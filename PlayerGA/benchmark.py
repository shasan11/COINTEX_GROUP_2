"""Reproducible batch benchmark for the CoinTex genetic agent.

The real game is graphical and event driven, which makes repeated controlled
trials slow and difficult to compare.  This module implements the same normalized
arena, movement rate, level timer, health loss, coin objective and danger types
in a deterministic simulator.  It is an evaluation harness, not a replacement
for final GUI play testing.

Run from the repository root:

    python PlayerGA/benchmark.py --output output/lab2
"""

from __future__ import annotations

import argparse
import csv
from dataclasses import asdict, dataclass, replace
import json
import math
from pathlib import Path
import random
import statistics
import sys
import time


PROJECT_DIR = Path(__file__).resolve().parents[1]
if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))

import levels
from PlayerGA.ga_core import (
    BASELINE_SETTINGS,
    IMPROVED_SETTINGS,
    EvolutionaryPlanner,
    baseline_fitness,
    clamp,
    distance,
    improved_fitness,
)


DT = 0.40
PLAYER_SPEED = 0.60
COLLECT_RADIUS = 0.060
DAMAGE_RADIUS = 0.105
FIRE_RANGE = 0.22


@dataclass
class TrialResult:
    configuration: str
    level: int
    world: int
    seed: int
    success: int
    elapsed_seconds: float
    remaining_health: float
    remaining_health_pct: float
    coins_collected: int
    total_coins: int
    generations: int


class Arena:
    """Small deterministic simulation of one CoinTex level."""

    def __init__(self, level_number: int, seed: int) -> None:
        self.level = levels.get_level(level_number)
        self.random = random.Random(seed * 1009 + level_number * 9176)
        self.player = [0.25, 0.22]
        self.health = float(self.level["player_health"])
        self.max_health = self.health
        self.time_left = float(self.level["time_limit"])
        self.elapsed = 0.0
        self.ammo = int(self.level["ammo"])
        self.reload_time = 0.0
        self.total_coins = int(self.level["coins"])
        section = (0.93 - 0.07) / self.total_coins
        self.coins = [
            (
                self.random.uniform(0.07 + section * i, 0.07 + section * (i + 1)),
                self.random.uniform(0.14, 0.88),
            )
            for i in range(self.total_coins)
        ]
        self.dangers: list[dict[str, float | bool | str]] = []
        for _ in range(int(self.level["monsters"])):
            angle = self.random.uniform(0.0, math.tau)
            self.dangers.append(
                {
                    "kind": "monster",
                    "x": self.random.uniform(0.45, 0.90),
                    "y": self.random.uniform(0.18, 0.88),
                    "r": 0.055,
                    "fx": math.cos(angle),
                    "fy": math.sin(angle),
                    "chasing": False,
                }
            )
        for index in range(int(self.level["fires"])):
            horizontal = index % 2 == 0
            self.dangers.append(
                {
                    "kind": "fire",
                    "x": self.random.uniform(0.07, 0.93),
                    "y": self.random.uniform(0.18, 0.86),
                    "r": 0.045,
                    "fx": 1.0 if horizontal else 0.0,
                    "fy": 0.0 if horizontal else 1.0,
                    "chasing": False,
                    "direction": 1.0,
                }
            )

    def snapshot(self) -> dict[str, object]:
        danger_records = [
            {
                "x": float(item["x"]),
                "y": float(item["y"]),
                "r": float(item["r"]),
                "fx": float(item["fx"]),
                "fy": float(item["fy"]),
                "chasing": bool(item["chasing"]),
            }
            for item in self.dangers
        ]
        nearest_monster = min(
            (
                distance(tuple(self.player), (float(item["x"]), float(item["y"])))
                for item in self.dangers
                if item["kind"] == "monster"
            ),
            default=None,
        )
        closest = min(
            ((float(item["x"]), float(item["y"])) for item in self.dangers),
            key=lambda point: distance(tuple(self.player), point),
            default=None,
        )
        flee_from = closest if closest and distance(tuple(self.player), closest) < 0.18 else None
        return {
            "coins": list(self.coins),
            "danger": [(float(item["x"]), float(item["y"])) for item in self.dangers],
            "dangers": danger_records,
            "player": tuple(self.player),
            "nearest_monster": nearest_monster,
            "flee_from": flee_from,
            "time_left": self.time_left,
            "time_limit": float(self.level["time_limit"]),
            "coins_left": len(self.coins),
            "ammo": self.ammo,
            "reloading": self.reload_time > 0.0,
        }

    def _move_player(self, target: tuple[float, float]) -> None:
        dx, dy = target[0] - self.player[0], target[1] - self.player[1]
        gap = math.hypot(dx, dy)
        step = PLAYER_SPEED * DT
        if gap <= step:
            self.player[:] = target
        elif gap > 0.0:
            self.player[0] += dx / gap * step
            self.player[1] += dy / gap * step
        self.player[0] = clamp(self.player[0], 0.07, 0.93)
        self.player[1] = clamp(self.player[1], 0.12, 0.90)

    def _move_dangers(self) -> None:
        for item in self.dangers:
            if item["kind"] == "monster":
                gap = distance(tuple(self.player), (float(item["x"]), float(item["y"])))
                chasing = float(self.level["chase_range"]) > 0.0 and gap <= float(self.level["chase_range"])
                item["chasing"] = chasing
                if chasing and gap > 0.0:
                    item["fx"] = (self.player[0] - float(item["x"])) / gap
                    item["fy"] = (self.player[1] - float(item["y"])) / gap
                speed = (0.11 if not chasing else 0.23) * float(self.level["enemy_speed_mult"])
            else:
                speed = 0.32 * float(self.level["enemy_speed_mult"]) / float(self.level["fire_speed"])
            item["x"] = float(item["x"]) + float(item["fx"]) * speed * DT
            item["y"] = float(item["y"]) + float(item["fy"]) * speed * DT
            if float(item["x"]) < 0.07 or float(item["x"]) > 0.93:
                item["fx"] = -float(item["fx"])
                item["x"] = clamp(float(item["x"]), 0.07, 0.93)
            if float(item["y"]) < 0.12 or float(item["y"]) > 0.90:
                item["fy"] = -float(item["fy"])
                item["y"] = clamp(float(item["y"]), 0.12, 0.90)

    def _shoot_if_needed(self) -> None:
        monsters = [item for item in self.dangers if item["kind"] == "monster"]
        if self.ammo <= 0 or not monsters:
            return
        nearest = min(monsters, key=lambda item: distance(tuple(self.player), (float(item["x"]), float(item["y"]))))
        if distance(tuple(self.player), (float(nearest["x"]), float(nearest["y"]))) <= FIRE_RANGE:
            self.ammo -= 1
            nearest["x"] = self.random.choice((0.08, 0.92))
            nearest["y"] = self.random.uniform(0.16, 0.88)
            nearest["chasing"] = False
            if self.ammo == 0 and self.level["gun_reload"]:
                self.reload_time = 5.0

    def advance(self, target: tuple[float, float]) -> None:
        self._move_player(target)
        self._move_dangers()
        self._shoot_if_needed()
        if self.reload_time > 0.0:
            self.reload_time = max(0.0, self.reload_time - DT)
            if self.reload_time == 0.0:
                self.ammo = int(self.level["ammo"])

        self.coins = [coin for coin in self.coins if distance(tuple(self.player), coin) > COLLECT_RADIUS]
        touching = any(
            distance(tuple(self.player), (float(item["x"]), float(item["y"])))
            < DAMAGE_RADIUS + float(item["r"]) * 0.35
            for item in self.dangers
        )
        if touching:
            self.health -= float(self.level["contact_damage"]) * DT
        self.time_left -= DT
        self.elapsed += DT

    @property
    def finished(self) -> bool:
        return not self.coins or self.health <= 0.0 or self.time_left <= 0.0


def run_trial_with_settings(
    configuration: str,
    level_number: int,
    seed: int,
    settings,
    fitness,
) -> TrialResult:
    """Run one trial with an explicit planner configuration."""

    arena = Arena(level_number, seed)
    planner = EvolutionaryPlanner(settings=settings, fitness_function=fitness, seed=seed * 7919 + level_number)
    while not arena.finished:
        target, _ = planner.step(arena.snapshot())
        arena.advance(target)
    collected = arena.total_coins - len(arena.coins)
    return TrialResult(
        configuration=configuration,
        level=level_number,
        world=int(arena.level["world"]),
        seed=seed,
        success=int(not arena.coins and arena.health > 0.0 and arena.time_left > 0.0),
        elapsed_seconds=round(arena.elapsed, 3),
        remaining_health=round(max(0.0, arena.health), 3),
        remaining_health_pct=round(100.0 * max(0.0, arena.health) / arena.max_health, 3),
        coins_collected=collected,
        total_coins=arena.total_coins,
        generations=planner.generation,
    )


def run_trial(configuration: str, level_number: int, seed: int) -> TrialResult:
    if configuration == "baseline":
        settings, fitness = BASELINE_SETTINGS, baseline_fitness
    elif configuration == "fitness_only":
        settings, fitness = BASELINE_SETTINGS, improved_fitness
    elif configuration == "operators_only":
        settings, fitness = IMPROVED_SETTINGS, baseline_fitness
    elif configuration == "improved":
        settings, fitness = IMPROVED_SETTINGS, improved_fitness
    else:
        raise ValueError(f"unknown configuration: {configuration}")

    return run_trial_with_settings(configuration, level_number, seed, settings, fitness)


def summarise(results: list[TrialResult]) -> list[dict[str, float | int | str]]:
    grouped: dict[tuple[str, int], list[TrialResult]] = {}
    for result in results:
        grouped.setdefault((result.configuration, result.level), []).append(result)
    rows: list[dict[str, float | int | str]] = []
    for (configuration, level_number), trials in sorted(grouped.items()):
        successes = [trial for trial in trials if trial.success]
        rows.append(
            {
                "configuration": configuration,
                "level": level_number,
                "runs": len(trials),
                "success_rate_pct": round(100.0 * sum(item.success for item in trials) / len(trials), 2),
                "mean_time_seconds": round(statistics.mean(item.elapsed_seconds for item in trials), 2),
                "mean_success_time_seconds": round(
                    statistics.mean(item.elapsed_seconds for item in successes), 2
                ) if successes else "",
                "mean_health_pct": round(statistics.mean(item.remaining_health_pct for item in trials), 2),
                "mean_coin_completion_pct": round(
                    statistics.mean(100.0 * item.coins_collected / item.total_coins for item in trials), 2
                ),
            }
        )
    return rows


def _write_csv(path: Path, rows: list[dict[str, object]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(rows[0]))
        writer.writeheader()
        writer.writerows(rows)


def make_chart(summary_rows: list[dict[str, object]], path: Path) -> bool:
    try:
        import matplotlib.pyplot as plt
    except ImportError:
        return False
    configurations = ["baseline", "fitness_only", "operators_only", "improved"]
    levels_used = sorted({int(row["level"]) for row in summary_rows})
    lookup = {(str(row["configuration"]), int(row["level"])): row for row in summary_rows}
    x = list(range(len(levels_used)))
    width = 0.19
    colors = ["#9ca3af", "#60a5fa", "#f59e0b", "#166534"]
    fig, axis = plt.subplots(figsize=(10.5, 5.5))
    for offset, (configuration, color) in enumerate(zip(configurations, colors)):
        values = [float(lookup[(configuration, level)]["success_rate_pct"]) for level in levels_used]
        positions = [value + (offset - 1.5) * width for value in x]
        axis.bar(positions, values, width=width, label=configuration.replace("_", " ").title(), color=color)
    axis.set_ylabel("Success rate (%)")
    axis.set_xlabel("CoinTex level")
    axis.set_xticks(x, [str(item) for item in levels_used])
    axis.set_ylim(0, 105)
    axis.grid(axis="y", alpha=0.25)
    axis.legend(ncol=2, frameon=False)
    fig.tight_layout()
    fig.savefig(path, dpi=180)
    plt.close(fig)
    return True


def run_benchmark(output_dir: Path, seeds: int, selected_levels: list[int]) -> dict[str, object]:
    started = time.perf_counter()
    configurations = ["baseline", "fitness_only", "operators_only", "improved"]
    results = [
        run_trial(configuration, level_number, seed)
        for configuration in configurations
        for level_number in selected_levels
        for seed in range(seeds)
    ]
    summary_rows = summarise(results)
    output_dir.mkdir(parents=True, exist_ok=True)
    _write_csv(output_dir / "benchmark_trials.csv", [asdict(result) for result in results])
    _write_csv(output_dir / "benchmark_summary.csv", summary_rows)
    chart_created = make_chart(summary_rows, output_dir / "success_rate_by_level.png")
    metadata = {
        "levels": selected_levels,
        "seeds_per_configuration_level": seeds,
        "total_trials": len(results),
        "dt_seconds": DT,
        "baseline_settings": asdict(BASELINE_SETTINGS),
        "improved_settings": asdict(IMPROVED_SETTINGS),
        "chart_created": chart_created,
        "elapsed_wall_seconds": round(time.perf_counter() - started, 3),
        "scope_note": (
            "Deterministic simulator using CoinTex level parameters. "
            "GUI play testing remains a separate final validation step."
        ),
    }
    (output_dir / "benchmark_metadata.json").write_text(json.dumps(metadata, indent=2), encoding="utf-8")
    return metadata


def run_parameter_tuning(output_dir: Path, seeds: int = 8) -> list[dict[str, object]]:
    """Evaluate a small population and mutation grid on middle/hard levels."""

    candidates = []
    all_trials: list[TrialResult] = []
    for population_size in (60, 120, 240):
        for mutation_probability in (0.15, 0.30, 0.50):
            settings = replace(
                IMPROVED_SETTINGS,
                name=f"pop{population_size}_mut{mutation_probability:.2f}",
                population_size=population_size,
                parent_count=max(12, population_size // 4),
                elite_count=max(3, population_size // 20),
                mutation_probability=mutation_probability,
            )
            trials = [
                run_trial_with_settings(settings.name, level_number, seed, settings, improved_fitness)
                for level_number in (20, 40)
                for seed in range(seeds)
            ]
            all_trials.extend(trials)
            successful = [trial for trial in trials if trial.success]
            candidates.append(
                {
                    "configuration": settings.name,
                    "population_size": population_size,
                    "mutation_probability": mutation_probability,
                    "trials": len(trials),
                    "success_rate_pct": round(100.0 * sum(item.success for item in trials) / len(trials), 2),
                    "mean_success_time_seconds": round(
                        statistics.mean(item.elapsed_seconds for item in successful), 2
                    ) if successful else "",
                    "mean_health_pct": round(statistics.mean(item.remaining_health_pct for item in trials), 2),
                    "mean_coin_completion_pct": round(
                        statistics.mean(100.0 * item.coins_collected / item.total_coins for item in trials), 2
                    ),
                }
            )
    _write_csv(output_dir / "tuning_trials.csv", [asdict(result) for result in all_trials])
    _write_csv(output_dir / "tuning_summary.csv", candidates)
    return candidates


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, default=PROJECT_DIR / "output" / "lab2")
    parser.add_argument("--seeds", type=int, default=12)
    parser.add_argument("--levels", type=int, nargs="+", default=[1, 20, 40, 60])
    parser.add_argument("--tune", action="store_true", help="also run the population/mutation grid")
    return parser.parse_args()


if __name__ == "__main__":
    arguments = parse_args()
    info = run_benchmark(arguments.output, arguments.seeds, arguments.levels)
    if arguments.tune:
        tuning = run_parameter_tuning(arguments.output)
        info["tuning_candidates"] = len(tuning)
    print(json.dumps(info, indent=2))
