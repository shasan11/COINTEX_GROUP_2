"""Pure genetic-algorithm components used by the CoinTex research agent.

The GUI adapter in :mod:`PlayerGA.main` supplies live game snapshots.  This
module deliberately has no Kivy or PyGAD dependency, which makes the fitness
functions and evolutionary operators reproducible in tests and batch
experiments.

A chromosome contains two real-valued genes, ``[x, y]``.  They describe the
next target point in normalized game coordinates.  Both fitness functions are
maximized.
"""

from __future__ import annotations

from dataclasses import dataclass
import math
import random
from typing import Callable, Iterable, Mapping, Sequence


Point = tuple[float, float]
Snapshot = Mapping[str, object]


@dataclass(frozen=True)
class GASettings:
    """Settings for one evolutionary planner configuration."""

    name: str
    population_size: int
    parent_count: int
    elite_count: int
    mutation_probability: float
    crossover: str = "uniform"
    tournament_size: int = 3
    adaptive_mutation: bool = False
    stagnation_generations: int = 6
    max_mutation_probability: float = 0.60

    def validate(self) -> None:
        if self.population_size < 4:
            raise ValueError("population_size must be at least 4")
        if not 2 <= self.parent_count <= self.population_size:
            raise ValueError("parent_count must be between 2 and population_size")
        if not 0 <= self.elite_count <= self.parent_count:
            raise ValueError("elite_count must be between 0 and parent_count")
        if not 0.0 <= self.mutation_probability <= 1.0:
            raise ValueError("mutation_probability must be in [0, 1]")
        if self.crossover not in {"uniform", "blend"}:
            raise ValueError("crossover must be 'uniform' or 'blend'")


BASELINE_SETTINGS = GASettings(
    name="baseline",
    population_size=60,
    parent_count=18,
    elite_count=0,
    mutation_probability=0.10,
    crossover="uniform",
)

IMPROVED_SETTINGS = GASettings(
    name="improved",
    population_size=60,
    parent_count=15,
    elite_count=3,
    mutation_probability=0.30,
    crossover="blend",
    tournament_size=4,
    adaptive_mutation=True,
    stagnation_generations=5,
    max_mutation_probability=0.65,
)


def clamp(value: float, low: float = 0.0, high: float = 1.0) -> float:
    return max(low, min(high, float(value)))


def distance(a: Point, b: Point) -> float:
    return math.hypot(a[0] - b[0], a[1] - b[1])


def point_segment_distance(point: Point, start: Point, end: Point) -> float:
    """Return the shortest distance between *point* and a line segment."""

    px, py = point
    ax, ay = start
    bx, by = end
    abx, aby = bx - ax, by - ay
    length_sq = abx * abx + aby * aby
    if length_sq == 0.0:
        t = 0.0
    else:
        t = ((px - ax) * abx + (py - ay) * aby) / length_sq
        t = clamp(t)
    return math.hypot(px - (ax + t * abx), py - (ay + t * aby))


def _danger_records(snapshot: Snapshot) -> list[dict[str, float | bool]]:
    records = snapshot.get("dangers")
    if records is not None:
        return [dict(item) for item in records]  # defensive copy for callers
    return [
        {"x": float(x), "y": float(y), "r": 0.05,
         "fx": 0.0, "fy": 0.0, "chasing": False}
        for x, y in snapshot.get("danger", [])
    ]


def baseline_fitness(snapshot: Snapshot | None, target_x: float, target_y: float) -> float:
    """Reproduce the tutorial's original first-coin, square-penalty fitness.

    The original code adds a constant reward for every danger outside a 0.3 by
    0.3 box around the candidate and subtracts 300 inside that box.  The zero
    distance guard below is the only numerical correction.
    """

    if not snapshot or not snapshot.get("coins"):
        return 0.0
    target = (float(target_x), float(target_y))
    coin = tuple(snapshot["coins"][0])
    score = 1.0 / (abs(target[0] - coin[0]) + abs(target[1] - coin[1]) + 1e-9)
    for danger in _danger_records(snapshot):
        if abs(target[0] - float(danger["x"])) <= 0.3 and abs(target[1] - float(danger["y"])) <= 0.3:
            score -= 300.0
        else:
            score += 100.0
    return score


def improved_fitness(snapshot: Snapshot | None, target_x: float, target_y: float) -> float:
    """Score a target using reachable coins, path safety, urgency and escape.

    Improvements over the baseline are designed around the game mechanics:

    * every remaining coin competes as a goal instead of list order deciding;
    * hazards are measured against the whole path from player to target;
    * chasing monsters are projected along their heading;
    * low time increases progress pressure without removing safety;
    * a nearby danger rewards movement away from it;
    * targets outside the playable area receive a strong penalty.
    """

    if not snapshot or not snapshot.get("coins"):
        return 0.0

    target = (float(target_x), float(target_y))
    player = tuple(snapshot["player"])
    if not (0.07 <= target[0] <= 0.93 and 0.12 <= target[1] <= 0.90):
        return -1000.0 - 500.0 * distance(target, (clamp(target[0]), clamp(target[1])))

    time_left = snapshot.get("time_left")
    time_limit = snapshot.get("time_limit")
    urgency = 0.0
    if isinstance(time_left, (int, float)) and isinstance(time_limit, (int, float)) and time_limit > 0:
        urgency = clamp(1.0 - float(time_left) / (0.60 * float(time_limit)))

    # Prefer a coin that is close to both the candidate and the player.  The
    # player-distance term prevents a far coin from winning only because the
    # candidate happened to land directly on it.
    progress_score = max(
        10.0 / (distance(target, tuple(coin)) + 0.025)
        - 7.0 * distance(player, tuple(coin))
        for coin in snapshot["coins"]
    )
    score = progress_score * (1.0 + 1.8 * urgency)

    dangers = _danger_records(snapshot)
    no_ammo = int(snapshot.get("ammo", 1)) <= 0
    closest_player_gap = math.inf
    for danger in dangers:
        chasing = bool(danger.get("chasing", False))
        lead = 0.14 if chasing else 0.0
        projected = (
            float(danger["x"]) + float(danger.get("fx", 0.0)) * lead,
            float(danger["y"]) + float(danger.get("fy", 0.0)) * lead,
        )
        radius = float(danger.get("r", 0.05))
        clearance = 0.12 + radius + (0.06 if chasing else 0.0) + (0.08 if chasing and no_ammo else 0.0)
        path_gap = point_segment_distance(projected, player, target)
        player_gap = distance(player, projected)
        closest_player_gap = min(closest_player_gap, player_gap)
        if path_gap < clearance:
            ratio = 1.0 - path_gap / max(clearance, 1e-9)
            score -= (430.0 - 120.0 * urgency) * ratio * ratio
        else:
            score += min(55.0, 80.0 * (path_gap - clearance))

    if dangers and closest_player_gap < 0.20:
        nearest = min(
            dangers,
            key=lambda item: distance(player, (float(item["x"]), float(item["y"]))),
        )
        away_gain = distance(target, (float(nearest["x"]), float(nearest["y"]))) - closest_player_gap
        score += 150.0 * away_gain

    # Short targets let the controller react again sooner and reduce overshoot.
    score -= 9.0 * distance(player, target)
    return score


FitnessFunction = Callable[[Snapshot | None, float, float], float]


class EvolutionaryPlanner:
    """A small real-valued GA that returns one target per generation."""

    def __init__(
        self,
        settings: GASettings = IMPROVED_SETTINGS,
        fitness_function: FitnessFunction = improved_fitness,
        seed: int | None = None,
    ) -> None:
        settings.validate()
        self.settings = settings
        self.fitness_function = fitness_function
        self.random = random.Random(seed)
        self.population: list[Point] = [self._random_point() for _ in range(settings.population_size)]
        self.best_fitness = -math.inf
        self.stagnation = 0
        self.generation = 0

    def _random_point(self) -> Point:
        return (self.random.uniform(0.07, 0.93), self.random.uniform(0.12, 0.90))

    def _select_parent(self, parents: Sequence[tuple[float, Point]]) -> Point:
        size = min(self.settings.tournament_size, len(parents))
        contestants = self.random.sample(list(parents), size)
        return max(contestants, key=lambda item: item[0])[1]

    def _crossover(self, mother: Point, father: Point) -> Point:
        if self.settings.crossover == "blend":
            alpha = self.random.random()
            return (
                clamp(alpha * mother[0] + (1.0 - alpha) * father[0], 0.07, 0.93),
                clamp(alpha * mother[1] + (1.0 - alpha) * father[1], 0.12, 0.90),
            )
        return (
            mother[0] if self.random.random() < 0.5 else father[0],
            mother[1] if self.random.random() < 0.5 else father[1],
        )

    def _mutation_probability(self) -> float:
        base = self.settings.mutation_probability
        if not self.settings.adaptive_mutation or self.stagnation < self.settings.stagnation_generations:
            return base
        extra_steps = self.stagnation - self.settings.stagnation_generations + 1
        return min(self.settings.max_mutation_probability, base + 0.07 * extra_steps)

    def _mutate(self, child: Point) -> Point:
        probability = self._mutation_probability()
        values = list(child)
        for index, bounds in enumerate(((0.07, 0.93), (0.12, 0.90))):
            if self.random.random() < probability:
                if self.settings.crossover == "blend" and self.random.random() < 0.70:
                    values[index] = clamp(values[index] + self.random.gauss(0.0, 0.13), *bounds)
                else:
                    values[index] = self.random.uniform(*bounds)
        return (values[0], values[1])

    def step(self, snapshot: Snapshot) -> tuple[Point, float]:
        ranked = sorted(
            ((self.fitness_function(snapshot, x, y), (x, y)) for x, y in self.population),
            key=lambda item: item[0],
            reverse=True,
        )
        generation_best, best_point = ranked[0]
        if generation_best > self.best_fitness + 1e-9:
            self.best_fitness = generation_best
            self.stagnation = 0
        else:
            self.stagnation += 1

        parents = ranked[: self.settings.parent_count]
        next_population = [point for _, point in ranked[: self.settings.elite_count]]
        while len(next_population) < self.settings.population_size:
            mother = self._select_parent(parents)
            father = self._select_parent(parents)
            next_population.append(self._mutate(self._crossover(mother, father)))
        self.population = next_population
        self.generation += 1
        return best_point, generation_best


def ranked_points(
    snapshot: Snapshot,
    points: Iterable[Point],
    fitness_function: FitnessFunction = improved_fitness,
) -> list[tuple[float, Point]]:
    """Convenience helper used by tests and small parameter studies."""

    return sorted(
        ((fitness_function(snapshot, x, y), (x, y)) for x, y in points),
        key=lambda item: item[0],
        reverse=True,
    )
