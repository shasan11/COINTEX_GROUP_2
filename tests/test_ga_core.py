import math
import unittest

from PlayerGA.ga_core import (
    BASELINE_SETTINGS,
    IMPROVED_SETTINGS,
    EvolutionaryPlanner,
    baseline_fitness,
    improved_fitness,
    point_segment_distance,
)


def snapshot(**overrides):
    value = {
        "coins": [(0.82, 0.75), (0.22, 0.78)],
        "player": (0.20, 0.20),
        "danger": [],
        "dangers": [],
        "flee_from": None,
        "time_left": 80.0,
        "time_limit": 100.0,
        "ammo": 1,
    }
    value.update(overrides)
    return value


class GACoreTests(unittest.TestCase):
    def test_point_segment_distance_projects_inside_segment(self):
        self.assertTrue(math.isclose(
            point_segment_distance((0.5, 0.4), (0.0, 0.0), (1.0, 0.0)), 0.4
        ))

    def test_baseline_fitness_only_targets_first_coin(self):
        state = snapshot()
        self.assertGreater(
            baseline_fitness(state, 0.82, 0.75),
            baseline_fitness(state, 0.22, 0.78),
        )

    def test_improved_fitness_can_choose_nearer_reachable_coin(self):
        state = snapshot()
        self.assertGreater(
            improved_fitness(state, 0.22, 0.78),
            improved_fitness(state, 0.82, 0.75),
        )

    def test_improved_fitness_penalises_danger_on_path(self):
        danger = {"x": 0.50, "y": 0.50, "r": 0.06, "fx": 0.0, "fy": 0.0, "chasing": False}
        state = snapshot(coins=[(0.82, 0.80)], dangers=[danger], danger=[(0.50, 0.50)])
        direct = improved_fitness(state, 0.82, 0.80)
        detour = improved_fitness(state, 0.30, 0.80)
        self.assertGreater(detour, direct)

    def test_planner_is_reproducible_for_fixed_seed(self):
        state = snapshot()
        first = EvolutionaryPlanner(IMPROVED_SETTINGS, improved_fitness, seed=44)
        second = EvolutionaryPlanner(IMPROVED_SETTINGS, improved_fitness, seed=44)
        self.assertEqual(
            [first.step(state) for _ in range(5)],
            [second.step(state) for _ in range(5)],
        )

    def test_adaptive_mutation_increases_after_stagnation(self):
        planner = EvolutionaryPlanner(IMPROVED_SETTINGS, improved_fitness, seed=7)
        planner.stagnation = IMPROVED_SETTINGS.stagnation_generations + 2
        self.assertGreater(
            planner._mutation_probability(), IMPROVED_SETTINGS.mutation_probability
        )

    def test_settings_have_equal_population_for_fair_comparison(self):
        self.assertEqual(
            BASELINE_SETTINGS.population_size, IMPROVED_SETTINGS.population_size
        )


if __name__ == "__main__":
    unittest.main()
