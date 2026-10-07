import unittest

from Fantasy import Player, recalculate_positional_ranks, shift_player_rank


class ShiftPlayerRankTests(unittest.TestCase):
	def setUp(self) -> None:
		self.players = [
			Player("A", "QB", 1, 1),
			Player("B", "RB", 2, 1),
			Player("C", "QB", 3, 2),
			Player("D", "QB", 4, 3),
		]

	def test_global_rank_move_up_shifts_intervening_players_down(self) -> None:
		shift_player_rank(self.players, 3, 1)

		self.assertEqual(
			{player.name: player.global_rank for player in self.players},
			{"A": 2, "B": 3, "C": 4, "D": 1},
		)

	def test_global_rank_move_down_shifts_intervening_players_up(self) -> None:
		shift_player_rank(self.players, 0, 4)

		self.assertEqual(
			{player.name: player.global_rank for player in self.players},
			{"A": 4, "B": 1, "C": 2, "D": 3},
		)

	def test_positional_rank_move_only_shifts_same_position(self) -> None:
		shift_player_rank(self.players, 3, 1, positional=True)

		self.assertEqual(
			{player.name: player.positional_rank for player in self.players},
			{"A": 2, "B": 1, "C": 3, "D": 1},
		)
		self.assertEqual(
			{player.name: player.global_rank for player in self.players},
			{"A": 1, "B": 2, "C": 3, "D": 4},
		)

	def test_positional_ranks_follow_global_order_after_global_move(self) -> None:
		self.players[0] = Player("A", "QB", 1, 3)
		self.players[1] = Player("B", "RB", 2, 1)
		self.players[2] = Player("C", "QB", 3, 1)
		self.players[3] = Player("D", "QB", 4, 2)

		shift_player_rank(self.players, 3, 1)
		recalculate_positional_ranks(self.players)

		self.assertEqual(
			{player.name: player.positional_rank for player in self.players},
			{"A": 2, "B": 1, "C": 3, "D": 1},
		)


if __name__ == "__main__":
	unittest.main()
