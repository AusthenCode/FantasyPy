import unittest

from Fantasy import Player, evaluate_trade


class TradeValidationRegressionTests(unittest.TestCase):
	def test_trade_rejects_players_outside_top_150(self) -> None:
		players = (
			Player("Eligible", "RB", 10, 1, custom_trade_value=20),
			Player("Outside Range", "WR", 151, 100, custom_trade_value=20),
		)

		with self.assertRaisesRegex(ValueError, "top 150"):
			evaluate_trade(["Eligible"], ["Outside Range"], players)

	def test_trade_rejects_unset_or_invalid_trade_values(self) -> None:
		eligible = Player("Eligible", "RB", 10, 1, custom_trade_value=20)
		for invalid_value in (None, 0, 101):
			with self.subTest(invalid_value=invalid_value):
				player = Player("Invalid", "WR", 20, 1, custom_trade_value=invalid_value)
				with self.assertRaisesRegex(ValueError, "Set each player's trade value"):
					evaluate_trade(["Eligible"], ["Invalid"], (eligible, player))


if __name__ == "__main__":
	unittest.main()
