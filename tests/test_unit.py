import unittest

from Fantasy import Player, calculated_trade_value, trade_verdict


class TradeValueUnitTests(unittest.TestCase):
	def test_calculated_trade_value_applies_rank_multipliers(self) -> None:
		top_five = Player("Top Five", "QB", 5, 1, custom_trade_value=20)
		top_ten = Player("Top Ten", "RB", 8, 1, custom_trade_value=20)
		other = Player("Other", "WR", 11, 1, custom_trade_value=20)

		self.assertEqual(calculated_trade_value(top_five), 40)
		self.assertEqual(calculated_trade_value(top_ten), 25)
		self.assertEqual(calculated_trade_value(other), 20)

	def test_trade_verdict_matches_difference_thresholds(self) -> None:
		expected = {
			-30: "They're robbing you",
			-15: "Bad trade for you",
			-5: "You're giving up too much",
			4.99: "Fair trade",
			5: "Good trade for you",
			15: "Great trade for you",
			30: "You're robbing them",
		}

		for difference, verdict in expected.items():
			with self.subTest(difference=difference):
				self.assertEqual(trade_verdict(difference), verdict)


if __name__ == "__main__":
	unittest.main()
