import contextlib
import io
import sys
import unittest
from unittest.mock import patch

from Fantasy import Player, main


class CommandLineAcceptanceTests(unittest.TestCase):
	def setUp(self) -> None:
		self.players = (
			Player("Quarterback", "QB", 1, 1, custom_trade_value=20),
			Player("Wide Receiver B", "WR", 4, 2, custom_trade_value=12),
			Player("Wide Receiver A", "WR", 3, 1, custom_trade_value=10),
		)

	def test_rankings_command_filters_sorts_and_limits_results(self) -> None:
		output = io.StringIO()
		with patch.object(sys, "argv", ["Fantasy.py", "rankings", "--position", "WR", "--sort", "position", "--limit", "1"]), \
				patch("Fantasy.load_players", return_value=self.players), \
				contextlib.redirect_stdout(output):
			main()

		self.assertIn("Wide Receiver A", output.getvalue())
		self.assertNotIn("Wide Receiver B", output.getvalue())
		self.assertNotIn("Quarterback", output.getvalue())

	def test_trade_command_displays_totals_and_verdict(self) -> None:
		output = io.StringIO()
		with patch.object(
			sys,
			"argv",
			["Fantasy.py", "trade", "--give", "Wide Receiver A", "--get", "Wide Receiver B"],
		), patch("Fantasy.load_players", return_value=self.players), contextlib.redirect_stdout(output):
			main()

		self.assertIn("10.0 given | 12.0 received", output.getvalue())
		self.assertIn("Difference: +2.0 (Fair trade)", output.getvalue())


if __name__ == "__main__":
	unittest.main()
