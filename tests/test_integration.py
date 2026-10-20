import json
import tempfile
import unittest
from io import BytesIO
from pathlib import Path
from unittest.mock import patch

from Fantasy import Player, fetch_api_players, read_excel, write_excel


class WorkbookIntegrationTests(unittest.TestCase):
	def test_player_data_round_trips_through_excel(self) -> None:
		players = (
			Player("Avery Example", "WR", 1, 1, "player-1", "SEA", 42),
			Player("Riley Example", "QB", 2, 1, "player-2", "BUF", 18),
		)
		with tempfile.TemporaryDirectory() as directory:
			workbook_path = Path(directory) / "players.xlsx"
			write_excel(players, str(workbook_path))

			loaded = read_excel(str(workbook_path))

		self.assertEqual(loaded, players)

	def test_api_players_export_and_reload_with_rankings(self) -> None:
		api_data = {
			"later": {
				"position": "RB",
				"search_rank": 20,
				"active": True,
				"first_name": "Riley",
				"last_name": "Runner",
				"player_id": "rb-1",
				"team": "BUF",
			},
			"first": {
				"position": "QB",
				"search_rank": 10,
				"active": True,
				"first_name": "Avery",
				"last_name": "Passer",
				"player_id": "qb-1",
				"team": "SEA",
			},
			"inactive": {
				"position": "WR",
				"search_rank": 1,
				"active": False,
				"first_name": "Inactive",
				"last_name": "Player",
				"player_id": "inactive-1",
			},
		}
		response = BytesIO(json.dumps(api_data).encode("utf-8"))
		with patch("Fantasy.urlopen", return_value=response):
			players = fetch_api_players()

		with tempfile.TemporaryDirectory() as directory:
			workbook_path = Path(directory) / "players.xlsx"
			write_excel(players, str(workbook_path))
			loaded = read_excel(str(workbook_path))

		self.assertEqual(
			[(player.name, player.global_rank, player.positional_rank) for player in loaded],
			[("Avery Passer", 1, 1), ("Riley Runner", 2, 1)],
		)
		self.assertEqual([player.player_id for player in loaded], ["qb-1", "rb-1"])


if __name__ == "__main__":
	unittest.main()
