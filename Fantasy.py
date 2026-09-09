"""Fantasy football trade calculator for QB, RB, WR, and TE players."""

from __future__ import annotations

import argparse
import json
import os
from pathlib import Path
import tkinter as tk
from urllib.request import Request, urlopen
from dataclasses import dataclass
from typing import Iterable
from tkinter import messagebox, ttk

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill


POSITIONS = ("QB", "RB", "WR", "TE")
SLEEPER_PLAYERS_URL = "https://api.sleeper.app/v1/players/nfl"
DEFAULT_EXCEL_FILE = "players.xlsx"


@dataclass(frozen=True)
class Player:
	name: str
	position: str
	global_rank: int
	positional_rank: int
	player_id: str = ""
	team: str = ""


def trade_value(player: Player) -> float:
	"""Convert global rank into a simple, consistent trade value."""
	return round(1000 / (player.global_rank + 10), 1)


def find_player(name: str, players: tuple[Player, ...]) -> Player:
	player_lookup = {player.name.casefold(): player for player in players}
	player = player_lookup.get(name.strip().casefold())
	if player is None:
		suggestions = [
			candidate.name
			for candidate in players
			if name.casefold() in candidate.name.casefold()
		]
		hint = f" Did you mean: {', '.join(suggestions)}?" if suggestions else ""
		raise ValueError(f"Player not found: {name}.{hint}")
	return player


def print_rankings(players: tuple[Player, ...], position: str | None, limit: int, sort_by: str) -> None:
	players = [player for player in players if not position or player.position == position]
	if sort_by == "position":
		players.sort(key=lambda player: (player.position, player.positional_rank))
	else:
		players.sort(key=lambda player: player.global_rank)

	print(f"{'Global':<8}{'Pos':<6}{'Pos Rank':<10}{'Player':<24}{'Value':>7}")
	print("-" * 55)
	for player in players[:limit]:
		print(
			f"{player.global_rank:<8}{player.position:<6}"
			f"{player.positional_rank:<10}{player.name:<24}"
			f"{trade_value(player):>7.1f}"
		)


def parse_player_names(names: Iterable[str], players: tuple[Player, ...]) -> list[Player]:
	return [find_player(name, players) for name in names]


def evaluate_trade(give_names: list[str], get_names: list[str], players: tuple[Player, ...]) -> None:
	giving = parse_player_names(give_names, players)
	receiving = parse_player_names(get_names, players)
	give_total = sum(trade_value(player) for player in giving)
	get_total = sum(trade_value(player) for player in receiving)
	difference = get_total - give_total

	print("You give:")
	for player in giving:
		print(f"  {player.name} ({player.position}{player.positional_rank}) = {trade_value(player):.1f}")
	print("You get:")
	for player in receiving:
		print(f"  {player.name} ({player.position}{player.positional_rank}) = {trade_value(player):.1f}")
	print(f"\nTotals: {give_total:.1f} given | {get_total:.1f} received")
	if difference >= 5:
		verdict = "Strong value for you"
	elif difference <= -5:
		verdict = "You are giving up too much"
	else:
		verdict = "Fair trade"
	print(f"Difference: {difference:+.1f} ({verdict})")


def fetch_api_players() -> tuple[Player, ...]:
	request = Request(SLEEPER_PLAYERS_URL, headers={"User-Agent": "FantasyPy/1.0"})
	with urlopen(request, timeout=30) as response:
		data = json.load(response)

	api_players = []
	for player in data.values():
		position = player.get("position")
		search_rank = player.get("search_rank")
		if position not in POSITIONS or not player.get("active") or search_rank is None:
			continue
		name = " ".join(filter(None, (player.get("first_name"), player.get("last_name"))))
		if name:
			api_players.append((name, position, int(search_rank), player["player_id"], player.get("team") or "FA"))

	api_players.sort(key=lambda player: player[2])
	position_counts = {position: 0 for position in POSITIONS}
	players = []
	for global_rank, (name, position, _, player_id, team) in enumerate(api_players, start=1):
		position_counts[position] += 1
		players.append(Player(name, position, global_rank, position_counts[position], player_id, team))
	return tuple(players)


def write_excel(players: tuple[Player, ...], output_path: str) -> None:
	workbook = Workbook()
	worksheet = workbook.active
	worksheet.title = "Players"
	worksheet.append(["Name", "Position", "Team", "Global Rank", "Positional Rank", "Trade Value", "Player ID"])
	for player in players:
		worksheet.append([
			player.name,
			player.position,
			player.team,
			player.global_rank,
			player.positional_rank,
			trade_value(player),
			player.player_id,
		])
	for cell in worksheet[1]:
		cell.font = Font(bold=True, color="FFFFFF")
		cell.fill = PatternFill("solid", fgColor="1F4E78")
	worksheet.freeze_panes = "A2"
	worksheet.auto_filter.ref = worksheet.dimensions
	for column, width in {"A": 24, "B": 12, "C": 10, "D": 14, "E": 18, "F": 14, "G": 14}.items():
		worksheet.column_dimensions[column].width = width
	workbook.save(output_path)


def read_excel(input_path: str) -> tuple[Player, ...]:
	workbook = load_workbook(input_path, read_only=True, data_only=True)
	worksheet = workbook["Players"]
	players = []
	for row in worksheet.iter_rows(min_row=2, values_only=True):
		if not row[0]:
			continue
		players.append(Player(str(row[0]), str(row[1]), int(row[3]), int(row[4]), str(row[6] or ""), str(row[2] or "")))
	return tuple(players)


def load_players(input_path: str | None) -> tuple[Player, ...]:
	return read_excel(input_path or DEFAULT_EXCEL_FILE)


class RankingsWindow:
	def __init__(self, root: tk.Tk, excel_path: str) -> None:
		self.root = root
		self.excel_path = excel_path
		self.players: tuple[Player, ...] = ()
		self.position = tk.StringVar(value="ALL")
		self.sort_by = tk.StringVar(value="global")
		self.search = tk.StringVar()
		root.title("FantasyPy Rankings")
		root.geometry("900x620")
		root.minsize(700, 420)
		self.build_layout()
		self.load_rankings()

	def build_layout(self) -> None:
		style = ttk.Style()
		style.configure("Treeview", rowheight=28, font=("Segoe UI", 10))
		style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"))

		header = ttk.Frame(self.root, padding=(18, 16, 18, 8))
		header.pack(fill="x")
		ttk.Label(header, text="FantasyPy Rankings", font=("Segoe UI", 20, "bold")).pack(anchor="w")
		ttk.Label(header, text=f"Source: {Path(self.excel_path).name}").pack(anchor="w", pady=(2, 0))

		controls = ttk.Frame(self.root, padding=(18, 8))
		controls.pack(fill="x")
		ttk.Label(controls, text="Position").pack(side="left")
		position_box = ttk.Combobox(controls, textvariable=self.position, values=("ALL", *POSITIONS), state="readonly", width=8)
		position_box.pack(side="left", padx=(6, 18))
		position_box.bind("<<ComboboxSelected>>", lambda _event: self.refresh_table())
		ttk.Label(controls, text="Sort by").pack(side="left")
		sort_box = ttk.Combobox(controls, textvariable=self.sort_by, values=("global", "position"), state="readonly", width=10)
		sort_box.pack(side="left", padx=(6, 18))
		sort_box.bind("<<ComboboxSelected>>", lambda _event: self.refresh_table())
		ttk.Label(controls, text="Search").pack(side="left")
		search_box = ttk.Entry(controls, textvariable=self.search, width=24)
		search_box.pack(side="left", padx=(6, 8))
		search_box.bind("<KeyRelease>", lambda _event: self.refresh_table())
		ttk.Button(controls, text="Refresh", command=self.load_rankings).pack(side="right", padx=(8, 0))
		tk.Button(controls, text="Open Excel", command=self.open_excel).pack(side="right")

		table_frame = ttk.Frame(self.root, padding=(18, 0, 18, 18))
		table_frame.pack(fill="both", expand=True)
		columns = ("global", "position", "positional", "name", "team", "value")
		self.table = ttk.Treeview(table_frame, columns=columns, show="headings")
		for column, heading, width in (
			("global", "Global", 80),
			("position", "Position", 90),
			("positional", "Pos. Rank", 100),
			("name", "Player", 280),
			("team", "Team", 90),
			("value", "Trade Value", 110),
		):
			self.table.heading(column, text=heading)
			self.table.column(column, width=width, anchor="center" if column != "name" else "w")
		scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.table.yview)
		self.table.configure(yscrollcommand=scrollbar.set)
		self.table.pack(side="left", fill="both", expand=True)
		scrollbar.pack(side="right", fill="y")
		self.status = ttk.Label(self.root, padding=(18, 0, 18, 10))
		self.status.pack(fill="x")

	def load_rankings(self) -> None:
		try:
			self.players = load_players(self.excel_path)
			self.refresh_table()
		except (OSError, KeyError, TypeError, ValueError) as error:
			messagebox.showerror("Could not load rankings", str(error))

	def refresh_table(self) -> None:
		players = [player for player in self.players if self.position.get() == "ALL" or player.position == self.position.get()]
		query = self.search.get().strip().casefold()
		if query:
			players = [player for player in players if query in player.name.casefold() or query in player.team.casefold()]
		if self.sort_by.get() == "position":
			players.sort(key=lambda player: (player.position, player.positional_rank))
		else:
			players.sort(key=lambda player: player.global_rank)
		for item in self.table.get_children():
			self.table.delete(item)
		for player in players:
			self.table.insert("", "end", values=(player.global_rank, player.position, player.positional_rank, player.name, player.team, f"{trade_value(player):.1f}"))
		self.status.configure(text=f"Showing {len(players):,} players | Edit {self.excel_path} and click Refresh to update")

	def open_excel(self) -> None:
		try:
			os.startfile(Path(self.excel_path).resolve())
		except OSError as error:
			messagebox.showerror("Could not open Excel", str(error))


def launch_gui(excel_path: str = DEFAULT_EXCEL_FILE) -> None:
	root = tk.Tk()
	RankingsWindow(root, excel_path)
	root.mainloop()


def build_parser() -> argparse.ArgumentParser:
	parser = argparse.ArgumentParser(description="Rank and compare fantasy football players.")
	subparsers = parser.add_subparsers(dest="command")

	rankings = subparsers.add_parser("rankings", help="Show global or positional rankings.")
	rankings.add_argument("--position", choices=POSITIONS, help="Only show one position.")
	rankings.add_argument("--sort", choices=("global", "position"), default="global")
	rankings.add_argument("--limit", type=int, default=25)
	rankings.add_argument("--excel", default=DEFAULT_EXCEL_FILE, help="Read rankings from an Excel workbook.")

	trade = subparsers.add_parser("trade", help="Compare the value of two sides of a trade.")
	trade.add_argument("--give", nargs="+", required=True, help="Player names you give away.")
	trade.add_argument("--get", nargs="+", required=True, help="Player names you receive.")
	trade.add_argument("--excel", default=DEFAULT_EXCEL_FILE, help="Read players from an Excel workbook.")

	fetch = subparsers.add_parser("fetch", help="Fetch current NFL players from Sleeper and export them to Excel.")
	fetch.add_argument("--output", default=DEFAULT_EXCEL_FILE, help="Excel file to create.")

	gui = subparsers.add_parser("gui", help="Open the rankings window.")
	gui.add_argument("--excel", default=DEFAULT_EXCEL_FILE, help="Read rankings from an Excel workbook.")
	return parser


def main() -> None:
	args = build_parser().parse_args()
	try:
		if args.command is None:
			launch_gui()
		elif args.command == "gui":
			launch_gui(args.excel)
		elif args.command == "rankings":
			if args.limit < 1:
				raise ValueError("--limit must be at least 1")
			print_rankings(load_players(args.excel), args.position, args.limit, args.sort)
		elif args.command == "trade":
			evaluate_trade(args.give, args.get, load_players(args.excel))
		else:
			players = fetch_api_players()
			write_excel(players, args.output)
			print(f"Saved {len(players)} players to {Path(args.output).resolve()}")
	except OSError as error:
		raise SystemExit(f"Error accessing data: {error}") from error
	except (KeyError, TypeError, ValueError) as error:
		raise SystemExit(f"Error reading player data: {error}") from error


if __name__ == "__main__":
	main()
