"""Fantasy football trade calculator for QB, RB, WR, and TE players."""

from __future__ import annotations

import argparse
import json
import os
from pathlib import Path
import tkinter as tk
from urllib.request import Request, urlopen
from dataclasses import dataclass, replace
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
	workbook.close()
	return tuple(players)


def save_rankings(input_path: str, players: tuple[Player, ...]) -> None:
	workbook = load_workbook(input_path)
	worksheet = workbook["Players"]
	columns = {cell.value: cell.column for cell in worksheet[1]}
	global_column = columns["Global Rank"]
	positional_column = columns["Positional Rank"]
	value_column = columns["Trade Value"]
	id_column = columns["Player ID"]
	name_column = columns["Name"]
	position_column = columns["Position"]
	team_column = columns["Team"]
	player_lookup = {
		("id", player.player_id) if player.player_id else ("name", player.name.casefold(), player.position, player.team.casefold()): player
		for player in players
	}
	updated = 0
	for row in range(2, worksheet.max_row + 1):
		player_id = worksheet.cell(row, id_column).value
		if player_id is not None:
			key = ("id", str(player_id))
		else:
			name = worksheet.cell(row, name_column).value
			position = worksheet.cell(row, position_column).value
			team = worksheet.cell(row, team_column).value or ""
			key = ("name", str(name).casefold(), str(position), str(team).casefold())
		player = player_lookup.get(key)
		if player is None:
			continue
		worksheet.cell(row, global_column, player.global_rank)
		worksheet.cell(row, positional_column, player.positional_rank)
		worksheet.cell(row, value_column, trade_value(player))
		updated += 1
	if updated != len(player_lookup):
		workbook.close()
		raise ValueError("Could not match every ranked player to a row in the workbook.")
	workbook.save(input_path)
	workbook.close()


def load_players(input_path: str | None) -> tuple[Player, ...]:
	return read_excel(input_path or DEFAULT_EXCEL_FILE)


class RankingsWindow:
	def __init__(self, root: tk.Tk, excel_path: str) -> None:
		self.root = root
		self.excel_path = excel_path
		self.players: tuple[Player, ...] = ()
		self.dirty = False
		self.selected_player_index: int | None = None
		self.position = tk.StringVar(value="ALL")
		self.sort_by = tk.StringVar(value="global")
		self.search = tk.StringVar()
		self.edit_global_rank = tk.StringVar()
		self.edit_positional_rank = tk.StringVar()
		self.selected_player_name = tk.StringVar(value="Select a player from the list")
		root.title("FantasyPy Rankings")
		root.geometry("1000x720")
		root.minsize(760, 660)
		self.build_layout()
		self.load_rankings()

	def build_layout(self) -> None:
		style = ttk.Style()
		style.theme_use("clam")
		background = "#151515"
		panel = "#202020"
		orange = "#F47721"
		white = "#F4F4F4"
		muted = "#B8B8B8"
		self.root.configure(background=background)
		style.configure("TFrame", background=background)
		style.configure("TLabel", background=background, foreground=white, font=("Segoe UI", 10))
		style.configure("TButton", font=("Segoe UI", 10, "bold"), padding=(12, 7), background=panel, foreground=white)
		style.map("TButton", background=[("active", "#383838")], foreground=[("active", white)])
		style.configure("Accent.TButton", background=orange, foreground="#111111")
		style.map("Accent.TButton", background=[("active", "#FF964C")], foreground=[("active", "#111111")])
		style.configure("TEntry", fieldbackground=panel, foreground=white, insertcolor=white)
		style.configure("TCombobox", fieldbackground=panel, background=panel, foreground=white, arrowcolor=orange)
		style.map("TCombobox", fieldbackground=[("readonly", panel)], foreground=[("readonly", white)])
		style.configure("Treeview", rowheight=30, font=("Segoe UI", 10), background=panel, fieldbackground=panel, foreground=white)
		style.map("Treeview", background=[("selected", orange)], foreground=[("selected", "#111111")])
		style.configure("Compact.Treeview", rowheight=18, font=("Segoe UI", 9), background=panel, fieldbackground=panel, foreground=white)
		style.map("Compact.Treeview", background=[("selected", orange)], foreground=[("selected", "#111111")])
		style.configure("Compact.Treeview.Heading", font=("Segoe UI", 8, "bold"), background="#101010", foreground=orange, padding=(5, 3))
		style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"), background="#101010", foreground=orange, padding=(8, 9))
		style.map("Treeview.Heading", background=[("active", "#303030")])
		style.configure("TNotebook", background=background, borderwidth=0)
		style.configure("TNotebook.Tab", background=panel, foreground=white, padding=(16, 9), font=("Segoe UI", 10, "bold"))
		style.map("TNotebook.Tab", background=[("selected", orange), ("active", "#383838")], foreground=[("selected", "#111111"), ("active", white)])

		header = ttk.Frame(self.root, padding=(18, 16, 18, 8))
		header.pack(fill="x")
		tk.Label(header, text="FantasyPy Rankings", background=background, foreground=white, font=("Segoe UI", 21, "bold")).pack(anchor="w")
		tk.Label(header, text=f"SOURCE  /  {Path(self.excel_path).name}", background=background, foreground=muted).pack(anchor="w", pady=(3, 0))

		self.notebook = ttk.Notebook(self.root)
		self.notebook.pack(fill="both", expand=True, padx=18, pady=(10, 0))
		self.global_tab = ttk.Frame(self.notebook)
		self.positional_tab = ttk.Frame(self.notebook)
		self.edit_tab = ttk.Frame(self.notebook)
		self.notebook.add(self.global_tab, text="Global Rankings")
		self.notebook.add(self.positional_tab, text="Positional Rankings")
		self.notebook.add(self.edit_tab, text="Edit Rankings")

		global_controls = ttk.Frame(self.global_tab, padding=(0, 8))
		global_controls.pack(fill="x")
		tk.Label(global_controls, text="Position").pack(side="left")
		position_box = ttk.Combobox(global_controls, textvariable=self.position, values=("ALL", *POSITIONS), state="readonly", width=8)
		position_box.pack(side="left", padx=(6, 18))
		position_box.bind("<<ComboboxSelected>>", lambda _event: self.refresh_table())
		tk.Label(global_controls, text="Sort by").pack(side="left")
		sort_box = ttk.Combobox(global_controls, textvariable=self.sort_by, values=("global", "position"), state="readonly", width=10)
		sort_box.pack(side="left", padx=(6, 18))
		sort_box.bind("<<ComboboxSelected>>", lambda _event: self.refresh_table())
		tk.Label(global_controls, text="Search").pack(side="left")
		search_box = ttk.Entry(global_controls, textvariable=self.search, width=24)
		search_box.pack(side="left", padx=(6, 8))
		search_box.bind("<KeyRelease>", lambda _event: self.refresh_table())
		tk.Button(global_controls, text="Refresh", command=self.load_rankings).pack(side="right", padx=(8, 0))
		tk.Button(global_controls, text="Open Excel", command=self.open_excel).pack(side="right")

		self.table = self.create_rankings_table(self.global_tab, include_positional_rank=False)
		self.positional_tables: dict[str, ttk.Treeview] = {}
		self.build_positional_board()

		edit_content = ttk.Frame(self.edit_tab, padding=(0, 8, 0, 12))
		edit_content.pack(fill="both", expand=True)
		ttk.Label(edit_content, text="Select a player to edit their global and positional ranks.", foreground=muted).pack(anchor="w", pady=(0, 8))
		self.editor_table = self.create_rankings_table(edit_content)
		self.editor_table.bind("<<TreeviewSelect>>", self.select_player)

		player_form = ttk.Frame(edit_content, padding=(14, 12))
		player_form.pack(fill="x", pady=(12, 0))
		ttk.Label(player_form, textvariable=self.selected_player_name, font=("Segoe UI", 12, "bold")).grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 10))
		ttk.Label(player_form, text="Global Rank").grid(row=1, column=0, sticky="w")
		global_rank_entry = ttk.Entry(player_form, textvariable=self.edit_global_rank, width=10)
		global_rank_entry.grid(row=2, column=0, sticky="w", padx=(0, 18))
		tk.Label(player_form, text="Positional Rank").grid(row=1, column=1, sticky="w")
		positional_rank_entry = ttk.Entry(player_form, textvariable=self.edit_positional_rank, width=10)
		positional_rank_entry.grid(row=2, column=1, sticky="w", padx=(0, 18))
		self.apply_button = ttk.Button(player_form, text="Apply to Player", command=self.apply_rank_edits)
		self.apply_button.grid(row=2, column=2, sticky="w", padx=(0, 8))
		ttk.Button(player_form, text="Save Changes", style="Accent.TButton", command=self.save_changes).grid(row=2, column=3, sticky="w")
		player_form.columnconfigure(4, weight=1)

		self.status = ttk.Label(self.root, padding=(18, 0, 18, 10))
		self.status.pack(fill="x")

	def create_rankings_table(self, parent: ttk.Frame, include_positional_rank: bool = True) -> ttk.Treeview:
		table_frame = ttk.Frame(parent)
		table_frame.pack(fill="both", expand=True)
		columns = ("global", "position", "positional", "name", "team") if include_positional_rank else ("global", "position", "name", "team")
		table = ttk.Treeview(table_frame, columns=columns, show="headings", selectmode="browse")
		column_details = (
			("global", "Global", 80),
			("position", "Position", 90),
			("positional", "Pos. Rank", 100),
			("name", "Player", 280),
			("team", "Team", 90),
		)
		if not include_positional_rank:
			column_details = tuple(column for column in column_details if column[0] != "positional")
		for column, heading, width in column_details:
			table.heading(column, text=heading)
			table.column(column, width=width, anchor="center" if column != "name" else "w")
		scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=table.yview)
		table.configure(yscrollcommand=scrollbar.set)
		table.pack(side="left", fill="both", expand=True)
		scrollbar.pack(side="right", fill="y")
		return table

	def build_positional_board(self) -> None:
		board = ttk.Frame(self.positional_tab, padding=(0, 4, 0, 6))
		board.pack(fill="both", expand=True)
		for index, position in enumerate(POSITIONS):
			group = ttk.Frame(board, padding=4)
			group.grid(row=index // 2, column=index % 2, sticky="nsew", padx=5, pady=5)
			ttk.Label(group, text=position, foreground="#F47721", font=("Segoe UI", 13, "bold")).pack(anchor="w", pady=(0, 5))
			table = ttk.Treeview(group, columns=("rank", "name", "team"), show="headings", height=10, style="Compact.Treeview")
			for column, heading, width in (("rank", "Rank", 58), ("name", "Player", 230), ("team", "Team", 70)):
				table.heading(column, text=heading)
				table.column(column, width=width, anchor="center" if column != "name" else "w")
			table.pack(fill="both", expand=True)
			self.positional_tables[position] = table
		for row in range(2):
			board.rowconfigure(row, weight=1)
		for column in range(2):
			board.columnconfigure(column, weight=1)

	def load_rankings(self) -> None:
		if self.dirty and not messagebox.askyesno("Discard edits?", "You have unsaved ranking changes. Discard them and reload the workbook?"):
			return
		try:
			self.players = load_players(self.excel_path)
			self.dirty = False
			self.refresh_table()
			self.refresh_positional_tables()
			self.refresh_editor_table()
		except (OSError, KeyError, TypeError, ValueError) as error:
			messagebox.showerror("Could not load rankings", str(error))

	def refresh_table(self) -> None:
		players = [
			(index, player)
			for index, player in enumerate(self.players)
			if player.team.casefold() != "fa"
			and (self.position.get() == "ALL" or player.position == self.position.get())
		]
		query = self.search.get().strip().casefold()
		if query:
			players = [(index, player) for index, player in players if query in player.name.casefold() or query in player.team.casefold()]
		if self.sort_by.get() == "position":
			players.sort(key=lambda item: (item[1].position, item[1].positional_rank))
		else:
			players.sort(key=lambda item: item[1].global_rank)
		for item in self.table.get_children():
			self.table.delete(item)
		for index, player in players:
			self.table.insert("", "end", iid=str(index), values=(player.global_rank, player.position, player.name, player.team))
		self.style_table_rows(self.table)
		message = "Unsaved changes" if self.dirty else ""
		self.status.configure(text=f"Showing {len(players):,} players  |  {message}")

	def refresh_positional_tables(self) -> None:
		for position, table in self.positional_tables.items():
			players = [player for player in self.players if player.position == position and player.team.casefold() != "fa"]
			players.sort(key=lambda player: (player.positional_rank, player.global_rank))
			for item in table.get_children():
				table.delete(item)
			for player in players[:10]:
				table.insert("", "end", values=(player.positional_rank, player.name, player.team))
			self.style_table_rows(table)

	def refresh_editor_table(self, selected_index: int | None = None) -> None:
		for item in self.editor_table.get_children():
			self.editor_table.delete(item)
		players = [(index, player) for index, player in enumerate(self.players) if player.team.casefold() != "fa"]
		players.sort(key=lambda item: item[1].global_rank)
		for index, player in players:
			self.editor_table.insert("", "end", iid=str(index), values=(player.global_rank, player.position, player.positional_rank, player.name, player.team))
		self.style_table_rows(self.editor_table)
		if selected_index is not None and self.editor_table.exists(str(selected_index)):
			self.editor_table.selection_set(str(selected_index))
			self.editor_table.focus(str(selected_index))
			self.set_selected_player(selected_index)
		else:
			self.selected_player_index = None
			self.selected_player_name.set("Select a player from the list")
			self.edit_global_rank.set("")
			self.edit_positional_rank.set("")

	def style_table_rows(self, table: ttk.Treeview) -> None:
		table.tag_configure("even", background="#202020")
		table.tag_configure("odd", background="#292929")
		for index, item in enumerate(table.get_children()):
			table.item(item, tags=("even" if index % 2 == 0 else "odd",))

	def select_player(self, _event: tk.Event) -> None:
		selection = self.editor_table.selection()
		if selection:
			self.set_selected_player(int(selection[0]))

	def set_selected_player(self, player_index: int) -> None:
		self.selected_player_index = player_index
		player = self.players[player_index]
		self.selected_player_name.set(f"{player.name}  /  {player.position}  /  {player.team}")
		self.edit_global_rank.set(str(player.global_rank))
		self.edit_positional_rank.set(str(player.positional_rank))

	def apply_rank_edits(self) -> None:
		if self.selected_player_index is None:
			messagebox.showinfo("Select a player", "Select a player before applying rank changes.")
			return
		try:
			global_rank = int(self.edit_global_rank.get())
			positional_rank = int(self.edit_positional_rank.get())
			if global_rank < 1 or positional_rank < 1:
				raise ValueError
		except ValueError:
			messagebox.showerror("Invalid rank", "Enter whole-number ranks greater than zero.")
			return
		players = list(self.players)
		players[self.selected_player_index] = replace(players[self.selected_player_index], global_rank=global_rank, positional_rank=positional_rank)
		self.players = tuple(players)
		self.dirty = True
		self.refresh_table()
		self.refresh_positional_tables()
		self.refresh_editor_table(self.selected_player_index)
		self.status.configure(text="Rank changes applied. Select Save Changes to write them to Excel.")

	def save_changes(self) -> None:
		if not self.dirty:
			self.status.configure(text="No unsaved changes")
			return
		try:
			save_rankings(self.excel_path, self.players)
		except (OSError, KeyError, TypeError, ValueError) as error:
			messagebox.showerror("Could not save rankings", str(error))
			return
		self.dirty = False
		self.refresh_table()
		self.refresh_positional_tables()
		self.refresh_editor_table(self.selected_player_index)
		self.status.configure(text=f"Rankings saved to {Path(self.excel_path).name}")

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
