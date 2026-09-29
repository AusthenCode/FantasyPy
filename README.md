# FantasyPy

A fantasy football trade calculator for quarterbacks, running backs, wide receivers, and tight ends.

The calculator can fetch current NFL player data from Sleeper's public API and save it as an Excel workbook. The workbook includes player name, position, team, global rank, positional rank, trade value, and Sleeper player ID.

## Usage

Open the rankings screen. This reads `players.xlsx` and does not print rankings to the terminal:

```powershell
python Fantasy.py
```

The window uses a black, orange, and white theme. Free agents are omitted from the rankings and search results. The **Global Rankings** tab shows global rank and lets you filter by QB, RB, WR, or TE, sort, and search players. **Positional Rankings** shows the top 10 players at QB, RB, WR, and TE. On the **Edit Rankings** tab, select a player to load their ranks into the form, then choose **Apply to Player** and **Save Changes** to update `players.xlsx`. Trade values continue to be calculated and stored in the workbook, but are not shown in the rankings window. **Refresh** reloads the workbook and prompts before discarding unsaved ranking changes. You can also launch the editor explicitly:

```powershell
python Fantasy.py gui --excel players.xlsx
```

The command-line board is still available when needed:

```powershell
python Fantasy.py rankings
```

Fetch the player list into Excel:

```powershell
python Fantasy.py fetch --output players.xlsx
```

Use the Excel rankings in the calculator:

```powershell
python Fantasy.py rankings --excel players.xlsx --position WR --sort position --limit 10
python Fantasy.py trade --excel players.xlsx --give "CeeDee Lamb" --get "Breece Hall" "Sam LaPorta"
```

Show the top wide receivers by positional rank:

```powershell
python Fantasy.py rankings --position WR --sort position --limit 10
```

Compare a trade. Quote player names containing spaces:

```powershell
python Fantasy.py trade --give "CeeDee Lamb" --get "Breece Hall" "Sam LaPorta"
```

The API's player ranking order becomes global rank, and positional rank is calculated separately within QB, RB, WR, and TE. The rankings screen and command-line commands read from `players.xlsx`; changes saved in the rankings window are available to both.