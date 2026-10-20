# Week 2 - Rankings foundation

## Tasks for the week
- Build the initial rankings application.
- Obtain player data and store rankings in a usable format.
- Provide ways to view rankings and compare players.

## Sprint 1 plan
I added the first full version of `Fantasy.py`, the initial `players.xlsx` workbook, and the `openpyxl` dependency. The application defines player records, retrieves active QB/RB/WR/TE data from Sleeper, writes and reads the Excel workbook, displays rankings, and supports command-line ranking and trade workflows.

## Code updates
Current locations in `Fantasy.py`:
- Lines 26-34: `Player` data model.
- Lines 36-161: ranking and trade calculations, player lookup, and command-line output.
- Lines 163-235: Sleeper data retrieval and Excel workbook read/write.
- Lines 984-1029: command-line arguments and application entry point.

## Line of Code
- **Lines changed:** 1000
- **Planned features:** 0/3
- **Planned requirements:** 1/6

## Outcome
The project moved from an initial description to a working rankings and trade-calculation foundation backed by a workbook.
