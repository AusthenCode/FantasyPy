# Week 5 - Rankings interface and tabs

## Tasks for the week
- Add a graphical interface for using the rankings.
- Organize the application into screens for its main workflows.
- Add controls to browse and search player rankings.

## Sprint 1 plan
I expanded `Fantasy.py` with a Tkinter interface and created separate Trade Calculator, Global Rankings, Positional Rankings, and Edit Rankings tabs. I added reusable visual components and connected the rankings view to the Excel workbook, including search and refresh controls.

## Code updates
Current locations in `Fantasy.py`:
- Lines 288-399: reusable rounded controls and trade-player display components.
- Lines 400-552: `RankingsWindow` setup and tab layout.
- Lines 780-838: positional board, workbook loading, and global-rank table refresh.
- Lines 978-982: GUI launch function.

## Line of Code
- **Lines changed:** 1000
- **Planned features:** 1/3
- **Planned requirements:** 3/6

## Outcome
The project gained a desktop interface alongside its existing command-line workflows.
