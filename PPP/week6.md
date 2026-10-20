# Week 6 - Trade calculator improvements

## Tasks for the week
- Fix and complete the trade-calculator workflow.
- Make player selection and trade evaluation usable from the rankings window.

## Sprint 1 plan
I made a substantial update to `Fantasy.py` focused on fixing the trade calculator. The current implementation lets me build the two sides of a trade, select players, calculate totals, and show a verdict. Trade values are checked before calculation, and the supported player range is limited to the top 150 global rankings.

## Code updates
Current locations in `Fantasy.py`:
- Lines 73-104: trade-value rules and trade verdict thresholds.
- Lines 137-161: command-line trade evaluation and validation.
- Lines 553-779: GUI trade calculator, player selection, and trade calculation.

## Line of Code
- **Lines changed:** 1060
- **Planned features:** 2/3
- **Planned requirements:** 5/6

## Outcome
The trade workflow was developed beyond the initial command-line foundation and integrated into the graphical application.
