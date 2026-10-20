# Week 8 - Rank changes and tests

## Tasks for the week
- Make rank edits shift affected players consistently.
- Recalculate positional rankings when global rankings change.
- Add tests for rank movement behavior.

## Sprint 1 plan
I changed rank editing so moving one player shifts the players between the old and new positions rather than leaving duplicate or skipped ranks. Global rank edits recalculate positional ranks in global-rank order; positional-only edits shift players at the same position. I added unit tests for moving ranks up and down and for positional-rank behavior.

## Code updates
- `Fantasy.py` lines 36-71: `shift_player_rank` and `recalculate_positional_ranks`.
- `Fantasy.py` lines 918-977: apply and save ranking edits in the GUI.
- `tests/test_rank_shifts.py` lines 1-51: tests for global and positional rank shifts.

## Line of Code
- **Lines changed:** 1200
- **Planned features:** 3/3
- **Planned requirements:** 6/6

## Outcome
Rank-edit behavior became more predictable, with automated tests covering the key cases.
