**# Match Report Generator

Turns StatsBomb CSV exports (all-events + crosses, for two teams) into:

1. A **team KPI comparison CSV** (+ per-player stat CSVs)
2. A formatted **PowerPoint report** (title slide, KPI table(s), KPI
   comparison chart, top-performers tables, and an average-position
   image slide)

## Files

| File | Purpose |
|---|---|
| `stats_processor.py` | Loads StatsBomb CSVs and calculates player/team KPIs. Also builds the two-team comparison table (folding in a manually-entered PPDA value). |
| `pptx_generator.py` | Reads the CSVs (+ optional average-position image) and builds the `.pptx` deck. No Streamlit dependency — usable standalone. |
| `streamlit_app.py` | The Streamlit UI: file uploads, PPDA / team-name / color inputs, and download buttons for the CSVs + final deck. |

## Running the app

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

## Workflow

1. Upload the **All Events** and **Crosses** CSVs for Brooklyn FC (BKFC) and
   for the opponent (4 files total).
2. Enter **BKFC ppda** and **opponent ppda** manually — PPDA isn't
   derivable from the standard StatsBomb CSV export, so this is a
   direct numeric input.
3. Enter each team's **display name** (as it should appear in the
   deck) and pick a **brand color** for each.
4. Optionally upload a **BKFC average position PNG** — it becomes its
   own slide in the deck.
5. Click **Generate report** to see the KPI table and player tables
   in-app, then download:
   - `match_report.pptx`
   - `team_comparison.csv`
   - `<team>_players.csv` / `<opponent>_players.csv`

## Using `pptx_generator.py` standalone

```python
import pptx_generator as pg

pg.generate_report(
    comparison_csv="team_comparison.csv",
    team_name="Brooklyn FC",
    team_color="#0B3D91",
    opponent_name="Sporting JAX",
    opponent_color="#F2A900",
    output_path="match_report.pptx",
    match_title="Sporting JAX vs. Brooklyn FC",
    team_player_stats_csv="bkfc_players.csv",       # optional
    opponent_player_stats_csv="opp_players.csv",     # optional
    avg_position_image="avg_position.png",           # optional
    subtitle="USL Championship - 2026-07-19",        # optional
)
```

## Notes / customizing KPIs

- The set of KPIs that appear in the comparison table is defined in
  `stats_processor.COMPARISON_METRICS` — add/remove tuples of
  `(display label, team_stats key)` there.
- The KPIs charted on the bar-chart slide are controlled by
  `pptx_generator.CHART_METRICS` (must match labels used in
  `COMPARISON_METRICS`).
- The "Top Performers" table columns are controlled by
  `pptx_generator.TOP_PLAYERS_DISPLAY_COLS`, and which stat they're
  ranked by is `TOP_PLAYERS_SORT_COL`.
**
