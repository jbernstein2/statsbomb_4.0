"""
stats_processor.py

Processes StatsBomb CSV exports (all-events + crosses) and calculates
player- and team-level stats & KPIs for a single team.

Also includes helpers for assembling a two-team comparison table (used
by the Streamlit app / pptx generator), selecting top/bottom performers
for a stat, and resolving player positions (used to optionally exclude
goalkeepers from individual player rankings).

ASSUMPTIONS (flag these if your raw exports differ):
  * The crosses CSV has a boolean column "Blocked Event Type" marking
    a cross that was blocked.
  * Goalkeepers are identified via a "Player Position" column with the
    value "GK".

NOTE: the "forward first pass from outside the box" KPI is currently
disabled (not calculated, not included in the deck) per request.
"""

import pandas as pd
from collections import defaultdict


# ============================================================
# CONSTANTS
# ============================================================

TOUCH_EVENTS = [
    "ball_receipt",
    "pass",
    "carry",
    "dribble",
    "shot",
    "clearance"
]

PENALTY_ENTRY_EVENTS = [
    "carry",
    "dribble"
]

GOALKEEPER_POSITION_LABEL = "GK"
PLAYER_POSITION_COL = "Player Position"

# ------------------------------------------------------------
# KPIs organized into game phases, in the order they should
# appear in the deck. Each metric is (display label, stats key,
# direction) where direction is "higher" or "lower" indicating
# which side wins that category. Metrics with direction=None
# (see SUPPLEMENTARY_METRICS) are informational only and are not
# scored on the Full-Time summary slide.
#
# "ppda" is a manually-entered, team-only value (not derivable
# from the CSV export) - it has no per-player breakdown.
# ------------------------------------------------------------

GAME_PHASES = [
    {
        "name": "Build-Up & Possession",
        "metrics": [
            ("% Pass Forward", "forward_pass_percentage", "higher"),
            ("Possession Lost", "possession_lost", "lower"),
            ("Possession Lost (Defensive Half)", "possession_lost_defensive_half", "lower"),
        ],
    },
    {
        "name": "Defending",
        "metrics": [
            ("Attacking Half Recoveries", "attacking_half_recoveries", "higher"),
            ("PPDA (Attacking Half)", "ppda", "lower"),
        ],
    },
    {
        "name": "Progression",
        "metrics": [
            ("Touches in Attacking 1/3", "touches_attacking_third", "higher"),
            ("Completed Passes to Attacking 1/3", "passes_to_attacking_third", "higher"),
            ("PA Entry \u2013 Passes", "passes_to_penalty_area", "higher"),
            ("PA Entry \u2013 Dribble", "penalty_entries", "higher"),
        ],
    },
    {
        "name": "Final Third / Red Zone",
        "metrics": [
            ("Red Zone Touches", "red_zone_touches", "higher"),
            ("Red Zone Shots", "red_zone_shots", "higher"),
            ("Completed Passes to Red Zone", "red_zone_passes_completed", "higher"),
        ],
    },
    {
        "name": "Crossing",
        "metrics": [
            ("Open Play Crosses", "open_play_crosses", "higher"),
            ("Successful Open Play Crosses", "successful_open_play_crosses", "higher"),
            ("Blocked Crosses", "blocked_crosses", "lower"),
        ],
    },
]

# Flat (label, key, direction) list of every headline KPI, in
# game-phase order. This drives the Team Totals / Full-Time slides.
FLAT_MAIN_METRICS = [
    metric
    for phase in GAME_PHASES
    for metric in phase["metrics"]
]

# Same list but without PPDA (which has no per-player breakdown) -
# used for the per-metric player breakdown / appendix slides.
PLAYER_METRICS = [m for m in FLAT_MAIN_METRICS if m[1] != "ppda"]

# Quick lookup: stats key -> "higher" or "lower"
METRIC_DIRECTIONS = {key: direction for _, key, direction in FLAT_MAIN_METRICS}

# Backward-compatible (label, key) view of the headline metrics.
MAIN_METRICS = [(label, key) for label, key, _ in FLAT_MAIN_METRICS]

# Extra metrics that don't get their own slide but are useful in the
# supplementary KPI table / downloadable CSVs.
SUPPLEMENTARY_METRICS = [
    ("Total Passes", "total_passes", None),
    ("Forward Passes (Outside Box)", "forward_passes_outside_box", None),
]

# Ordered list of (Display Label, key, direction) used to build the
# full team comparison table (headline KPIs + supplementary KPIs).
COMPARISON_METRICS = FLAT_MAIN_METRICS + SUPPLEMENTARY_METRICS


# ============================================================
# DATA LOADING
# ============================================================

def load_team_data(events_file, crosses_file):
    """
    Load StatsBomb CSV exports.

    Parameters:
        events_file: All Events CSV path (or file-like object)
        crosses_file: Crosses CSV path (or file-like object)

    Returns:
        events dataframe
        crosses dataframe
    """

    events = pd.read_csv(events_file)
    crosses = pd.read_csv(crosses_file)

    events = clean_booleans(events)
    crosses = clean_booleans(crosses)

    return events, crosses


# ============================================================
# BOOLEAN CLEANING
# ============================================================

def clean_booleans(df):

    boolean_cols = [
        "Forward Pass",
        "In Attacking Half",
        "In Defensive Half",
        "In Attacking Third",
        "Into Attacking Third",
        "Into Opposition Penalty Box",
        "Open Play",
        "Blocked Event Type",
    ]

    for col in boolean_cols:

        if col in df.columns:

            df[col] = (
                df[col]
                .fillna(False)
                .astype(str)
                .str.lower()
                .eq("true")
            )

    return df


# ============================================================
# FIELD DEFINITIONS
# ============================================================

def is_penalty_area(x, y):

    return (
        pd.notna(x)
        and pd.notna(y)
        and x > 102
        and 18 < y < 62
    )


def is_red_zone(x, y):

    return (
        pd.notna(x)
        and pd.notna(y)
        and 100 < x < 110
        and 20 < y < 60
    )


# ============================================================
# PLAYER POSITIONS (for the optional goalkeeper filter)
# ============================================================

def get_player_positions(events):
    """
    Returns a {player: position} dict resolved from the events CSV's
    "Player Position" column (e.g. "GK", "CB", "ST", ...). Used to
    optionally exclude goalkeepers from individual player rankings.
    Team-level totals always include every player regardless of
    position.
    """

    if PLAYER_POSITION_COL not in events.columns or "Player" not in events.columns:
        return {}

    valid = events.dropna(subset=[PLAYER_POSITION_COL])

    if valid.empty:
        return {}

    return valid.groupby("Player")[PLAYER_POSITION_COL].first().to_dict()


def is_goalkeeper(position_value):
    return str(position_value).strip().upper() == GOALKEEPER_POSITION_LABEL


# ============================================================
# PLAYER STAT CALCULATION
# ============================================================

def calculate_player_stats(events, crosses):

    stats = defaultdict(lambda: {

        "forward_passes": 0,
        "total_passes": 0,
        "forward_pass_percentage": 0,

        "forward_passes_outside_box": 0,

        "attacking_half_recoveries": 0,

        "touches_attacking_third": 0,

        "passes_to_attacking_third": 0,

        "passes_to_penalty_area": 0,

        "penalty_entries": 0,

        "red_zone_touches": 0,

        "red_zone_shots": 0,

        "failed_passes": 0,

        "dispossessed": 0,

        "possession_lost": 0,

        "failed_passes_defensive_half": 0,

        "dispossessed_defensive_half": 0,

        "possession_lost_defensive_half": 0,

        "open_play_crosses": 0,

        "successful_open_play_crosses": 0,

        "blocked_crosses": 0,

        "red_zone_passes_attempted": 0,

        "red_zone_passes_completed": 0

    })

    # ----------------------------
    # PASSING
    # ----------------------------

    passes = events[
        events["Event Type"] == "pass"
    ]

    completed_passes = passes[
        passes["Outcome"] == "complete"
    ]

    for player, count in passes.groupby("Player").size().items():
        stats[player]["total_passes"] = count

    forward_passes = completed_passes[
        completed_passes["Forward Pass"]
    ]

    for player, count in forward_passes.groupby("Player").size().items():
        stats[player]["forward_passes"] = count

    for _, row in forward_passes.iterrows():

        if not is_penalty_area(
            row["Start X"],
            row["Start Y"]
        ):

            stats[row["Player"]]["forward_passes_outside_box"] += 1

    # ----------------------------
    # RECOVERIES
    # ----------------------------

    recoveries = events[
        (events["Event Type"] == "ball_recovery")
        &
        (events["In Attacking Half"])
    ]

    for player, count in recoveries.groupby("Player").size().items():
        stats[player]["attacking_half_recoveries"] = count

    # ----------------------------
    # TOUCHES
    # ----------------------------

    touches = events[
        (events["Event Type"].isin(TOUCH_EVENTS))
        &
        (events["In Attacking Third"])
    ]

    for player, count in touches.groupby("Player").size().items():
        stats[player]["touches_attacking_third"] = count

    # ----------------------------
    # PROGRESSION
    # ----------------------------

    final_third = completed_passes[
        completed_passes["Into Attacking Third"]
    ]

    for player, count in final_third.groupby("Player").size().items():
        stats[player]["passes_to_attacking_third"] = count

    box_passes = completed_passes[
        completed_passes["Into Opposition Penalty Box"]
    ]

    for player, count in box_passes.groupby("Player").size().items():
        stats[player]["passes_to_penalty_area"] = count

    entries = events[
        (events["Event Type"].isin(PENALTY_ENTRY_EVENTS))
        &
        (events["Into Opposition Penalty Box"])
    ]

    for player, count in entries.groupby("Player").size().items():
        stats[player]["penalty_entries"] = count

    # ----------------------------
    # RED ZONE
    # ----------------------------

    for _, row in events[
        events["Event Type"].isin(TOUCH_EVENTS)
    ].iterrows():

        if is_red_zone(
            row["Start X"],
            row["Start Y"]
        ):

            stats[row["Player"]]["red_zone_touches"] += 1

    for _, row in events[
        events["Event Type"] == "shot"
    ].iterrows():

        if is_red_zone(
            row["Start X"],
            row["Start Y"]
        ):

            stats[row["Player"]]["red_zone_shots"] += 1

    for _, row in passes.iterrows():

        if is_red_zone(
            row["Start X"],
            row["Start Y"]
        ):

            stats[row["Player"]]["red_zone_passes_attempted"] += 1

    for _, row in completed_passes.iterrows():

        if is_red_zone(
            row["Start X"],
            row["Start Y"]
        ):

            stats[row["Player"]]["red_zone_passes_completed"] += 1

    # ----------------------------
    # POSSESSION LOST
    # ----------------------------

    failed_passes = passes[
        passes["Outcome"] != "complete"
    ]

    for player, count in failed_passes.groupby("Player").size().items():
        stats[player]["failed_passes"] = count

    dispossessed = events[
        events["Event Type"] == "dispossessed"
    ]

    for player, count in dispossessed.groupby("Player").size().items():
        stats[player]["dispossessed"] = count

    failed_def = failed_passes[
        failed_passes["In Defensive Half"]
    ]

    for player, count in failed_def.groupby("Player").size().items():
        stats[player]["failed_passes_defensive_half"] = count

    dispossessed_def = dispossessed[
        dispossessed["In Defensive Half"]
    ]

    for player, count in dispossessed_def.groupby("Player").size().items():
        stats[player]["dispossessed_defensive_half"] = count

    # ----------------------------
    # CROSSES
    # ----------------------------

    crosses_open = crosses[
        crosses["Open Play"]
    ]

    for player, count in crosses_open.groupby("Player").size().items():
        stats[player]["open_play_crosses"] = count

    crosses_success = crosses[
        (crosses["Open Play"])
        &
        (crosses["Outcome"] == "complete")
    ]

    for player, count in crosses_success.groupby("Player").size().items():
        stats[player]["successful_open_play_crosses"] = count

    if "Blocked Event Type" in crosses.columns:

        crosses_blocked = crosses[
            crosses["Blocked Event Type"]
        ]

        for player, count in crosses_blocked.groupby("Player").size().items():
            stats[player]["blocked_crosses"] = count

    # ----------------------------
    # FINAL CALCULATIONS
    # ----------------------------

    for player in stats:

        total = stats[player]["total_passes"]

        if total:
            stats[player]["forward_pass_percentage"] = (
                stats[player]["forward_passes"]
                /
                total
                *
                100
            )

        stats[player]["possession_lost"] = (
            stats[player]["failed_passes"]
            +
            stats[player]["dispossessed"]
        )

        stats[player]["possession_lost_defensive_half"] = (
            stats[player]["failed_passes_defensive_half"]
            +
            stats[player]["dispossessed_defensive_half"]
        )

    return dict(stats)


# ============================================================
# GOALS (for the title-slide scoreline)
# ============================================================

def count_goals(events):
    """Count goals scored by a team from its All Events export."""
    shots = events[events["Event Type"] == "shot"]
    return int((shots["Outcome"] == "goal").sum())


# ============================================================
# TEAM CALCULATION
# ============================================================

def calculate_team_stats(player_stats):
    """
    Sums player_stats into team totals. Goalkeepers are always
    included here regardless of the "include goalkeepers in player
    rankings" UI toggle - that toggle only affects which players show
    up in the individual player breakdown charts.

    Note: this does NOT include "ppda", since PPDA is a manually
    entered, team-only value. Callers should set
    team_stats["ppda"] = <value> after calling this function.
    """

    if not player_stats:
        # No player events found - return a zeroed-out stat block so
        # downstream code (comparison tables, pptx) doesn't crash.
        return {key: 0 for _, key, _ in COMPARISON_METRICS if key != "ppda"}

    df = pd.DataFrame(player_stats).T

    team_stats = {}

    for column in df.columns:

        if column != "forward_pass_percentage":

            team_stats[column] = df[column].sum()

    if team_stats["total_passes"]:

        team_stats["forward_pass_percentage"] = (
            team_stats["forward_passes"]
            /
            team_stats["total_passes"]
            *
            100
        )

    else:

        team_stats["forward_pass_percentage"] = 0

    return team_stats


# ============================================================
# MAIN PIPELINE FUNCTION
# ============================================================

def analyze_team(events_file, crosses_file):
    """
    Complete StatsBomb processing pipeline.

    Returns:
        {
            "player_stats": dataframe,
            "team_stats": dictionary
        }
    """

    events, crosses = load_team_data(
        events_file,
        crosses_file
    )

    player_stats = calculate_player_stats(
        events,
        crosses
    )

    team_stats = calculate_team_stats(
        player_stats
    )

    return {

        "player_stats": pd.DataFrame(player_stats).T,

        "team_stats": team_stats

    }


# ============================================================
# TOP / BOTTOM PERFORMER SELECTION (for player breakdown charts)
# ============================================================

def best_worst_player(player_df, metric_col, direction="higher", include_goalkeepers=True,
                       position_col=PLAYER_POSITION_COL):
    """
    Returns ((best_player, best_value), (worst_player, worst_value)) for a
    metric, respecting direction ("higher" or "lower" is better). Returns
    (None, None) for a side if there's no eligible player data.

    include_goalkeepers=False drops rows where position_col == "GK" before
    picking the best/worst - used for the individual player callouts on
    each metric slide. Team totals are unaffected by this filter.
    """

    if player_df is None or player_df.empty or metric_col not in player_df.columns \
            or "Player" not in player_df.columns:
        return (None, None), (None, None)

    df = player_df.copy()

    if not include_goalkeepers and position_col in df.columns:
        df = df[~df[position_col].apply(is_goalkeeper)]

    df = df[["Player", metric_col]].dropna()

    if df.empty:
        return (None, None), (None, None)

    ascending = (direction == "lower")
    df_sorted = df.sort_values(metric_col, ascending=ascending)

    best_row = df_sorted.iloc[0]
    worst_row = df_sorted.iloc[-1]

    return (best_row["Player"], best_row[metric_col]), (worst_row["Player"], worst_row[metric_col])


def top_bottom_players(player_df, metric_col, n=3, include_goalkeepers=True,
                        position_col=PLAYER_POSITION_COL):
    """
    Returns a small dataframe with just the top-n and bottom-n players
    for a given metric (sorted best-to-worst), so bar charts stay
    readable. If there are 2n or fewer players to begin with, all of
    them are returned (no need to trim).

    include_goalkeepers=False drops rows where position_col == "GK"
    before selecting the top/bottom players. This only affects which
    players are eligible for the ranking - team totals are unaffected.
    """

    if player_df is None or player_df.empty or metric_col not in player_df.columns \
            or "Player" not in player_df.columns:
        return pd.DataFrame(columns=["Player", metric_col])

    df = player_df.copy()

    if not include_goalkeepers and position_col in df.columns:
        df = df[~df[position_col].apply(is_goalkeeper)]

    df = df[["Player", metric_col]].fillna(0)
    df = df.sort_values(metric_col, ascending=False)

    if len(df) <= 2 * n:
        return df

    return pd.concat([df.head(n), df.tail(n)])


# ============================================================
# TWO-TEAM COMPARISON TABLE (for the pptx report)
# ============================================================

def build_team_comparison(
    team_stats,
    opponent_stats,
    team_name,
    opponent_name,
    metrics=None,
):
    """
    Build a tidy comparison dataframe with one row per KPI and one
    column per team, ready to be written to CSV and consumed by the
    pptx generator.

    PPDA (key "ppda") is a manually-entered value - not present in the
    StatsBomb CSV export - so callers should merge it into team_stats
    / opponent_stats (e.g. team_stats["ppda"] = 9.4) before calling
    this function. Pass metrics=stats_processor.FLAT_MAIN_METRICS
    (etc.) to build a table scoped to a subset of KPIs, e.g. a single
    game phase.
    """

    if metrics is None:
        metrics = COMPARISON_METRICS

    rows = []

    for metric in metrics:

        label, key = metric[0], metric[1]
        team_val = team_stats.get(key, 0)
        opponent_val = opponent_stats.get(key, 0)

        rows.append({
            "Metric": label,
            team_name: round(float(team_val if team_val is not None else 0), 2),
            opponent_name: round(float(opponent_val if opponent_val is not None else 0), 2),
        })

    return pd.DataFrame(rows)


def player_stats_to_csv_df(player_stats_df, positions=None):
    """
    Normalize a player_stats dataframe (indexed by player name) into a
    flat dataframe with 'Player' as a regular column, suitable for
    writing to CSV. If a {player: position} dict is passed in (see
    get_player_positions), a "Player Position" column is attached.
    """

    df = player_stats_df.copy()
    df.index.name = "Player"
    df = df.reset_index()

    if positions:
        df[PLAYER_POSITION_COL] = df["Player"].map(positions)

    return df
