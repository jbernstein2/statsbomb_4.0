"""
streamlit_app.py

Streamlit front-end that ties stats_processor.py and pptx_generator.py
together into a single report-generation workflow.
"""

import io
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

import stats_processor as sp
import pptx_generator as pg


st.set_page_config(page_title="Match Report Generator", layout="wide")

# Fixed defaults for Brooklyn FC
BKFC_NAME = "Brooklyn FC"
BKFC_COLOR = "#D4AF37"  # Gold

# Standard web color name mapping for typed text input
COLOR_NAME_MAP = {
    "black": "#000000",
    "gold": "#D4AF37",
    "silver": "#C0C0C0",
    "white": "#FFFFFF",
    "red": "#FF0000",
    "blue": "#0000FF",
    "navy": "#000080",
    "green": "#008000",
    "yellow": "#FFFF00",
    "orange": "#FFA500",
    "purple": "#800080",
    "grey": "#808080",
    "gray": "#808080",
}


def parse_color_input(color_str: str, default_hex: str = "#000000") -> tuple[str, bool]:
    """
    Parses user-typed color input (hex or color name).
    Returns (hex_code, is_valid). Defaults to `default_hex` if input is empty.
    """
    raw = color_str.strip().lower()
    if not raw:
        return default_hex, True

    if raw in COLOR_NAME_MAP:
        return COLOR_NAME_MAP[raw], True

    hex_clean = raw.lstrip("#")
    if len(hex_clean) == 6:
        try:
            int(hex_clean, 16)
            return f"#{hex_clean.upper()}", True
        except ValueError:
            return "", False
    elif len(hex_clean) == 3:
        try:
            full_hex = "".join([c * 2 for c in hex_clean])
            int(full_hex, 16)
            return f"#{full_hex.upper()}", True
        except ValueError:
            return "", False

    return "", False


def main():
    st.title("⚽ Match Report Generator")
    st.caption(
        "Upload StatsBomb CSV exports for both teams, fill in a couple of "
        "manual values, and generate a CSV + PowerPoint report organized "
        "by game phase (build-up, defending, progression, final third, crossing)."
    )

    # ------------------------------------------------------------
    # 1. File uploads
    # ------------------------------------------------------------
    st.header("1. Upload CSV files")

    col_bkfc, col_opp = st.columns(2)

    with col_bkfc:
        st.subheader("Brooklyn FC")
        bkfc_events_file = st.file_uploader(
            "All Events CSV — Brooklyn FC", type="csv", key="bkfc_events"
        )
        bkfc_crosses_file = st.file_uploader(
            "Crosses CSV — Brooklyn FC", type="csv", key="bkfc_crosses"
        )

    with col_opp:
        st.subheader("Opponent")
        opp_events_file = st.file_uploader(
            "All Events CSV — Opponent", type="csv", key="opp_events"
        )
        opp_crosses_file = st.file_uploader(
            "Crosses CSV — Opponent", type="csv", key="opp_crosses"
        )

    # ------------------------------------------------------------
    # 2. Manual PPDA input
    # ------------------------------------------------------------
    st.header("2. PPDA — Attacking Half (manual entry)")
    col_ppda_1, col_ppda_2 = st.columns(2)
    with col_ppda_1:
        bkfc_ppda = st.number_input("Input BKFC attacking-half PPDA", min_value=0.0, step=0.1, format="%.2f")
    with col_ppda_2:
        opponent_ppda = st.number_input("Input opponent attacking-half PPDA", min_value=0.0, step=0.1, format="%.2f")

    # ------------------------------------------------------------
    # 3. Team details (Brooklyn fixed; Opponent inputs)
    # ------------------------------------------------------------
    st.header("3. Team details")
    col_team_1, col_team_2 = st.columns(2)

    with col_team_1:
        st.subheader("Brooklyn FC")
        st.info(f"**Team Name:** {BKFC_NAME}\n\n**Team Color:** Gold (`{BKFC_COLOR}`)")

    with col_team_2:
        st.subheader("Opponent Details")
        opponent_display_name = st.text_input(
            "Opponent team name * (Required)",
            value="",
            placeholder="e.g. Rhode Island FC"
        )
        opponent_color_raw = st.text_input(
            "Opponent team color (Optional - type color name or hex code)",
            value="",
            placeholder="Defaults to black if blank"
        )

    # ------------------------------------------------------------
    # 4. Player ranking options
    # ------------------------------------------------------------
    st.header("4. Player ranking options")
    include_goalkeepers = st.checkbox(
        "Include goalkeepers in individual player stat rankings (best/worst-3 charts)",
        value=False,
        help="Goalkeepers are always included in team-level totals regardless of this "
             "setting — this only controls whether they're eligible to appear in the "
             "per-player best/worst-3 breakdown charts. Detected via a \"GK\" value in "
             "the Player Position column.",
    )

    # ------------------------------------------------------------
    # 5. Average position + shot chart images (both teams)
    # ------------------------------------------------------------
    st.header("5. Team visuals")
    st.caption("Upload average positioning and shot chart images for both teams.")
    col_visuals_bkfc, col_visuals_opp = st.columns(2)
    with col_visuals_bkfc:
        st.subheader("Brooklyn FC")
        bkfc_avg_position_file = st.file_uploader(
            "Average Position — Brooklyn FC", type=["png", "jpg", "jpeg"], key="bkfc_avg_position"
        )
        bkfc_shot_chart_file = st.file_uploader(
            "Shot Chart — Brooklyn FC", type=["png", "jpg", "jpeg"], key="bkfc_shot_chart"
        )
    with col_visuals_opp:
        st.subheader("Opponent")
        opp_avg_position_file = st.file_uploader(
            "Average Position — Opponent", type=["png", "jpg", "jpeg"], key="opp_avg_position"
        )
        opp_shot_chart_file = st.file_uploader(
            "Shot Chart — Opponent", type=["png", "jpg", "jpeg"], key="opp_shot_chart"
        )

    # ------------------------------------------------------------
    # 6. Optional report metadata
    # ------------------------------------------------------------
    st.header("6. Report details")
    col_meta_1, col_meta_2 = st.columns(2)
    with col_meta_1:
        kicker = st.text_input("Eyebrow / kicker text", value="STATSBOMB MATCH RECAP")
        match_date = st.text_input("Match date (footer, e.g. 7-18-2026)", value="")
    with col_meta_2:
        subtitle = st.text_input(
            "Subtitle",
            value="STATISTICAL MATCH ANALYSIS  \u2022  PLAYER & TEAM PERFORMANCE BREAKDOWN",
        )
        include_supplementary_table = st.checkbox(
            "Include supplementary KPI table (secondary stats)", value=True
        )

    st.divider()

    required_files_present = all([
        bkfc_events_file, bkfc_crosses_file, opp_events_file, opp_crosses_file
    ])

    if not required_files_present:
        st.info("Upload all four CSV files above to enable report generation.")
        return

    # Validation for Opponent Name
    if not opponent_display_name.strip():
        st.error("Opponent team name is required.")
        return

    if opponent_display_name.strip().lower() == BKFC_NAME.lower():
        st.error("Opponent team name cannot be 'Brooklyn FC'.")
        return

    # Validation & Parsing for Opponent Color
    opponent_color, is_color_valid = parse_color_input(opponent_color_raw, default_hex="#000000")
    if not is_color_valid:
        st.error("Invalid opponent color. Please enter a valid hex code (e.g. #FF0000) or color name (e.g. 'black', 'navy', 'red').")
        return

    generate = st.button("Generate report", type="primary")

    if not generate:
        return

    # ------------------------------------------------------------
    # 7. Run the stats pipeline
    # ------------------------------------------------------------
    with st.spinner("Crunching the numbers..."):
        try:
            bkfc_events_df, bkfc_crosses_df = sp.load_team_data(bkfc_events_file, bkfc_crosses_file)
            opp_events_df, opp_crosses_df = sp.load_team_data(opp_events_file, opp_crosses_file)
        except Exception as e:
            st.error(f"Failed to process the uploaded CSVs: {e}")
            return

        bkfc_player_stats = sp.calculate_player_stats(bkfc_events_df, bkfc_crosses_df)
        bkfc_team_stats = sp.calculate_team_stats(bkfc_player_stats)
        bkfc_team_stats["ppda"] = bkfc_ppda
        bkfc_goals = sp.count_goals(bkfc_events_df)
        bkfc_positions = sp.get_player_positions(bkfc_events_df)

        opp_player_stats = sp.calculate_player_stats(opp_events_df, opp_crosses_df)
        opp_team_stats = sp.calculate_team_stats(opp_player_stats)
        opp_team_stats["ppda"] = opponent_ppda
        opp_goals = sp.count_goals(opp_events_df)
        opp_positions = sp.get_player_positions(opp_events_df)

        comparison_df = sp.build_team_comparison(
            bkfc_team_stats,
            opp_team_stats,
            BKFC_NAME,
            opponent_display_name.strip(),
        )

        bkfc_players_df = sp.player_stats_to_csv_df(pd.DataFrame(bkfc_player_stats).T, positions=bkfc_positions)
        opponent_players_df = sp.player_stats_to_csv_df(pd.DataFrame(opp_player_stats).T, positions=opp_positions)

    st.success("Stats processed.")

    st.subheader("Final Score")
    st.markdown(f"**{BKFC_NAME} {bkfc_goals} — {opp_goals} {opponent_display_name.strip()}**")

    st.subheader("Team KPI Comparison")
    st.dataframe(comparison_df, use_container_width=True)

    tab_bkfc, tab_opp = st.tabs([f"{BKFC_NAME} players", f"{opponent_display_name.strip()} players"])
    with tab_bkfc:
        st.dataframe(bkfc_players_df, use_container_width=True)
    with tab_opp:
        st.dataframe(opponent_players_df, use_container_width=True)

    # ------------------------------------------------------------
    # 8. Write CSVs + build the pptx in a temp workspace
    # ------------------------------------------------------------
    with st.spinner("Building the slide deck..."):
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)

            comparison_csv_path = tmp / "team_comparison.csv"
            bkfc_players_csv_path = tmp / f"{BKFC_NAME}_players.csv"
            opponent_players_csv_path = tmp / f"{opponent_display_name.strip()}_players.csv"

            comparison_df.to_csv(comparison_csv_path, index=False)
            bkfc_players_df.to_csv(bkfc_players_csv_path, index=False)
            opponent_players_df.to_csv(opponent_players_csv_path, index=False)

            def _save_upload(uploaded_file, stem):
                if uploaded_file is None:
                    return None
                path = tmp / f"{stem}{Path(uploaded_file.name).suffix}"
                path.write_bytes(uploaded_file.getvalue())
                return path

            bkfc_avg_position_path = _save_upload(bkfc_avg_position_file, "bkfc_avg_position")
            bkfc_shot_chart_path = _save_upload(bkfc_shot_chart_file, "bkfc_shot_chart")
            opp_avg_position_path = _save_upload(opp_avg_position_file, "opp_avg_position")
            opp_shot_chart_path = _save_upload(opp_shot_chart_file, "opp_shot_chart")

            pptx_output_path = tmp / "match_report.pptx"

            try:
                pg.generate_report(
                    comparison_csv=comparison_csv_path,
                    team_name=BKFC_NAME,
                    team_color=BKFC_COLOR,
                    opponent_name=opponent_display_name.strip(),
                    opponent_color=opponent_color,
                    output_path=pptx_output_path,
                    team_score=bkfc_goals,
                    opponent_score=opp_goals,
                    date_label=match_date or None,
                    kicker=kicker,
                    subtitle=subtitle,
                    team_player_stats_csv=bkfc_players_csv_path,
                    opponent_player_stats_csv=opponent_players_csv_path,
                    team_avg_position_image=bkfc_avg_position_path,
                    team_shot_chart_image=bkfc_shot_chart_path,
                    opponent_avg_position_image=opp_avg_position_path,
                    opponent_shot_chart_image=opp_shot_chart_path,
                    include_supplementary_table=include_supplementary_table,
                    include_goalkeepers=include_goalkeepers,
                )
            except Exception as e:
                st.error(f"Failed to build the PowerPoint report: {e}")
                return

            pptx_bytes = pptx_output_path.read_bytes()
            comparison_csv_bytes = comparison_csv_path.read_bytes()
            bkfc_players_csv_bytes = bkfc_players_csv_path.read_bytes()
            opponent_players_csv_bytes = opponent_players_csv_path.read_bytes()

    st.success("Report generated!")

    st.header("7. Downloads")
    col_dl_1, col_dl_2 = st.columns(2)
    with col_dl_1:
        st.download_button(
            "⬇️ Download PowerPoint report",
            data=pptx_bytes,
            file_name="match_report.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            type="primary",
        )
    with col_dl_2:
        st.download_button(
            "⬇️ Download team comparison CSV",
            data=comparison_csv_bytes,
            file_name="team_comparison.csv",
            mime="text/csv",
        )

    col_dl_3, col_dl_4 = st.columns(2)
    with col_dl_3:
        st.download_button(
            f"⬇️ Download {BKFC_NAME} player stats CSV",
            data=bkfc_players_csv_bytes,
            file_name=f"{BKFC_NAME}_players.csv",
            mime="text/csv",
        )
    with col_dl_4:
        st.download_button(
            f"⬇️ Download {opponent_display_name.strip()} player stats CSV",
            data=opponent_players_csv_bytes,
            file_name=f"{opponent_display_name.strip()}_players.csv",
            mime="text/csv",
        )


if __name__ == "__main__":
    main()
