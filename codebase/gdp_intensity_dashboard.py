#%%
"""
Interactive Plotly dashboard for the GDP-intensity vs 9th projection comparison.

Reads data/temp/gdp_intensity_comparison.csv (output of gdp_intensity_comparison_workflow.py)
and writes a self-contained HTML file to data/temp/gdp_intensity_dashboard.html.

No server required — open the HTML in any browser.
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots

REPO_ROOT = Path(__file__).resolve().parents[1]
CURRENT_DIR = Path.cwd()
if CURRENT_DIR != REPO_ROOT:
    os.chdir(REPO_ROOT)
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

INPUT_CSV = "data/temp/gdp_intensity_comparison.csv"
OUTPUT_HTML = "data/temp/gdp_intensity_dashboard.html"

ECONOMY_LABELS = {
    "20USA": "USA (20)",
    "05PRC": "China (05)",
    "08JPN": "Japan (08)",
    "01AUS": "Australia (01)",
    "09ROK": "Korea (09)",
    "11MEX": "Mexico (11)",
}

FLOW_COLORS = {
    "Agriculture":        {"gdp": "#2196F3", "ninth": "#0D47A1"},
    "Fishing":            {"gdp": "#4CAF50", "ninth": "#1B5E20"},
    "Non-specified others": {"gdp": "#FF9800", "ninth": "#E65100"},
}

FLOWS = ["Agriculture", "Fishing", "Non-specified others"]


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def economy_label(key: str) -> str:
    return ECONOMY_LABELS.get(key, key)


def build_energy_figure(df: pd.DataFrame, economies: list[str]) -> go.Figure:
    """
    3-column subplot (one per flow), one row per economy via dropdown.
    Shows GDP-derived (solid) vs 9th TGT (dashed) in PJ.
    """
    fig = make_subplots(
        rows=1,
        cols=3,
        subplot_titles=FLOWS,
        shared_yaxes=False,
        horizontal_spacing=0.08,
    )

    # Build one set of traces per economy; only the first economy is visible.
    all_traces: list[tuple[go.Scatter, int]] = []  # (trace, col)

    for econ in economies:
        econ_df = df[df["economy_key"] == econ]
        visible = econ == economies[0]

        for col_idx, flow in enumerate(FLOWS, start=1):
            flow_df = econ_df[econ_df["sector_label"] == flow].sort_values("year")
            colors = FLOW_COLORS.get(flow, {"gdp": "#888", "ninth": "#333"})

            # GDP-derived line
            all_traces.append((
                go.Scatter(
                    x=flow_df["year"],
                    y=flow_df["gdp_energy_pj"],
                    name="GDP-derived",
                    mode="lines+markers",
                    line=dict(color=colors["gdp"], width=2),
                    marker=dict(size=5),
                    legendgroup=f"gdp_{flow}",
                    showlegend=(col_idx == 1),
                    visible=visible,
                    hovertemplate="%{x}: %{y:.2f} PJ<extra>GDP-derived</extra>",
                ),
                col_idx,
            ))

            # 9th TGT line
            all_traces.append((
                go.Scatter(
                    x=flow_df["year"],
                    y=flow_df["ninth_energy_pj"],
                    name="9th TGT",
                    mode="lines+markers",
                    line=dict(color=colors["ninth"], width=2, dash="dash"),
                    marker=dict(size=5, symbol="diamond"),
                    legendgroup=f"ninth_{flow}",
                    showlegend=(col_idx == 1),
                    visible=visible,
                    hovertemplate="%{x}: %{y:.2f} PJ<extra>9th TGT</extra>",
                ),
                col_idx,
            ))

    for trace, col in all_traces:
        fig.add_trace(trace, row=1, col=col)

    traces_per_economy = len(FLOWS) * 2  # gdp + ninth per flow

    # Dropdown buttons — one per economy
    buttons = []
    for i, econ in enumerate(economies):
        vis = [False] * len(all_traces)
        start = i * traces_per_economy
        for j in range(traces_per_economy):
            vis[start + j] = True
        buttons.append(dict(
            label=economy_label(econ),
            method="update",
            args=[
                {"visible": vis},
                {"title": f"Energy (PJ) — {economy_label(econ)}"},
            ],
        ))

    fig.update_layout(
        title=f"Energy (PJ) — {economy_label(economies[0])}",
        updatemenus=[dict(
            buttons=buttons,
            direction="down",
            x=0.0,
            xanchor="left",
            y=1.18,
            yanchor="top",
            showactive=True,
            bgcolor="#f0f0f0",
            bordercolor="#cccccc",
        )],
        legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5),
        height=480,
        margin=dict(t=120, b=100),
        plot_bgcolor="#fafafa",
    )
    fig.update_xaxes(title_text="Year", tickmode="linear", dtick=5)
    fig.update_yaxes(title_text="PJ")
    return fig


def build_diff_figure(df: pd.DataFrame, economies: list[str]) -> go.Figure:
    """
    3-column subplot showing diff % (GDP vs 9th TGT) per flow, per economy.
    All economies are shown as separate lines (no dropdown) for direct comparison.
    """
    fig = make_subplots(
        rows=1,
        cols=3,
        subplot_titles=FLOWS,
        shared_yaxes=True,
        horizontal_spacing=0.06,
    )

    palette = ["#E91E63", "#2196F3", "#4CAF50", "#FF9800", "#9C27B0", "#00BCD4"]

    for econ_idx, econ in enumerate(economies):
        econ_df = df[df["economy_key"] == econ]
        color = palette[econ_idx % len(palette)]

        for col_idx, flow in enumerate(FLOWS, start=1):
            flow_df = econ_df[econ_df["sector_label"] == flow].sort_values("year")
            valid = flow_df.dropna(subset=["diff_pct"])
            if valid.empty:
                continue

            fig.add_trace(
                go.Scatter(
                    x=valid["year"],
                    y=valid["diff_pct"],
                    name=economy_label(econ),
                    mode="lines+markers",
                    line=dict(color=color, width=2),
                    marker=dict(size=5),
                    legendgroup=econ,
                    showlegend=(col_idx == 1),
                    hovertemplate="%{x}: %{y:.1f}%<extra>" + economy_label(econ) + "</extra>",
                ),
                row=1,
                col=col_idx,
            )

    # Zero reference line per subplot
    for col_idx in range(1, 4):
        fig.add_hline(y=0, line_dash="dot", line_color="grey", line_width=1, row=1, col=col_idx)

    fig.update_layout(
        title="Divergence: GDP-derived vs 9th TGT (diff %)",
        legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5),
        height=450,
        margin=dict(t=80, b=100),
        plot_bgcolor="#fafafa",
    )
    fig.update_xaxes(title_text="Year", tickmode="linear", dtick=5)
    fig.update_yaxes(title_text="Diff %", col=1)
    return fig


def build_intensity_figure(df: pd.DataFrame, economies: list[str]) -> go.Figure:
    """
    Implied energy intensity (energy_pj / gdp_index) shown alongside the flat
    GDP-derived intensity, to visualise how much the 9th bakes in efficiency gains.

    Uses ninth_energy_pj / gdp_index as the 'implied 9th intensity'.
    Since GDP-derived intensity is flat (= base-year energy), we show the 9th
    implied intensity declining relative to it.
    """
    # We need the GDP index — reconstruct from the ratio gdp_energy_pj / base_energy.
    # base_energy = gdp_energy_pj at base year (index = 1 there).
    fig = make_subplots(
        rows=1,
        cols=3,
        subplot_titles=FLOWS,
        shared_yaxes=False,
        horizontal_spacing=0.08,
    )

    palette = ["#E91E63", "#2196F3", "#4CAF50", "#FF9800", "#9C27B0", "#00BCD4"]

    for econ_idx, econ in enumerate(economies):
        econ_df = df[df["economy_key"] == econ]
        color = palette[econ_idx % len(palette)]

        for col_idx, flow in enumerate(FLOWS, start=1):
            flow_df = econ_df[econ_df["sector_label"] == flow].sort_values("year")

            # Recover base-year energy (= flat GDP-derived intensity since index=1 at base)
            # gdp_energy_pj = base_energy * gdp_index  → gdp_index = gdp_energy / base_energy
            # We can infer base_energy from the fact that at base year gdp_index == 1,
            # but the base year row has NaN (it's in ESTO only). Use year with min diff% ≈ 0
            # as proxy, or just use the first non-NaN gdp_energy row.
            valid = flow_df.dropna(subset=["gdp_energy_pj", "ninth_energy_pj"])
            if valid.empty or valid["gdp_energy_pj"].iloc[0] == 0:
                continue

            # gdp_index for each year = gdp_energy / gdp_energy[first year]
            first_gdp = valid["gdp_energy_pj"].iloc[0]
            gdp_index = valid["gdp_energy_pj"] / first_gdp

            # Implied 9th intensity = ninth_energy / gdp_index  (relative to first year)
            implied_ninth = valid["ninth_energy_pj"] / gdp_index

            fig.add_trace(
                go.Scatter(
                    x=valid["year"],
                    y=implied_ninth,
                    name=economy_label(econ),
                    mode="lines+markers",
                    line=dict(color=color, width=2),
                    marker=dict(size=5),
                    legendgroup=econ,
                    showlegend=(col_idx == 1),
                    hovertemplate="%{x}: %{y:.2f}<extra>" + economy_label(econ) + "</extra>",
                ),
                row=1,
                col=col_idx,
            )

    # Flat reference line = 1.0 (GDP-derived, constant intensity)
    for col_idx in range(1, 4):
        fig.add_hline(
            y=valid["ninth_energy_pj"].iloc[0] if not valid.empty else 1,
            line_dash="dot",
            line_color="grey",
            line_width=1,
            annotation_text="GDP flat",
            annotation_position="right",
            row=1,
            col=col_idx,
        )

    fig.update_layout(
        title="Implied 9th energy intensity relative to GDP (indexed to first projection year)",
        legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5),
        height=450,
        margin=dict(t=80, b=100),
        plot_bgcolor="#fafafa",
    )
    fig.update_xaxes(title_text="Year", tickmode="linear", dtick=5)
    fig.update_yaxes(title_text="Intensity (indexed)")
    return fig


# ---------------------------------------------------------------------------
# Assemble and save
# ---------------------------------------------------------------------------

def build_dashboard(
    input_csv: str = INPUT_CSV,
    output_html: str = OUTPUT_HTML,
) -> Path:
    df = pd.read_csv(input_csv)
    economies = [e for e in ECONOMY_LABELS if e in df["economy_key"].unique()]
    if not economies:
        # Fall back to whatever is in the file
        economies = sorted(df["economy_key"].unique().tolist())

    print(f"Building dashboard for economies: {economies}")

    fig_energy = build_energy_figure(df, economies)
    fig_diff   = build_diff_figure(df, economies)
    fig_intens = build_intensity_figure(df, economies)

    # Combine into one HTML file
    html_parts = []
    html_parts.append("""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>GDP-intensity vs 9th projection</title>
<style>
  body { font-family: Arial, sans-serif; background: #fff; margin: 0; padding: 20px; }
  h1   { font-size: 1.4em; color: #333; margin-bottom: 4px; }
  p    { color: #666; font-size: 0.9em; margin-top: 0; }
  .section { margin-bottom: 40px; }
  hr   { border: none; border-top: 1px solid #eee; margin: 30px 0; }
</style>
</head>
<body>
<h1>GDP-intensity vs 9th Projection &mdash; Minor Demand Sectors</h1>
<p>Agriculture &bull; Fishing &bull; Non-specified others &nbsp;|&nbsp;
   Base year: 2022 &nbsp;|&nbsp; GDP scenario: Reference &nbsp;|&nbsp;
   Intensity held constant at base year for GDP approach</p>
<hr>
""")

    html_parts.append('<div class="section">')
    html_parts.append(fig_energy.to_html(full_html=False, include_plotlyjs="cdn"))
    html_parts.append("</div><hr>")

    html_parts.append('<div class="section">')
    html_parts.append(fig_diff.to_html(full_html=False, include_plotlyjs=False))
    html_parts.append("</div><hr>")

    html_parts.append('<div class="section">')
    html_parts.append(fig_intens.to_html(full_html=False, include_plotlyjs=False))
    html_parts.append("</div>")

    html_parts.append("</body></html>")

    out_path = Path(output_html)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text("\n".join(html_parts), encoding="utf-8")
    print(f"Dashboard saved to: {out_path.resolve()}")
    return out_path


#%%
if __name__ == "__main__":
    path = build_dashboard()
    # Auto-open in default browser
    import webbrowser
    webbrowser.open(path.resolve().as_uri())
