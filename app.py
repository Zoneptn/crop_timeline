import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from pathlib import Path


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="Product Coverage",
    page_icon="🎯",
    layout="wide"
)


# ============================================================
# FILE (same workbook as the crop timeline dashboard)
# ============================================================

FILE_PATH = Path("crop_timeline.xlsx")


# ============================================================
# LOAD DATA
# (Same detection logic as the crop timeline app: match by
# sheet NAME first, fall back to matching by COLUMNS if a
# sheet gets renamed. Sheet order never matters.)
# ============================================================

@st.cache_data
def load_data(file_path):

    if not file_path.exists():
        st.error(f"File not found: {file_path}")
        st.stop()

    excel = pd.ExcelFile(file_path)

    def clean_columns(df):

        df = df.copy()

        df.columns = (
            df.columns
            .astype(str)
            .str.strip()
            .str.lower()
            .str.replace(" ", "_")
        )

        return df

    raw_sheets = {
        name: clean_columns(pd.read_excel(excel, sheet_name=name))
        for name in excel.sheet_names
    }

    sheet_lookup = {
        name.strip().lower(): name
        for name in raw_sheets.keys()
    }

    name_candidates = {
        "timeline": ["crop_stage", "timeline", "crop_timeline"],
        "pests": ["crop_pest", "crop_pests"],
        "weeds": ["crop_weeds", "crop_weed"],
        "diseases": ["crop_disease", "crop_diseases"],
        "weed_her": ["weed_her"],
        "pest_ins": ["pest_ins"],
        "disease_fun": ["disease_fun"]
    }

    assigned = {}
    detected_sheets = {}

    for role, candidates in name_candidates.items():

        for candidate in candidates:

            key = candidate.strip().lower()

            if key in sheet_lookup:
                actual_name = sheet_lookup[key]
                assigned[role] = raw_sheets[actual_name]
                detected_sheets[role] = f"{actual_name} (matched by name)"
                break

    timeline = assigned.get("timeline")
    pests = assigned.get("pests")
    weeds = assigned.get("weeds")
    diseases = assigned.get("diseases")
    weed_her = assigned.get("weed_her")
    pest_ins = assigned.get("pest_ins")
    disease_fun = assigned.get("disease_fun")

    # Fallback: match by columns for anything not found by name
    matched_sheet_names = {
        v.split(" (matched")[0] for v in detected_sheets.values()
    }

    leftover_sheets = {
        name: df
        for name, df in raw_sheets.items()
        if name not in matched_sheet_names
    }

    for sheet_name, df in leftover_sheets.items():

        cols = set(df.columns)

        if timeline is None and {"start_day", "end_day", "stage_id", "crop_id"}.issubset(cols):
            timeline = df
            detected_sheets["timeline"] = f"{sheet_name} (matched by columns)"

        elif pests is None and {"pest_id", "crop_id"}.issubset(cols) and "ins_id" not in cols:
            pests = df
            detected_sheets["pests"] = f"{sheet_name} (matched by columns)"

        elif weeds is None and {"weed_id", "crop_id"}.issubset(cols) and "her_id" not in cols:
            weeds = df
            detected_sheets["weeds"] = f"{sheet_name} (matched by columns)"

        elif diseases is None and {"disease_id", "crop_id"}.issubset(cols) and "fun_id" not in cols:
            diseases = df
            detected_sheets["diseases"] = f"{sheet_name} (matched by columns)"

        elif pest_ins is None and {"pest_id", "ins_id"}.issubset(cols):
            pest_ins = df
            detected_sheets["pest_ins"] = f"{sheet_name} (matched by columns)"

        elif weed_her is None and {"weed_id", "her_id"}.issubset(cols):
            weed_her = df
            detected_sheets["weed_her"] = f"{sheet_name} (matched by columns)"

        elif disease_fun is None and {"disease_id", "fun_id"}.issubset(cols):
            disease_fun = df
            detected_sheets["disease_fun"] = f"{sheet_name} (matched by columns)"

    required = {
        "crop timeline (crop_id, crop, stage_id, stage, start_day, end_day)": timeline,
        "pest stages (crop_id, pest_id, ...)": pests,
        "weed stages (crop_id, weed_id, ...)": weeds,
        "disease stages (crop_id, disease_id, ...)": diseases
    }

    missing = [label for label, df in required.items() if df is None]

    if missing:
        st.error(
            "Could not identify the following sheet(s):\n\n"
            + "\n".join(f"- {m}" for m in missing)
        )
        st.stop()

    for label, df in (
        ("weed_her", weed_her),
        ("pest_ins", pest_ins),
        ("disease_fun", disease_fun)
    ):
        if df is None:
            st.warning(
                f"Could not find a '{label}' sheet — that category "
                f"will show as 'no product' everywhere."
            )

    weed_her = weed_her if weed_her is not None else pd.DataFrame()
    pest_ins = pest_ins if pest_ins is not None else pd.DataFrame()
    disease_fun = disease_fun if disease_fun is not None else pd.DataFrame()

    datasets = [timeline, pests, weeds, diseases, weed_her, pest_ins, disease_fun]

    for df in datasets:
        for id_col in ("crop_id", "stage_id", "pest_id", "weed_id", "disease_id"):
            if id_col in df.columns:
                df[id_col] = df[id_col].astype(str).str.strip().str.lower()

    timeline["crop"] = timeline["crop"].astype(str).str.strip()
    timeline["start_day"] = pd.to_numeric(timeline["start_day"], errors="coerce")
    timeline["end_day"] = pd.to_numeric(timeline["end_day"], errors="coerce")

    return timeline, pests, weeds, diseases, weed_her, pest_ins, disease_fun


(
    timeline,
    pests,
    weeds,
    diseases,
    weed_her,
    pest_ins,
    disease_fun
) = load_data(FILE_PATH)


# ============================================================
# HEADER + SIDEBAR
# ============================================================

st.title("🎯 Product Coverage")

st.caption(
    "Green = we have a registered product for this threat. "
    "Red = we don't (either no product exists, or only competitors do)."
)

st.divider()

st.sidebar.header("Settings")

if st.sidebar.button("🔄 Reload data"):
    load_data.clear()
    st.rerun()

our_company = st.sidebar.text_input(
    "Your company name",
    value="SAC",
    help="Matched case-insensitively against the 'company_name' column."
).strip().lower()

crop_list = timeline["crop"].dropna().sort_values().unique()

selected_crop = st.sidebar.selectbox("Select Crop", crop_list)


# ============================================================
# FILTER TO SELECTED CROP
# ============================================================

crop_id = timeline.loc[timeline["crop"] == selected_crop, "crop_id"].iloc[0]

crop_timeline = timeline[timeline["crop_id"] == crop_id].copy().sort_values("start_day")

crop_pests = pests[pests["crop_id"] == crop_id].copy()
crop_weeds = weeds[weeds["crop_id"] == crop_id].copy()
crop_diseases = diseases[diseases["crop_id"] == crop_id].copy()

stage_lookup = crop_timeline[["stage_id", "stage", "start_day", "end_day"]].drop_duplicates()

crop_pests = crop_pests.merge(stage_lookup, on="stage_id", how="left")
crop_weeds = crop_weeds.merge(stage_lookup, on="stage_id", how="left")
crop_diseases = crop_diseases.merge(stage_lookup, on="stage_id", how="left")


# ============================================================
# COVERAGE LOOKUP
# For each item id, figure out:
#   - covered: True if any product row's company matches ours
#   - companies: the set of companies that DO have a product
#     (useful to see who the competitors are)
# ============================================================

def build_coverage_map(product_df, id_column, our_company_lower):

    coverage = {}

    if product_df.empty or id_column not in product_df.columns:
        return coverage

    for item_id, group in product_df.groupby(id_column):

        companies = (
            group.get("company_name", pd.Series(dtype=str))
            .dropna()
            .astype(str)
            .str.strip()
            .unique()
            .tolist()
        )

        is_covered = any(
            our_company_lower in c.lower()
            for c in companies
        )

        coverage[item_id] = {
            "covered": is_covered,
            "companies": companies
        }

    return coverage


weed_coverage = build_coverage_map(weed_her, "weed_id", our_company)
pest_coverage = build_coverage_map(pest_ins, "pest_id", our_company)
disease_coverage = build_coverage_map(disease_fun, "disease_id", our_company)


COVERED_COLOR = "#2ca02c"    # green
NOT_COVERED_COLOR = "#d62728"  # red


# ============================================================
# CATEGORY GROUPS
# ============================================================

CATEGORY_GROUPS = [
    {
        "label": "Weed",
        "icon": "🌱",
        "data": crop_weeds,
        "name_col": "weed_name_en",
        "thai_col": "weed_name_th",
        "id_col": "weed_id",
        "coverage_map": weed_coverage
    },
    {
        "label": "Insect",
        "icon": "🐛",
        "data": crop_pests,
        "name_col": "pest_name_en",
        "thai_col": "pest_name_th",
        "id_col": "pest_id",
        "coverage_map": pest_coverage
    },
    {
        "label": "Disease",
        "icon": "🍄",
        "data": crop_diseases,
        "name_col": "disease_name_en",
        "thai_col": "disease_name_th",
        "id_col": "disease_id",
        "coverage_map": disease_coverage
    }
]


# ============================================================
# CATEGORY SELECTOR
# ============================================================

category_options = [f"{g['icon']} {g['label']}s" for g in CATEGORY_GROUPS]

selected_categories = st.multiselect(
    "Categories to show",
    options=category_options,
    default=category_options
)

active_groups = [
    g for g, opt in zip(CATEGORY_GROUPS, category_options)
    if opt in selected_categories
]

if not active_groups:
    st.info("Select at least one category above to see the chart.")
    st.stop()


# ============================================================
# SUMMARY METRICS (coverage %)
# ============================================================

summary_cols = st.columns(len(active_groups))

for col, group in zip(summary_cols, active_groups):

    data = group["data"]
    name_col = group["name_col"]
    id_col = group["id_col"]
    coverage_map = group["coverage_map"]

    if data.empty or name_col not in data.columns:
        col.metric(f"{group['icon']} {group['label']}s covered", "—")
        continue

    unique_ids = data[id_col].dropna().unique()

    total = len(unique_ids)

    covered = sum(
        1 for i in unique_ids
        if coverage_map.get(i, {}).get("covered", False)
    )

    pct = f"{covered}/{total}" if total else "—"

    col.metric(f"{group['icon']} {group['label']}s covered", pct)


st.divider()


# ============================================================
# ROW SIZING
# ============================================================

for group in active_groups:

    if group["data"].empty or group["name_col"] not in group["data"].columns:
        group["item_count"] = 0
    else:
        group["item_count"] = group["data"][group["name_col"]].nunique()


row_heights = [max(g["item_count"], 1) for g in active_groups]

subplot_titles = [
    f"{g['icon']} {g['label']}s ({g['item_count']})"
    for g in active_groups
]

num_rows = len(active_groups)


# ============================================================
# BUILD FIGURE
# ============================================================

fig = make_subplots(
    rows=num_rows,
    cols=1,
    shared_xaxes=True,
    row_heights=row_heights,
    vertical_spacing=0.08,
    subplot_titles=subplot_titles
)

any_data = False
legend_shown = {"covered": False, "not_covered": False}

for row_idx, group in enumerate(active_groups, start=1):

    chart_data = group["data"]
    name_column = group["name_col"]
    thai_column = group["thai_col"]
    id_column = group["id_col"]
    coverage_map = group["coverage_map"]

    if chart_data.empty or name_column not in chart_data.columns:
        continue

    unique_items = chart_data[name_column].dropna().unique()

    for name in unique_items:

        item_data = chart_data[chart_data[name_column] == name].sort_values("start_day")

        for _, row in item_data.iterrows():

            start = row["start_day"]
            end = row["end_day"]

            if pd.isna(start) or pd.isna(end):
                continue

            duration = end - start

            thai_name = row.get(thai_column, "")

            item_id = row.get(id_column)

            info = coverage_map.get(item_id, {"covered": False, "companies": []})

            is_covered = info["covered"]

            companies = info["companies"]

            color = COVERED_COLOR if is_covered else NOT_COVERED_COLOR

            if is_covered:
                status_line = f"✅ We have a product ({our_company.upper()})"
            elif companies:
                status_line = (
                    "❌ No product from us — only: "
                    + ", ".join(companies)
                )
            else:
                status_line = "❌ No registered product at all"

            hover_text = (
                f"<b>{name}</b><br>"
                f"{thai_name}<br><br>"
                f"Stage: {row['stage']}<br>"
                f"Day {start:,.0f} – {end:,.0f}<br><br>"
                f"{status_line}"
            )

            legend_key = "covered" if is_covered else "not_covered"
            show_legend = not legend_shown[legend_key]
            legend_shown[legend_key] = True

            fig.add_trace(
                go.Bar(
                    x=[duration],
                    y=[name],
                    base=[start],
                    orientation="h",
                    marker_color=color,
                    hovertemplate=hover_text + "<extra></extra>",
                    name="We have a product" if is_covered else "No product from us",
                    legendgroup=legend_key,
                    showlegend=show_legend
                ),
                row=row_idx,
                col=1
            )

            any_data = True

    fig.update_yaxes(
        autorange="reversed",
        automargin=True,
        title="",
        row=row_idx,
        col=1
    )


# ============================================================
# STAGE BOUNDARIES + LABELS
# ============================================================

for _, row in crop_timeline.iterrows():

    fig.add_vline(x=row["start_day"], line_width=1, line_dash="dot", row="all", col=1)

if not crop_timeline.empty:
    fig.add_vline(x=crop_timeline["end_day"].max(), line_width=1, line_dash="dot", row="all", col=1)

for _, row in crop_timeline.iterrows():

    start = row["start_day"]
    end = row["end_day"]
    midpoint = (start + end) / 2

    stage_text = f"<b>{row['stage']}</b><br>Day {start:,.0f}–{end:,.0f}"

    xref_bottom = f"x{num_rows}" if num_rows > 1 else "x"

    fig.add_annotation(
        x=midpoint, y=-0.15, xref=xref_bottom, yref="paper",
        text=stage_text, showarrow=False, align="center", font=dict(size=11)
    )

    fig.add_annotation(
        x=midpoint, y=1.15, xref="x", yref="paper",
        text=stage_text, showarrow=False, align="center", font=dict(size=11)
    )


# ============================================================
# LAYOUT
# ============================================================

total_items = sum(row_heights)

chart_height = max(600, total_items * 38 + 260)

fig.update_layout(
    barmode="overlay",
    height=chart_height,
    margin=dict(l=20, r=20, t=180, b=180),
    hovermode="closest",
    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0)
)

fig.update_xaxes(
    title="Days After Planting", side="bottom", showgrid=True, zeroline=False,
    row=num_rows, col=1
)

for r in range(1, num_rows):
    fig.update_xaxes(showgrid=True, zeroline=False, row=r, col=1)


# ============================================================
# DISPLAY
# ============================================================

if not any_data:
    st.info(f"No data recorded for {selected_crop}.")
else:
    st.plotly_chart(fig, width='stretch')

st.caption(
    "Hover over a bar to see which companies (if any) have a "
    "registered product for that threat."
)
