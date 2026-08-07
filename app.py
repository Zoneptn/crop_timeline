import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from pathlib import Path
import unicodedata


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="Crop Timeline",
    page_icon="🌱",
    layout="wide"
)


# ============================================================
# FILE
# ============================================================

FILE_PATH = Path("crop_timeline.xlsx")


# ============================================================
# LOAD DATA
# ============================================================

@st.cache_data
def load_data(file_path):

    if not file_path.exists():
        st.error(f"File not found: {file_path}")
        st.stop()

    excel = pd.ExcelFile(file_path)

    # -----------------------------------------
    # Read every sheet and clean its columns
    # first. We then figure out WHICH sheet is
    # which by looking at its columns, not its
    # position or name — this way it doesn't
    # matter what order the sheets are in, or
    # if more sheets get added later.
    # -----------------------------------------

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

    # Known tab names -> role. Listed in priority
    # order; first match wins.
    name_candidates = {
        "timeline": ["crop_stage", "timeline", "crop_timeline"],
        "pests": ["crop_pest", "crop_pests"],
        "weeds": ["crop_weeds", "crop_weed"],
        "diseases": ["crop_disease", "crop_diseases"],
        "weed_her": ["weed_her"],
        "pest_ins": ["pest_ins"],
        "disease_fun": ["disease_fun"],
        "stage_fert": ["stage_fert", "crop_fert", "fert_stage", "fertilizer"]
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
    stage_fert = assigned.get("stage_fert")

    # -----------------------------------------
    # Fallback: for any role NOT matched by name
    # (e.g. a sheet got renamed), try matching by
    # column signature among leftover sheets.
    # -----------------------------------------

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

        elif stage_fert is None and {"stage_id", "fertilizer_id"}.issubset(cols):
            stage_fert = df
            detected_sheets["stage_fert"] = f"{sheet_name} (matched by columns)"

        else:
            detected_sheets[f"UNMATCHED: {sheet_name}"] = list(cols)

    # -----------------------------------------
    # The 4 core sheets are required
    # -----------------------------------------

    required = {
        "crop timeline (crop_id, crop, stage_id, stage, start_day, end_day)": timeline,
        "pest stages (crop_id, pest_id, ...)": pests,
        "weed stages (crop_id, weed_id, ...)": weeds,
        "disease stages (crop_id, disease_id, ...)": diseases
    }

    missing = [label for label, df in required.items() if df is None]

    if missing:
        st.error(
            "Could not identify the following sheet(s) by their "
            "columns — check the workbook structure:\n\n"
            + "\n".join(f"- {m}" for m in missing)
        )
        st.stop()

    # -----------------------------------------
    # Product sheets are optional — warn but
    # don't crash if one is missing
    # -----------------------------------------

    for label, df in (
        ("weed_her", weed_her),
        ("pest_ins", pest_ins),
        ("disease_fun", disease_fun),
        ("stage_fert", stage_fert)
    ):
        if df is None:
            st.warning(
                f"Could not find a '{label}' sheet (matched by columns). "
                f"That category will be unavailable."
            )

    weed_her = weed_her if weed_her is not None else pd.DataFrame()
    pest_ins = pest_ins if pest_ins is not None else pd.DataFrame()
    disease_fun = disease_fun if disease_fun is not None else pd.DataFrame()
    stage_fert = stage_fert if stage_fert is not None else pd.DataFrame()

    # -----------------------------------------
    # Guard against Excel silently converting a
    # value like "20-10-10" into an actual date.
    # If that happened, pandas reads it back as a
    # Timestamp — reconstruct it as day-month-year
    # text so it displays as a formula again, not
    # a date.
    # -----------------------------------------

    def fix_date_like_column(df, col):

        if col not in df.columns or df.empty:
            return df

        def restore_text(value):

            if isinstance(value, pd.Timestamp):
                return f"{value.day}-{value.month}-{value.year % 100}"

            if pd.isna(value):
                return ""

            return str(value).strip()

        df[col] = df[col].apply(restore_text)

        return df

    stage_fert = fix_date_like_column(stage_fert, "fertilizer_formula")

    # -----------------------------------------
    # Real-world column typos in the 3 product
    # sheets: "concetration" (missing an 'n') and
    # "formuation_type" (missing an 'l'). Rename
    # them to the canonical spelling so the rest
    # of the app doesn't need to know about this.
    # -----------------------------------------

    typo_fix_map = {
        "concetration": "concentration",
        "formuation_type": "formulation_type"
    }

    for df in (weed_her, pest_ins, disease_fun):
        df.rename(columns=typo_fix_map, inplace=True)

    # -----------------------------------------
    # crop_disease has a duplicated "stage_id"
    # column in the workbook — pandas renames the
    # second one to "stage_id.1" on read. The sheet
    # also carries its own start_day/end_day,
    # which would collide with stage_lookup's
    # versions when merged later (producing
    # start_day_x/start_day_y and silently
    # breaking the chart). Drop all of these here
    # so diseases gets its start_day/end_day the
    # same way every other sheet does: from the
    # stage merge, single source of truth.
    # -----------------------------------------

    dup_stage_cols = [
        col for col in diseases.columns
        if col.startswith("stage_id.")
    ]

    drop_cols = [
        col for col in (["start_day", "end_day"] + dup_stage_cols)
        if col in diseases.columns
    ]

    if drop_cols:
        diseases = diseases.drop(columns=drop_cols)

    datasets = [
        timeline,
        pests,
        weeds,
        diseases,
        weed_her,
        pest_ins,
        disease_fun,
        stage_fert
    ]

    # Normalize id columns used for joins
    for df in datasets:

        for id_col in (
            "crop_id",
            "stage_id",
            "pest_id",
            "weed_id",
            "disease_id",
            "fertilizer_id"
        ):

            if id_col in df.columns:
                df[id_col] = (
                    df[id_col]
                    .astype(str)
                    .str.strip()
                    .str.lower()
                )

    # Clean crop name
    timeline["crop"] = (
        timeline["crop"]
        .astype(str)
        .str.strip()
    )

    # Numeric days
    timeline["start_day"] = pd.to_numeric(
        timeline["start_day"],
        errors="coerce"
    )

    timeline["end_day"] = pd.to_numeric(
        timeline["end_day"],
        errors="coerce"
    )

    return (
        timeline,
        pests,
        weeds,
        diseases,
        weed_her,
        pest_ins,
        disease_fun,
        stage_fert,
        detected_sheets
    )


(
    timeline,
    pests,
    weeds,
    diseases,
    weed_her,
    pest_ins,
    disease_fun,
    stage_fert,
    detected_sheets
) = load_data(FILE_PATH)


# ============================================================
# DEBUG: which physical sheet mapped to which role
# ============================================================

with st.expander("🔍 Sheet detection (debug)", expanded=False):

    st.write(
        "This shows which sheet in your workbook was matched to "
        "each expected role, based on its columns. If a role is "
        "missing or points to the wrong sheet, check that sheet's "
        "column names."
    )

    st.json(detected_sheets)

    st.write("Column counts detected per role:")

    st.write(
        {
            "timeline": list(timeline.columns),
            "pests": list(pests.columns),
            "weeds": list(weeds.columns),
            "diseases": list(diseases.columns),
            "weed_her": list(weed_her.columns),
            "pest_ins": list(pest_ins.columns),
            "disease_fun": list(disease_fun.columns),
            "stage_fert": list(stage_fert.columns)
        }
    )


# ============================================================
# HEADER
# ============================================================

st.title("🌱 Crop Growth Timeline")

st.caption(
    "Crop development stages with insect pests, diseases and weeds."
)

st.divider()


# ============================================================
# SIDEBAR
# ============================================================

st.sidebar.header("Dashboard")

if st.sidebar.button("🔄 Reload data"):
    load_data.clear()
    st.rerun()

dashboard_view = st.sidebar.radio(
    "View",
    ["📅 Reference Timeline", "🎯 Product Coverage"],
    index=0
)

st.sidebar.divider()

st.sidebar.header("Crop Selection")

crop_list = (
    timeline["crop"]
    .dropna()
    .sort_values()
    .unique()
)

selected_crop = st.sidebar.selectbox(
    "Select Crop",
    crop_list
)

our_company_raw = "SAC"

if dashboard_view == "🎯 Product Coverage":

    st.sidebar.divider()

    st.sidebar.header("Coverage Settings")

    our_company_raw = st.sidebar.text_input(
        "Your company name",
        value="SAC",
        help=(
            "Matched against the 'company_name' column, ignoring case, "
            "extra whitespace, and Unicode variations."
        )
    )


# ============================================================
# GET CROP ID
# ============================================================

crop_id = (
    timeline.loc[
        timeline["crop"] == selected_crop,
        "crop_id"
    ]
    .iloc[0]
)


# ============================================================
# FILTER CROP
# ============================================================

crop_timeline = timeline[
    timeline["crop_id"] == crop_id
].copy()

crop_pests = pests[
    pests["crop_id"] == crop_id
].copy()

crop_weeds = weeds[
    weeds["crop_id"] == crop_id
].copy()

crop_diseases = diseases[
    diseases["crop_id"] == crop_id
].copy()

# Fertilizer sheet carries crop_id directly, same as
# the other 3 sheets, so filter the same way.
if stage_fert.empty or "crop_id" not in stage_fert.columns:
    crop_fert = stage_fert.copy()
else:
    crop_fert = stage_fert[
        stage_fert["crop_id"] == crop_id
    ].copy()

    # The sheet also carries its own "stage" column,
    # which would collide with stage_lookup's "stage"
    # column on merge (producing stage_x/stage_y). Drop
    # it here — we'll get the authoritative version
    # (with matching start_day/end_day) from the merge.
    if "stage" in crop_fert.columns:
        crop_fert = crop_fert.drop(columns=["stage"])


crop_timeline = crop_timeline.sort_values(
    "start_day"
)


# ============================================================
# MERGE TIMELINE DAYS INTO PEST / WEED / DISEASE / FERTILIZER
# ============================================================

stage_lookup = crop_timeline[
    [
        "stage_id",
        "stage",
        "start_day",
        "end_day"
    ]
].drop_duplicates()


crop_pests = crop_pests.merge(
    stage_lookup,
    on="stage_id",
    how="left"
)


crop_weeds = crop_weeds.merge(
    stage_lookup,
    on="stage_id",
    how="left"
)


crop_diseases = crop_diseases.merge(
    stage_lookup,
    on="stage_id",
    how="left"
)


if not crop_fert.empty and "stage_id" in crop_fert.columns:
    crop_fert = crop_fert.merge(
        stage_lookup,
        on="stage_id",
        how="left"
    )


# ============================================================
# SUMMARY
# ============================================================

st.header(f"🌿 {selected_crop.title()}")


def safe_nunique(df, col):

    if col not in df.columns:

        st.warning(
            f"Column '{col}' not found — showing 0. "
            f"Check the '🔍 Sheet detection' panel above."
        )

        return 0

    return df[col].nunique()


total_days = crop_timeline["end_day"].max()

total_stages = safe_nunique(crop_timeline, "stage_id")

total_pests = safe_nunique(crop_pests, "pest_id")

total_diseases = safe_nunique(crop_diseases, "disease_id")

total_weeds = safe_nunique(crop_weeds, "weed_id")


c1, c2, c3, c4, c5 = st.columns(5)


c1.metric(
    "Growth Period",
    f"{total_days:,.0f} days"
)

c2.metric(
    "Growth Stages",
    total_stages
)

c3.metric(
    "Insect Pests",
    total_pests
)

c4.metric(
    "Diseases",
    total_diseases
)

c5.metric(
    "Weeds",
    total_weeds
)


st.divider()


# ============================================================
# TEXT NORMALIZATION
# (used for matching company names — handles case, extra
# whitespace, non-breaking spaces, and Unicode variations
# such as differently-composed Thai combining characters,
# which can make two visually-identical strings compare as
# unequal with a plain .strip().lower())
# ============================================================

def normalize_text(value):

    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""

    text = str(value)

    # Thai SARA AM decomposed form (NIKHAHIT + SARA AA) ->
    # precomposed form. Standard Unicode NFC normalization does
    # NOT fix this on its own — Thai text extracted from PDFs or
    # typed on different systems can end up in either form even
    # though both render identically as "ำ".
    text = text.replace("\u0e4d\u0e32", "\u0e33")

    text = unicodedata.normalize("NFC", text)

    # Strip invisible formatting characters (zero-width space,
    # zero-width joiner, BOM, etc.) — these don't count as
    # whitespace to Python's .strip()/.split(), so a company name
    # with one hidden inside it would otherwise silently fail to
    # match even though it looks identical on screen.
    text = "".join(
        ch for ch in text
        if unicodedata.category(ch) != "Cf"
    )

    # Collapse every whitespace variant (including non-breaking
    # space) down to single regular spaces.
    text = "".join(
        " " if ch.isspace() else ch
        for ch in text
    )

    text = " ".join(text.split())

    return text.strip()


# ============================================================
# PRODUCT LOOKUPS (shared by both views)
# ============================================================

PRODUCT_COLUMNS = [
    "brand_name",
    "company_name",
    "common_name",
    "concentration",
    "formulation_type"
]


def build_product_map(product_df, id_column):

    if product_df.empty or id_column not in product_df.columns:
        return {}

    return (
        product_df
        .groupby(id_column)
        .apply(lambda d: d.to_dict("records"))
        .to_dict()
    )


def format_products(product_map, product_id):

    products = product_map.get(product_id, [])

    if not products:
        return "No registered products"

    lines = []

    for p in products:

        parts = [
            str(p.get(col, "")).strip()
            for col in PRODUCT_COLUMNS
            if str(p.get(col, "")).strip()
            and str(p.get(col, "")).strip().lower() != "nan"
        ]

        lines.append("• " + " | ".join(parts))

    return "<br>".join(lines)


def compute_coverage(product_map, item_id, our_company_lower):
    """Returns (is_covered, list_of_normalized_company_names, our_products).

    our_products is the list of raw product records (brand_name,
    common_name, etc.) whose company matches ours — used to show
    exactly which product(s) we have, not just that we have one.
    """

    products = product_map.get(item_id, [])

    companies = sorted({
        normalize_text(p.get("company_name", ""))
        for p in products
        if normalize_text(p.get("company_name", ""))
    })

    our_products = [
        p for p in products
        if our_company_lower in normalize_text(p.get("company_name", "")).lower()
    ]

    is_covered = len(our_products) > 0

    return is_covered, companies, our_products


def format_our_products(our_products):
    """Bullet list of brand + common name for each of our products.
    Single product still gets a bullet, for visual consistency."""

    lines = []

    for p in our_products:

        brand = str(p.get("brand_name", "")).strip()
        common = str(p.get("common_name", "")).strip()

        parts = [
            v for v in (brand, common)
            if v and v.lower() != "nan"
        ]

        lines.append("• " + " — ".join(parts) if parts else "• (unnamed product)")

    return "<br>".join(lines) if lines else "• (product details unavailable)"


weed_product_map = build_product_map(weed_her, "weed_id")

pest_product_map = build_product_map(pest_ins, "pest_id")

disease_product_map = build_product_map(disease_fun, "disease_id")


COVERED_COLOR = "#2ca02c"     # green
NOT_COVERED_COLOR = "#d62728"  # red


# ============================================================
# SHARED GANTT FIGURE BUILDER
# Used by BOTH views, so they can never drift out of sync —
# stage boundaries, top/bottom labels, x-axis range, and
# layout are all defined exactly once.
# ============================================================

def build_gantt_figure(active_groups, crop_timeline, show_legend=False):

    for group in active_groups:

        data = group["data"]
        name_col = group["name_col"]

        if data.empty or name_col not in data.columns:
            group["item_count"] = 0
        else:
            group["item_count"] = data[name_col].nunique()

    row_heights = [max(g["item_count"], 1) for g in active_groups]

    subplot_titles = [
        f"{g['icon']} {g['label']}s ({g['item_count']})"
        for g in active_groups
    ]

    num_rows = len(active_groups)

    fig = make_subplots(
        rows=num_rows,
        cols=1,
        shared_xaxes=True,
        row_heights=row_heights,
        vertical_spacing=0.06,
        subplot_titles=subplot_titles
    )

    any_data = False
    legend_shown = {}

    for row_idx, group in enumerate(active_groups, start=1):

        data = group["data"]
        name_col = group["name_col"]
        get_trace_info = group["get_trace_info"]

        if data.empty or name_col not in data.columns:
            continue

        unique_items = data[name_col].dropna().unique()

        for name in unique_items:

            item_data = data[data[name_col] == name].sort_values("start_day")

            for _, row in item_data.iterrows():

                start = row["start_day"]
                end = row["end_day"]

                if pd.isna(start) or pd.isna(end):
                    continue

                duration = end - start

                color, hover_text, legend_key, trace_name = get_trace_info(row, name)

                show_this_legend = not legend_shown.get(legend_key, False)
                legend_shown[legend_key] = True

                fig.add_trace(
                    go.Bar(
                        x=[duration],
                        y=[name],
                        base=[start],
                        orientation="h",
                        marker_color=color,
                        hovertemplate=hover_text + "<extra></extra>",
                        name=trace_name,
                        legendgroup=legend_key,
                        showlegend=show_this_legend
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

    # --------------------------------------------------
    # STAGE BOUNDARIES (span all active rows)
    # --------------------------------------------------

    for _, row in crop_timeline.iterrows():

        fig.add_vline(
            x=row["start_day"],
            line_width=1,
            line_dash="dot",
            row="all",
            col=1
        )

    if not crop_timeline.empty:

        fig.add_vline(
            x=crop_timeline["end_day"].max(),
            line_width=1,
            line_dash="dot",
            row="all",
            col=1
        )

    # --------------------------------------------------
    # STAGE LABELS — top AND bottom of the whole figure
    # --------------------------------------------------

    xref_bottom = f"x{num_rows}" if num_rows > 1 else "x"

    for _, row in crop_timeline.iterrows():

        start = row["start_day"]
        end = row["end_day"]
        midpoint = (start + end) / 2

        stage_text = (
            f"<b>{row['stage']}</b>"
            f"<br>"
            f"Day {start:,.0f}–{end:,.0f}"
        )

        fig.add_annotation(
            x=midpoint, y=-0.2, xref=xref_bottom, yref="paper",
            text=stage_text, showarrow=False, align="center",
            font=dict(size=11)
        )

        fig.add_annotation(
            x=midpoint, y=1.2, xref="x", yref="paper",
            text=stage_text, showarrow=False, align="center",
            font=dict(size=11)
        )

    # --------------------------------------------------
    # LAYOUT
    # --------------------------------------------------

    total_items = sum(row_heights)

    chart_height = max(700, total_items * 38 + 320)

    layout_kwargs = dict(
        barmode="overlay",
        height=chart_height,
        margin=dict(l=20, r=20, t=270, b=250),
        hovermode="closest",
        showlegend=show_legend
    )

    if show_legend:
        layout_kwargs["legend"] = dict(
            orientation="h", yanchor="bottom", y=1.3, xanchor="left", x=0
        )

    fig.update_layout(**layout_kwargs)

    # Always show the FULL crop lifetime on the x-axis
    # (day 0 to the final stage's end_day), regardless of
    # where the actual bars happen to start/end — otherwise
    # Plotly autoscales to just the bar data and can visually
    # crop the timeline.
    if not crop_timeline.empty:
        lifetime_end = crop_timeline["end_day"].max()
        pad = max(2, lifetime_end * 0.02)
        x_range = [0 - pad, lifetime_end + pad]
    else:
        x_range = None

    fig.update_xaxes(
        title="Days After Planting", side="bottom", showgrid=True,
        zeroline=False, range=x_range, row=num_rows, col=1
    )

    for r in range(1, num_rows):
        fig.update_xaxes(
            showgrid=True, zeroline=False, range=x_range, row=r, col=1
        )

    return fig, any_data


# ============================================================
# CATEGORY SELECTOR (shared by both views)
# ============================================================

category_options = ["🌱 Weeds", "🐛 Insects", "🍄 Diseases", "🧪 Fertilizer"]

selected_categories = st.multiselect(
    "Categories to show",
    options=category_options,
    default=category_options,
    help="Remove a category to hide its panel. Leave all selected to see everything together."
)


# ============================================================
# REFERENCE TIMELINE — groups
# ============================================================

def build_reference_groups():

    def weed_info(row, name):
        thai_name = row.get("weed_name_th", "")
        product_text = format_products(weed_product_map, row.get("weed_id"))
        hover = (
            f"<b>{name}</b><br>{thai_name}<br><br>"
            f"Stage: {row['stage']}<br>"
            f"Day {row['start_day']:,.0f} – {row['end_day']:,.0f}<br><br>"
            f"<b>Products:</b><br>{product_text}"
        )
        return "#2ca02c", hover, "weed_ref", "Weeds"

    def pest_info(row, name):
        thai_name = row.get("pest_name_th", "")
        product_text = format_products(pest_product_map, row.get("pest_id"))
        hover = (
            f"<b>{name}</b><br>{thai_name}<br><br>"
            f"Stage: {row['stage']}<br>"
            f"Day {row['start_day']:,.0f} – {row['end_day']:,.0f}<br><br>"
            f"<b>Products:</b><br>{product_text}"
        )
        return "#d62728", hover, "insect_ref", "Insects"

    def disease_info(row, name):
        thai_name = row.get("disease_name_th", "")
        product_text = format_products(disease_product_map, row.get("disease_id"))
        hover = (
            f"<b>{name}</b><br>{thai_name}<br><br>"
            f"Stage: {row['stage']}<br>"
            f"Day {row['start_day']:,.0f} – {row['end_day']:,.0f}<br><br>"
            f"<b>Products:</b><br>{product_text}"
        )
        return "#9467bd", hover, "disease_ref", "Diseases"

    def fert_info(row, name):
        hover = (
            f"<b>🧪 {name}</b><br>"
            f"Brand: {row.get('fertilizer_brand', '—')}<br>"
            f"Company: {row.get('fertilizer_company', '—')}<br><br>"
            f"Stage: {row['stage']}<br>"
            f"Day {row['start_day']:,.0f} – {row['end_day']:,.0f}"
        )
        return "#1f77b4", hover, "fert_ref", "Fertilizer"

    return [
        {
            "label": "Weed", "icon": "🌱", "option_label": "🌱 Weeds", "data": crop_weeds,
            "name_col": "weed_name_en", "thai_col": "weed_name_th",
            "id_col": "weed_id", "mode": "product_lookup",
            "get_trace_info": weed_info
        },
        {
            "label": "Insect", "icon": "🐛", "option_label": "🐛 Insects", "data": crop_pests,
            "name_col": "pest_name_en", "thai_col": "pest_name_th",
            "id_col": "pest_id", "mode": "product_lookup",
            "get_trace_info": pest_info
        },
        {
            "label": "Disease", "icon": "🍄", "option_label": "🍄 Diseases", "data": crop_diseases,
            "name_col": "disease_name_en", "thai_col": "disease_name_th",
            "id_col": "disease_id", "mode": "product_lookup",
            "get_trace_info": disease_info
        },
        {
            "label": "Fertilizer", "icon": "🧪", "option_label": "🧪 Fertilizer", "data": crop_fert,
            "name_col": "fertilizer_formula", "thai_col": None,
            "id_col": "fertilizer_id", "mode": "direct",
            "get_trace_info": fert_info
        }
    ]


# ============================================================
# PRODUCT COVERAGE — groups
# ============================================================

def build_coverage_groups(our_company_lower):

    coverage_debug_rows = []

    def make_threat_info(product_map, thai_col, id_col, category_label):

        def info(row, name):

            item_id = row.get(id_col)

            is_covered, companies, our_products = compute_coverage(
                product_map, item_id, our_company_lower
            )

            thai_name = row.get(thai_col, "") if thai_col else ""

            color = COVERED_COLOR if is_covered else NOT_COVERED_COLOR

            if is_covered:
                product_list = format_our_products(our_products)
                status = f"✅ We have a product:<br>{product_list}"
            elif companies:
                status = "❌ No product from us — only: " + ", ".join(companies)
            else:
                status = "❌ No registered product at all"

            hover = (
                f"<b>{name}</b><br>{thai_name}<br><br>"
                f"Stage: {row['stage']}<br>"
                f"Day {row['start_day']:,.0f} – {row['end_day']:,.0f}<br><br>"
                f"{status}"
            )

            legend_key = "covered" if is_covered else "not_covered"
            trace_name = "We have a product" if is_covered else "No product from us"

            our_product_names = "; ".join(
                " — ".join(
                    v for v in (
                        str(p.get("brand_name", "")).strip(),
                        str(p.get("common_name", "")).strip()
                    )
                    if v and v.lower() != "nan"
                )
                for p in our_products
            )

            coverage_debug_rows.append({
                "Category": category_label,
                "Item": name,
                "Item ID": item_id,
                "Covered": "✅" if is_covered else "❌",
                "Our product(s)": our_product_names if our_product_names else "—",
                "Companies found": ", ".join(companies) if companies else "(none)"
            })

            return color, hover, legend_key, trace_name

        return info

    def fert_info(row, name):
        hover = (
            f"<b>🧪 {name}</b><br>"
            f"Brand: {row.get('fertilizer_brand', '—')}<br>"
            f"Company: {row.get('fertilizer_company', '—')}<br><br>"
            f"Stage: {row['stage']}<br>"
            f"Day {row['start_day']:,.0f} – {row['end_day']:,.0f}"
        )
        return COVERED_COLOR, hover, "fertilizer", "Our fertilizer recommendation"

    groups = [
        {
            "label": "Weed", "icon": "🌱", "option_label": "🌱 Weeds", "data": crop_weeds,
            "name_col": "weed_name_en",
            "get_trace_info": make_threat_info(
                weed_product_map, "weed_name_th", "weed_id", "Weed"
            )
        },
        {
            "label": "Insect", "icon": "🐛", "option_label": "🐛 Insects", "data": crop_pests,
            "name_col": "pest_name_en",
            "get_trace_info": make_threat_info(
                pest_product_map, "pest_name_th", "pest_id", "Insect"
            )
        },
        {
            "label": "Disease", "icon": "🍄", "option_label": "🍄 Diseases", "data": crop_diseases,
            "name_col": "disease_name_en",
            "get_trace_info": make_threat_info(
                disease_product_map, "disease_name_th", "disease_id", "Disease"
            )
        },
        {
            "label": "Fertilizer", "icon": "🧪", "option_label": "🧪 Fertilizer", "data": crop_fert,
            "name_col": "fertilizer_formula",
            "get_trace_info": fert_info
        }
    ]

    return groups, coverage_debug_rows


def category_coverage_summary(data, id_col, product_map, our_company_lower):

    if data.empty or id_col not in data.columns:
        return 0, 0

    unique_ids = data[id_col].dropna().unique()

    total = len(unique_ids)

    covered = sum(
        1 for i in unique_ids
        if compute_coverage(product_map, i, our_company_lower)[0]
    )

    return covered, total


# ============================================================
# BUILD ACTIVE VIEW
# ============================================================

coverage_debug_rows = None

if dashboard_view == "📅 Reference Timeline":

    st.header("📅 Crop Timeline")

    st.caption(
        "Weeds, insects, diseases and fertilizer across crop growth "
        "stages — each row is one item, grouped by category."
    )

    all_groups = build_reference_groups()

    show_legend = False

else:

    st.header("🎯 Product Coverage")

    st.caption(
        "Green = we have a registered product for this threat. "
        "Red = we don't (either no product exists, or only "
        "competitors do)."
    )

    our_company_norm = normalize_text(our_company_raw)

    our_company_lower = our_company_norm.lower()

    cov_cols = st.columns(3)

    for col, (label, icon, data, id_col, pmap) in zip(
        cov_cols,
        [
            ("Weeds", "🌱", crop_weeds, "weed_id", weed_product_map),
            ("Insects", "🐛", crop_pests, "pest_id", pest_product_map),
            ("Diseases", "🍄", crop_diseases, "disease_id", disease_product_map)
        ]
    ):
        covered, total = category_coverage_summary(
            data, id_col, pmap, our_company_lower
        )
        pct = f"{covered}/{total}" if total else "—"
        col.metric(f"{icon} {label} covered", pct)

    st.divider()

    all_groups, coverage_debug_rows = build_coverage_groups(our_company_lower)

    show_legend = True


active_groups = [
    g for g in all_groups
    if g["option_label"] in selected_categories
]

if not active_groups:
    st.info("Select at least one category above to see the chart.")
    st.stop()

fig, any_data = build_gantt_figure(active_groups, crop_timeline, show_legend=show_legend)

if not any_data:
    st.info(f"No data recorded for {selected_crop}.")
else:
    st.plotly_chart(fig, width='stretch')

if dashboard_view == "📅 Reference Timeline":
    st.caption(
        "Hover over a bar to view the Thai name, growth stage, "
        "active period, and registered products."
    )
else:
    st.caption(
        "Hover over a bar to see which companies (if any) have a "
        "registered product for that threat."
    )

st.divider()


# ============================================================
# COVERAGE DEBUG (only in Product Coverage view)
# Shows exactly what company text was found per item, so a
# "should be green but shows red" case can be diagnosed
# directly instead of guessed at.
# ============================================================

if dashboard_view == "🎯 Product Coverage" and coverage_debug_rows:

    with st.expander("🔍 Coverage debug — see exactly what was matched", expanded=False):

        st.caption(
            f"Comparing against: \"{our_company_raw}\" "
            f"(normalized to: \"{our_company_norm}\")"
        )

        debug_df = (
            pd.DataFrame(coverage_debug_rows)
            .drop_duplicates(subset=["Category", "Item ID"])
            .sort_values(["Category", "Item"])
        )

        st.dataframe(debug_df, width='stretch', hide_index=True)

    st.divider()


# ============================================================
# PRODUCTS TABLE (only in Reference Timeline view)
# ============================================================

if dashboard_view == "📅 Reference Timeline":

    st.subheader("🧪 Registered Products")

    table_category = st.segmented_control(
        "Product category",
        options=["🐛 Insects", "🍄 Diseases", "🌱 Weeds", "🧪 Fertilizer"],
        default="🐛 Insects",
        label_visibility="collapsed"
    )

    reference_groups = build_reference_groups()

    table_group_map = {
        "🐛 Insects": reference_groups[1],
        "🍄 Diseases": reference_groups[2],
        "🌱 Weeds": reference_groups[0],
        "🧪 Fertilizer": reference_groups[3]
    }

    selected_group = table_group_map[table_category]

    chart_data = selected_group["data"]
    name_column = selected_group["name_col"]
    thai_column = selected_group["thai_col"]
    id_column = selected_group["id_col"]
    category_name = selected_group["label"]

    if selected_group["mode"] == "direct":

        if chart_data.empty:
            st.info(f"No fertilizer data recorded for {selected_crop}.")
        else:

            rename_map = {
                "stage": "Stage",
                "fertilizer_formula": "Formula",
                "fertilizer_brand": "Brand",
                "fertilizer_company": "Company"
            }

            display_cols = [
                col for col in
                ["stage", "fertilizer_formula", "fertilizer_brand", "fertilizer_company"]
                if col in chart_data.columns
            ]

            st.dataframe(
                chart_data[display_cols].drop_duplicates().rename(columns=rename_map),
                width='stretch',
                hide_index=True
            )

    else:

        product_df_map = {
            "Weed": weed_her,
            "Insect": pest_ins,
            "Disease": disease_fun
        }

        product_df = product_df_map[category_name]

        if chart_data.empty:
            st.info(f"No {category_name.lower()} data recorded for {selected_crop}.")
        elif product_df.empty or id_column not in product_df.columns:
            st.info(f"No product data available for {category_name.lower()}s.")
        else:

            item_lookup = chart_data[[id_column, name_column, thai_column]].drop_duplicates()

            product_table = item_lookup.merge(product_df, on=id_column, how="inner")

            if product_table.empty:
                st.info(
                    f"No registered products found for "
                    f"{selected_crop} {category_name.lower()}s."
                )
            else:

                rename_map = {
                    name_column: f"{category_name} (EN)",
                    thai_column: f"{category_name} (TH)",
                    "brand_name": "Brand",
                    "company_name": "Company",
                    "common_name": "Common Name",
                    "concentration": "Concentration",
                    "formulation_type": "Formulation"
                }

                display_cols = [
                    col for col in [name_column, thai_column] + PRODUCT_COLUMNS
                    if col in product_table.columns
                ]

                st.dataframe(
                    product_table[display_cols].rename(columns=rename_map),
                    width='stretch',
                    hide_index=True
                )

    st.divider()
