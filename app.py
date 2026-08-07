import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from pathlib import Path


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
if st.sidebar.button("🔄 Reload data"):
    load_data.clear()
    st.rerun()


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
# MAIN TIMELINE
# ============================================================

st.header(" Crop Timeline")

st.caption(
    "Weeds, insects, and diseases across crop growth stages — "
    "each row is one item, grouped by category."
)


# ============================================================
# PRODUCT LOOKUPS
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


weed_product_map = build_product_map(weed_her, "weed_id")

pest_product_map = build_product_map(pest_ins, "pest_id")

disease_product_map = build_product_map(disease_fun, "disease_id")


# ============================================================
# CATEGORY GROUPS (order = Weeds, Insects, Diseases)
# ============================================================

CATEGORY_GROUPS = [
    {
        "label": "Weed",
        "icon": "🌱",
        "data": crop_weeds,
        "name_col": "weed_name_en",
        "thai_col": "weed_name_th",
        "id_col": "weed_id",
        "product_map": weed_product_map,
        "color": "#2ca02c",
        "mode": "product_lookup"
    },
    {
        "label": "Insect",
        "icon": "🐛",
        "data": crop_pests,
        "name_col": "pest_name_en",
        "thai_col": "pest_name_th",
        "id_col": "pest_id",
        "product_map": pest_product_map,
        "color": "#d62728",
        "mode": "product_lookup"
    },
    {
        "label": "Disease",
        "icon": "🍄",
        "data": crop_diseases,
        "name_col": "disease_name_en",
        "thai_col": "disease_name_th",
        "id_col": "disease_id",
        "product_map": disease_product_map,
        "color": "#9467bd",
        "mode": "product_lookup"
    },
    {
        "label": "Fertilizer",
        "icon": "🧪",
        "data": crop_fert,
        "name_col": "fertilizer_formula",
        "thai_col": None,
        "id_col": "fertilizer_id",
        "product_map": None,
        "color": "#1f77b4",
        "mode": "direct"
    }
]


# ============================================================
# ROW SIZING (each group gets height proportional to its
# item count, with a minimum so empty groups don't vanish)
# ============================================================

for group in CATEGORY_GROUPS:

    if group["data"].empty or group["name_col"] not in group["data"].columns:
        group["item_count"] = 0
    else:
        group["item_count"] = group["data"][group["name_col"]].nunique()


row_heights = [
    max(g["item_count"], 1)
    for g in CATEGORY_GROUPS
]

subplot_titles = [
    f"{g['icon']} {g['label']}s ({g['item_count']})"
    for g in CATEGORY_GROUPS
]


# ============================================================
# BUILD UNIFIED FIGURE
# ============================================================

fig = make_subplots(
    rows=4,
    cols=1,
    shared_xaxes=True,
    row_heights=row_heights,
    vertical_spacing=0.06,
    subplot_titles=subplot_titles
)


any_data = False

for row_idx, group in enumerate(CATEGORY_GROUPS, start=1):

    chart_data = group["data"]

    name_column = group["name_col"]

    thai_column = group["thai_col"]

    id_column = group["id_col"]

    product_map = group["product_map"]

    color = group["color"]

    if chart_data.empty or name_column not in chart_data.columns:
        continue

    unique_items = (
        chart_data[name_column]
        .dropna()
        .unique()
    )

    for name in unique_items:

        item_data = chart_data[
            chart_data[name_column] == name
        ].sort_values("start_day")

        for _, row in item_data.iterrows():

            start = row["start_day"]
            end = row["end_day"]

            if pd.isna(start) or pd.isna(end):
                continue

            duration = end - start

            if group["mode"] == "direct":

                # Fertilizer: details live directly on this
                # row, no separate product-lookup table.
                hover_text = (
                    f"<b>🧪 {name}</b><br>"
                    f"Brand: {row.get('fertilizer_brand', '—')}<br>"
                    f"Company: {row.get('fertilizer_company', '—')}<br><br>"
                    f"Stage: {row['stage']}<br>"
                    f"Day {start:,.0f} – {end:,.0f}"
                )

            else:

                thai_name = row.get(thai_column, "")

                product_text = format_products(
                    product_map,
                    row.get(id_column)
                )

                hover_text = (
                    f"<b>{name}</b><br>"
                    f"{thai_name}<br><br>"
                    f"Stage: {row['stage']}<br>"
                    f"Day {start:,.0f} – {end:,.0f}<br><br>"
                    f"<b>Products:</b><br>{product_text}"
                )

            fig.add_trace(
                go.Bar(
                    x=[duration],
                    y=[name],
                    base=[start],
                    orientation="h",
                    marker_color=color,
                    hovertemplate=(
                        hover_text
                        + "<extra></extra>"
                    ),
                    showlegend=False
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
# STAGE BOUNDARIES (span all 4 groups)
# ============================================================

for _, row in crop_timeline.iterrows():

    start = row["start_day"]

    fig.add_vline(
        x=start,
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


# ============================================================
# STAGE LABELS (top AND bottom of the whole figure)
# ============================================================

for _, row in crop_timeline.iterrows():

    start = row["start_day"]
    end = row["end_day"]

    midpoint = (start + end) / 2

    stage_text = (
        f"<b>{row['stage']}</b>"
        f"<br>"
        f"Day {start:,.0f}–{end:,.0f}"
    )

    # Bottom label — below the last (Fertilizer) row
    fig.add_annotation(
        x=midpoint,
        y=-0.2,
        xref="x4",
        yref="paper",
        text=stage_text,
        showarrow=False,
        align="center",
        font=dict(size=11)
    )

    # Top label — above the first (Weeds) row, mirrored
    fig.add_annotation(
        x=midpoint,
        y=1.2,
        xref="x",
        yref="paper",
        text=stage_text,
        showarrow=False,
        align="center",
        font=dict(size=11)
    )


# ============================================================
# CHART LAYOUT
# ============================================================

total_items = sum(row_heights)

chart_height = max(
    700,
    total_items * 38 + 320
)

fig.update_layout(

    barmode="overlay",

    height=chart_height,

    margin=dict(
        l=20,
        r=20,
        t=250,
        b=250
    ),

    hovermode="closest",

    showlegend=False
)

fig.update_xaxes(
    title="Days After Planting",
    side="bottom",
    showgrid=True,
    zeroline=False,
    row=4,
    col=1
)

fig.update_xaxes(
    showgrid=True,
    zeroline=False,
    row=1,
    col=1
)

fig.update_xaxes(
    showgrid=True,
    zeroline=False,
    row=2,
    col=1
)

fig.update_xaxes(
    showgrid=True,
    zeroline=False,
    row=3,
    col=1
)


# ============================================================
# DISPLAY
# ============================================================

if not any_data:

    st.info(
        f"No weed, insect, disease, or fertilizer data recorded for {selected_crop}."
    )

else:

    st.plotly_chart(
        fig,
        width='stretch'
    )


st.caption(
    "Hover over a bar to view the Thai name, growth stage, "
    "active period, and registered products."
)


st.divider()


# ============================================================
# PRODUCTS TABLE
# ============================================================

st.subheader("🧪 Registered Products")

table_category = st.segmented_control(
    "Product category",
    options=["🐛 Insects", "🍄 Diseases", "🌱 Weeds", "🧪 Fertilizer"],
    default="🐛 Insects",
    label_visibility="collapsed"
)

table_group_map = {
    "🐛 Insects": CATEGORY_GROUPS[1],
    "🍄 Diseases": CATEGORY_GROUPS[2],
    "🌱 Weeds": CATEGORY_GROUPS[0],
    "🧪 Fertilizer": CATEGORY_GROUPS[3]
}

selected_group = table_group_map[table_category]

chart_data = selected_group["data"]
name_column = selected_group["name_col"]
thai_column = selected_group["thai_col"]
id_column = selected_group["id_col"]
category_name = selected_group["label"]

if selected_group["mode"] == "direct":

    # Fertilizer: the sheet already IS the detail table,
    # no separate product sheet to merge in.

    if chart_data.empty:

        st.info(
            f"No fertilizer data recorded for {selected_crop}."
        )

    else:

        rename_map = {
            "stage": "Stage",
            "fertilizer_formula": "Formula",
            "fertilizer_brand": "Brand",
            "fertilizer_company": "Company"
        }

        display_cols = [
            col for col in
            [
                "stage",
                "fertilizer_formula",
                "fertilizer_brand",
                "fertilizer_company"
            ]
            if col in chart_data.columns
        ]

        st.dataframe(
            chart_data[display_cols]
            .drop_duplicates()
            .rename(columns=rename_map),
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

        st.info(
            f"No {category_name.lower()} data recorded "
            f"for {selected_crop}."
        )

    elif product_df.empty or id_column not in product_df.columns:

        st.info(
            f"No product data available for {category_name.lower()}s."
        )

    else:

        item_lookup = chart_data[
            [id_column, name_column, thai_column]
        ].drop_duplicates()

        product_table = item_lookup.merge(
            product_df,
            on=id_column,
            how="inner"
        )

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
                col for col in
                [name_column, thai_column] + PRODUCT_COLUMNS
                if col in product_table.columns
            ]

            st.dataframe(
                product_table[display_cols].rename(
                    columns=rename_map
                ),
                width='stretch',
                hide_index=True
            )


st.divider()
