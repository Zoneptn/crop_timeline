import streamlit as st
import pandas as pd
import plotly.graph_objects as go
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

@st.cache_data
def load_data(file_path):

    if not file_path.exists():
        st.error(f"File not found: {file_path}")
        st.stop()

    excel = pd.ExcelFile(file_path)

    timeline = pd.read_excel(excel, sheet_name=0)
    pests = pd.read_excel(excel, sheet_name=1)
    weeds = pd.read_excel(excel, sheet_name=2)
    diseases = pd.read_excel(excel, sheet_name=3)

    # -----------------------------------------
    # Product link sheets (looked up by name,
    # since they may not always be in the same
    # position in the workbook)
    # -----------------------------------------

    sheet_lookup = {
        name.strip().lower(): name
        for name in excel.sheet_names
    }

    def load_sheet_by_name(target_name):

        key = target_name.strip().lower()

        if key in sheet_lookup:
            return pd.read_excel(
                excel,
                sheet_name=sheet_lookup[key]
            )

        st.warning(
            f"Sheet '{target_name}' not found in the workbook. "
            f"Product info for this category will be unavailable."
        )

        return pd.DataFrame()

    weed_her = load_sheet_by_name("weed_her")
    pest_ins = load_sheet_by_name("pest_ins")
    disease_fun = load_sheet_by_name("disease_fun")

    datasets = [
        timeline,
        pests,
        weeds,
        diseases,
        weed_her,
        pest_ins,
        disease_fun
    ]

    # Clean column names
    for df in datasets:

        df.columns = (
            df.columns
            .astype(str)
            .str.strip()
            .str.lower()
            .str.replace(" ", "_")
        )

        for id_col in (
            "crop_id",
            "stage_id",
            "pest_id",
            "weed_id",
            "disease_id"
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
        disease_fun
    )


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


crop_timeline = crop_timeline.sort_values(
    "start_day"
)


# ============================================================
# MERGE TIMELINE DAYS INTO PEST / WEED / DISEASE
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


# ============================================================
# SUMMARY
# ============================================================

st.header(f"🌿 {selected_crop.title()}")


total_days = crop_timeline["end_day"].max()

total_stages = crop_timeline[
    "stage_id"
].nunique()

total_pests = crop_pests[
    "pest_id"
].nunique()

total_diseases = crop_diseases[
    "disease_id"
].nunique()

total_weeds = crop_weeds[
    "weed_id"
].nunique()


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

st.header("📅 Crop Timeline")

st.caption(
    "Select a category to view its occurrence across crop growth stages."
)


# ============================================================
# CATEGORY SELECTOR
# ============================================================

timeline_category = st.segmented_control(
    "Timeline category",
    options=[
        "🐛 Insects",
        "🍄 Diseases",
        "🌱 Weeds"
    ],
    default="🐛 Insects",
    label_visibility="collapsed"
)


# ============================================================
# SELECT DATA BASED ON CATEGORY
# ============================================================

if timeline_category == "🐛 Insects":

    chart_data = crop_pests.copy()

    name_column = "pest_name_en"

    thai_column = "pest_name_th"

    category_name = "Insect"

    id_column = "pest_id"

    product_df = pest_ins


elif timeline_category == "🍄 Diseases":

    chart_data = crop_diseases.copy()

    name_column = "disease_name_en"

    thai_column = "disease_name_th"

    category_name = "Disease"

    id_column = "disease_id"

    product_df = disease_fun


else:

    chart_data = crop_weeds.copy()

    name_column = "weed_name_en"

    thai_column = "weed_name_th"

    category_name = "Weed"

    id_column = "weed_id"

    product_df = weed_her


# ============================================================
# PRODUCT LOOKUP (for hover text)
# ============================================================

PRODUCT_COLUMNS = [
    "brand_name",
    "company_name",
    "common_name",
    "concentration",
    "formulation_type"
]

if not product_df.empty and id_column in product_df.columns:

    product_map = (
        product_df
        .groupby(id_column)
        .apply(lambda d: d.to_dict("records"))
        .to_dict()
    )

else:

    product_map = {}


def format_products(product_id):

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


# ============================================================
# CREATE TIMELINE
# ============================================================

fig = go.Figure()


# ============================================================
# ADD ITEMS
# ============================================================

if not chart_data.empty:

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

            # Thai name
            thai_name = row.get(
                thai_column,
                ""
            )

            # Registered products for this item
            product_text = format_products(
                row.get(id_column)
            )

            # Hover information
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

                    hovertemplate=(
                        hover_text
                        + "<extra></extra>"
                    ),

                    showlegend=False
                )
            )


# ============================================================
# ADD STAGE BOUNDARIES
# ============================================================

for _, row in crop_timeline.iterrows():

    start = row["start_day"]
    end = row["end_day"]

    # -----------------------------------------
    # Vertical line at beginning of stage
    # -----------------------------------------

    fig.add_vline(
        x=start,
        line_width=1,
        line_dash="dot"
    )


    # -----------------------------------------
    # Stage label position
    # -----------------------------------------

    midpoint = (
        start + end
    ) / 2


    # -----------------------------------------
    # Stage name above chart
    # -----------------------------------------

    fig.add_annotation(

        x=midpoint,

        y=-0.22,

        xref="x",

        yref="paper",

        text=(
            f"<b>{row['stage']}</b>"
            f"<br>"
            f"Day {start:,.0f}–{end:,.0f}"
        ),

        showarrow=False,

        align="center",

        font=dict(
            size=11
        )
    )


# ============================================================
# FINAL STAGE BOUNDARY
# ============================================================

if not crop_timeline.empty:

    fig.add_vline(

        x=crop_timeline[
            "end_day"
        ].max(),

        line_width=1,

        line_dash="dot"
    )


# ============================================================
# CHART HEIGHT
# ============================================================

number_items = chart_data[
    name_column
].nunique()


chart_height = max(
    450,
    number_items * 42
)


# ============================================================
# CHART LAYOUT
# ============================================================

fig.update_layout(

    barmode="overlay",

    height=chart_height,

    xaxis=dict(

        title="Days After Planting",

        side="bottom",

        showgrid=True,

        zeroline=False
    ),

    yaxis=dict(

        title="",

        autorange="reversed",

        automargin=True
    ),

    margin=dict(

        l=20,

        r=20,

        t=30,

        b=180
    ),

    hovermode="closest",

    showlegend=False
)


# ============================================================
# DISPLAY
# ============================================================

if chart_data.empty:

    st.info(
        f"No {category_name.lower()} data "
        f"recorded for {selected_crop}."
    )

else:

    st.plotly_chart(
        fig,
        use_container_width=True
    )


st.caption(
    "Hover over a bar to view the Thai name, "
    "growth stage and active period."
)


st.divider()


# ============================================================
# PRODUCTS TABLE
# ============================================================

st.subheader(f"🧪 Registered Products — {category_name}s")

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
            use_container_width=True,
            hide_index=True
        )


st.divider()
