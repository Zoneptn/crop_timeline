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

    datasets = [timeline, pests, weeds, diseases]

    # Clean column names
    for df in datasets:

        df.columns = (
            df.columns
            .astype(str)
            .str.strip()
            .str.lower()
            .str.replace(" ", "_")
        )

        if "crop_id" in df.columns:
            df["crop_id"] = (
                df["crop_id"]
                .astype(str)
                .str.strip()
                .str.lower()
            )

        if "stage_id" in df.columns:
            df["stage_id"] = (
                df["stage_id"]
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

    return timeline, pests, weeds, diseases


timeline, pests, weeds, diseases = load_data(FILE_PATH)


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


fig = go.Figure()


# ============================================================
# FUNCTION TO ADD TIMELINE BARS
# ============================================================

def add_timeline_rows(
    df,
    name_column,
    category,
    symbol
):

    if df.empty:
        return

    # --------------------------------------------------------
    # A pest can appear in several stages.
    #
    # Example:
    # Thrips S03
    # Thrips S04
    #
    # Each stage is displayed as part of the same row.
    # --------------------------------------------------------

    for name in df[name_column].dropna().unique():

        item_data = df[
            df[name_column] == name
        ].sort_values("start_day")

        for _, row in item_data.iterrows():

            start = row["start_day"]
            end = row["end_day"]

            if pd.isna(start) or pd.isna(end):
                continue

            duration = end - start

            hover_text = (
                f"<b>{name}</b><br>"
                f"Category: {category}<br>"
                f"Stage: {row['stage']}<br>"
                f"Day {start:,.0f} – {end:,.0f}"
            )

            fig.add_trace(
                go.Bar(
                    x=[duration],
                    y=[f"{symbol} {name}"],
                    base=[start],
                    orientation="h",
                    name=category,
                    legendgroup=category,
                    showlegend=False,
                    hovertemplate=(
                        hover_text
                        + "<extra></extra>"
                    )
                )
            )


# ============================================================
# ADD INSECTS
# ============================================================

add_timeline_rows(
    crop_pests,
    "pest_name_en",
    "Insect",
    "🐛"
)


# ============================================================
# ADD DISEASES
# ============================================================

add_timeline_rows(
    crop_diseases,
    "disease_name_en",
    "Disease",
    "🍄"
)


# ============================================================
# ADD WEEDS
# ============================================================

add_timeline_rows(
    crop_weeds,
    "weed_name_en",
    "Weed",
    "🌱"
)


# ============================================================
# STAGE BACKGROUND + LABELS
# ============================================================

for i, (_, row) in enumerate(
    crop_timeline.iterrows()
):

    start = row["start_day"]
    end = row["end_day"]

    # Stage boundary
    fig.add_vline(
        x=start,
        line_width=1,
        line_dash="dot"
    )

    # Center of stage
    midpoint = (
        start + end
    ) / 2

    # Stage annotation
    fig.add_annotation(
        x=midpoint,
        y=1.08,
        xref="x",
        yref="paper",
        text=(
            f"<b>{row['stage']}</b>"
            f"<br>"
            f"Day {start:,.0f}–{end:,.0f}"
        ),
        showarrow=False,
        align="center",
        font=dict(size=11)
    )


# Last boundary
if not crop_timeline.empty:

    fig.add_vline(
        x=crop_timeline[
            "end_day"
        ].max(),
        line_width=1,
        line_dash="dot"
    )


# ============================================================
# CHART LAYOUT
# ============================================================

number_rows = (
    crop_pests["pest_name_en"].nunique()
    + crop_diseases["disease_name_en"].nunique()
    + crop_weeds["weed_name_en"].nunique()
)


fig.update_layout(

    barmode="overlay",

    height=max(
        550,
        number_rows * 38
    ),

    xaxis=dict(
        title="Days After Planting",
        side="top",
        showgrid=True,
        zeroline=False
    ),

    yaxis=dict(
        title="",
        autorange="reversed"
    ),

    margin=dict(
        l=20,
        r=20,
        t=120,
        b=30
    ),

    hovermode="closest",

    showlegend=False
)


st.plotly_chart(
    fig,
    use_container_width=True
)


st.caption(
    "Hover over a bar to see the growth stage and active period."
)


st.divider()


# ============================================================
# EXPLORE STAGE
# ============================================================

st.header("🔎 Explore Growth Stage")


stage_options = crop_timeline[
    [
        "stage_id",
        "stage",
        "start_day",
        "end_day"
    ]
].drop_duplicates()


stage_dict = dict(
    zip(
        stage_options["stage_id"],
        stage_options["stage"]
    )
)


selected_stage_id = st.selectbox(

    "Select Growth Stage",

    stage_options["stage_id"],

    format_func=lambda x:
        stage_dict.get(x, x)
)


selected_stage = stage_options[
    stage_options["stage_id"]
    == selected_stage_id
].iloc[0]


# ============================================================
# STAGE HEADER
# ============================================================

st.subheader(
    f"🌱 {selected_stage['stage']}"
)


start_day = selected_stage[
    "start_day"
]

end_day = selected_stage[
    "end_day"
]


c1, c2, c3 = st.columns(3)


c1.metric(
    "Start Day",
    f"{start_day:,.0f}"
)


c2.metric(
    "End Day",
    f"{end_day:,.0f}"
)


c3.metric(
    "Duration",
    f"{end_day - start_day:,.0f} days"
)


st.divider()


# ============================================================
# FILTER SELECTED STAGE
# ============================================================

stage_pests = crop_pests[
    crop_pests["stage_id"]
    == selected_stage_id
]


stage_diseases = crop_diseases[
    crop_diseases["stage_id"]
    == selected_stage_id
]


stage_weeds = crop_weeds[
    crop_weeds["stage_id"]
    == selected_stage_id
]


# ============================================================
# EXPLORE COLUMNS
# ============================================================

pest_col, disease_col, weed_col = st.columns(3)


# ============================================================
# INSECTS
# ============================================================

with pest_col:

    st.subheader(
        f"🐛 Insects ({len(stage_pests)})"
    )

    if stage_pests.empty:

        st.info(
            "No insect pests recorded."
        )

    else:

        display_pests = (
            stage_pests[
                [
                    "pest_name_en",
                    "pest_name_th"
                ]
            ]
            .drop_duplicates()
            .sort_values("pest_name_en")
        )

        for _, row in display_pests.iterrows():

            st.markdown(
                f"""
                **{row['pest_name_en']}**

                {row['pest_name_th']}
                """
            )

            st.divider()


# ============================================================
# DISEASES
# ============================================================

with disease_col:

    st.subheader(
        f"🍄 Diseases ({len(stage_diseases)})"
    )

    if stage_diseases.empty:

        st.info(
            "No diseases recorded."
        )

    else:

        disease_columns = [
            "disease_name_en",
            "disease_name_th"
        ]

        if "type" in stage_diseases.columns:
            disease_columns.append("type")

        display_diseases = (
            stage_diseases[
                disease_columns
            ]
            .drop_duplicates()
            .sort_values(
                "disease_name_en"
            )
        )

        for _, row in display_diseases.iterrows():

            st.markdown(
                f"""
                **{row['disease_name_en']}**

                {row['disease_name_th']}
                """
            )

            if (
                "type" in row.index
                and pd.notna(row["type"])
            ):

                st.caption(
                    row["type"]
                )

            st.divider()


# ============================================================
# WEEDS
# ============================================================

with weed_col:

    st.subheader(
        f"🌱 Weeds ({len(stage_weeds)})"
    )

    if stage_weeds.empty:

        st.info(
            "No weeds recorded."
        )

    else:

        display_weeds = (
            stage_weeds[
                [
                    "weed_name_en",
                    "weed_name_th"
                ]
            ]
            .drop_duplicates()
            .sort_values("weed_name_en")
        )

        for _, row in display_weeds.iterrows():

            st.markdown(
                f"""
                **{row['weed_name_en']}**

                {row['weed_name_th']}
                """
            )

            st.divider()


# ============================================================
# COMPLETE STAGE OVERVIEW
# ============================================================

st.divider()

st.header("📋 Stage Overview")


for _, stage_row in crop_timeline.iterrows():

    sid = stage_row["stage_id"]

    stage_pest_data = crop_pests[
        crop_pests["stage_id"] == sid
    ]

    stage_disease_data = crop_diseases[
        crop_diseases["stage_id"] == sid
    ]

    stage_weed_data = crop_weeds[
        crop_weeds["stage_id"] == sid
    ]


    with st.expander(

        f"{stage_row['stage']}  |  "
        f"Day {stage_row['start_day']:,.0f}"
        f"–"
        f"{stage_row['end_day']:,.0f}"

    ):

        col1, col2, col3 = st.columns(3)


        # ----------------------------------------------------
        # INSECTS
        # ----------------------------------------------------

        with col1:

            st.markdown("### 🐛 Insects")

            if stage_pest_data.empty:

                st.caption(
                    "No insect pests recorded"
                )

            else:

                temp = stage_pest_data[
                    [
                        "pest_name_en",
                        "pest_name_th"
                    ]
                ].drop_duplicates()

                for _, row in temp.iterrows():

                    st.markdown(
                        f"""
                        **{row['pest_name_en']}**  
                        {row['pest_name_th']}
                        """
                    )


        # ----------------------------------------------------
        # DISEASES
        # ----------------------------------------------------

        with col2:

            st.markdown("### 🍄 Diseases")

            if stage_disease_data.empty:

                st.caption(
                    "No diseases recorded"
                )

            else:

                cols = [
                    "disease_name_en",
                    "disease_name_th"
                ]

                if "type" in stage_disease_data.columns:
                    cols.append("type")

                temp = stage_disease_data[
                    cols
                ].drop_duplicates()

                for _, row in temp.iterrows():

                    st.markdown(
                        f"""
                        **{row['disease_name_en']}**  
                        {row['disease_name_th']}
                        """
                    )

                    if (
                        "type" in row.index
                        and pd.notna(row["type"])
                    ):

                        st.caption(
                            row["type"]
                        )


        # ----------------------------------------------------
        # WEEDS
        # ----------------------------------------------------

        with col3:

            st.markdown("### 🌱 Weeds")

            if stage_weed_data.empty:

                st.caption(
                    "No weeds recorded"
                )

            else:

                temp = stage_weed_data[
                    [
                        "weed_name_en",
                        "weed_name_th"
                    ]
                ].drop_duplicates()

                for _, row in temp.iterrows():

                    st.markdown(
                        f"""
                        **{row['weed_name_en']}**  
                        {row['weed_name_th']}
                        """
                    )