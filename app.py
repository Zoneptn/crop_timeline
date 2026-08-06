import streamlit as st
import pandas as pd
import plotly.express as px
from pathlib import Path


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="Crop Timeline Dashboard",
    page_icon="🌱",
    layout="wide"
)


# ============================================================
# FILE PATH
# ============================================================

# Change this if your Excel file is inside another folder.
# Example:
# FILE_PATH = Path("data/crop_timeline.xlsx")

FILE_PATH = Path("crop_timeline.xlsx")


# ============================================================
# LOAD DATA
# ============================================================

@st.cache_data
def load_data(file_path):

    if not file_path.exists():
        st.error(f"File not found: {file_path}")
        st.stop()

    excel_file = pd.ExcelFile(file_path)

    # --------------------------------------------------------
    # IMPORTANT
    # --------------------------------------------------------
    # This assumes the sheets are in this order:
    #
    # Sheet 1 = crop timeline
    # Sheet 2 = crop pest
    # Sheet 3 = crop weeds
    # Sheet 4 = crop disease
    #
    # Using sheet index avoids problems if sheet names differ
    # slightly from what we expect.
    # --------------------------------------------------------

    timeline = pd.read_excel(excel_file, sheet_name=0)
    pests = pd.read_excel(excel_file, sheet_name=1)
    weeds = pd.read_excel(excel_file, sheet_name=2)
    diseases = pd.read_excel(excel_file, sheet_name=3)

    # Clean column names
    for df in [timeline, pests, weeds, diseases]:
        df.columns = (
            df.columns
            .astype(str)
            .str.strip()
            .str.lower()
            .str.replace(" ", "_")
        )

    # Clean important text columns
    for df in [timeline, pests, weeds, diseases]:

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

    # Convert timeline days to numeric
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
    "Interactive crop development timeline with pests, weeds and diseases."
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

crop_rows = timeline[
    timeline["crop"] == selected_crop
]

crop_id = crop_rows["crop_id"].iloc[0]


# ============================================================
# FILTER DATA
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


# ============================================================
# SORT TIMELINE
# ============================================================

crop_timeline = crop_timeline.sort_values(
    "start_day"
)


# ============================================================
# BASIC CROP INFORMATION
# ============================================================

st.header(f"🌿 {selected_crop.title()}")

total_days = crop_timeline["end_day"].max()
total_stages = crop_timeline["stage_id"].nunique()
total_pests = crop_pests["pest_id"].nunique()
total_diseases = crop_diseases["disease_id"].nunique()
total_weeds = crop_weeds["weed_id"].nunique()


col1, col2, col3, col4, col5 = st.columns(5)

col1.metric(
    "Growth Period",
    f"{total_days:,.0f} days"
)

col2.metric(
    "Growth Stages",
    total_stages
)

col3.metric(
    "Pests",
    total_pests
)

col4.metric(
    "Diseases",
    total_diseases
)

col5.metric(
    "Weeds",
    total_weeds
)


st.divider()


# ============================================================
# TIMELINE CHART
# ============================================================

st.subheader("📅 Crop Growth Timeline")


timeline_chart = crop_timeline.copy()

# Plotly timeline works better with dates.
# We create an artificial starting date.
base_date = pd.Timestamp("2026-01-01")

timeline_chart["start_date"] = (
    base_date
    + pd.to_timedelta(
        timeline_chart["start_day"],
        unit="D"
    )
)

timeline_chart["end_date"] = (
    base_date
    + pd.to_timedelta(
        timeline_chart["end_day"],
        unit="D"
    )
)


fig = px.timeline(
    timeline_chart,
    x_start="start_date",
    x_end="end_date",
    y="stage",
    color="stage",
    hover_data={
        "stage_id": True,
        "start_day": True,
        "end_day": True,
        "start_date": False,
        "end_date": False
    }
)


fig.update_yaxes(
    autorange="reversed",
    title=""
)


fig.update_xaxes(
    title="Crop Development Period",
    tickformat="Day %j"
)


fig.update_layout(
    height=max(
        400,
        len(timeline_chart) * 70
    ),
    showlegend=False,
    margin=dict(
        l=20,
        r=20,
        t=20,
        b=20
    )
)


st.plotly_chart(
    fig,
    use_container_width=True
)


st.divider()


# ============================================================
# STAGE SELECTOR
# ============================================================

st.subheader("🔎 Explore Growth Stage")


stage_options = crop_timeline[
    ["stage_id", "stage"]
].drop_duplicates()


stage_dictionary = dict(
    zip(
        stage_options["stage_id"],
        stage_options["stage"]
    )
)


selected_stage_id = st.selectbox(
    "Select Growth Stage",
    stage_options["stage_id"],
    format_func=lambda x:
        f"{x.upper()} — {stage_dictionary.get(x, x)}"
)


selected_stage = crop_timeline[
    crop_timeline["stage_id"]
    == selected_stage_id
].iloc[0]


# ============================================================
# STAGE INFORMATION
# ============================================================

st.markdown(
    f"## 🌱 {selected_stage['stage']}"
)


stage_start = selected_stage["start_day"]
stage_end = selected_stage["end_day"]

duration = stage_end - stage_start


c1, c2, c3 = st.columns(3)

c1.metric(
    "Start Day",
    f"{stage_start:,.0f}"
)

c2.metric(
    "End Day",
    f"{stage_end:,.0f}"
)

c3.metric(
    "Duration",
    f"{duration:,.0f} days"
)


st.divider()


# ============================================================
# FILTER BY SELECTED STAGE
# ============================================================

stage_pests = crop_pests[
    crop_pests["stage_id"]
    == selected_stage_id
].copy()


stage_weeds = crop_weeds[
    crop_weeds["stage_id"]
    == selected_stage_id
].copy()


stage_diseases = crop_diseases[
    crop_diseases["stage_id"]
    == selected_stage_id
].copy()


# ============================================================
# THREE CATEGORY COLUMNS
# ============================================================

pest_col, disease_col, weed_col = st.columns(3)


# ============================================================
# PESTS
# ============================================================

with pest_col:

    st.subheader(
        f"🐛 Pests ({len(stage_pests)})"
    )

    if stage_pests.empty:

        st.info(
            "No pests recorded for this stage."
        )

    else:

        for _, row in stage_pests.iterrows():

            pest_en = row.get(
                "pest_name_en",
                ""
            )

            pest_th = row.get(
                "pest_name_th",
                ""
            )

            st.markdown(
                f"""
                **{pest_en}**

                {pest_th}
                """
            )

            st.caption(
                f"ID: {row['pest_id']}"
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
            "No diseases recorded for this stage."
        )

    else:

        for _, row in stage_diseases.iterrows():

            disease_en = row.get(
                "disease_name_en",
                ""
            )

            disease_th = row.get(
                "disease_name_th",
                ""
            )

            disease_type = row.get(
                "type",
                ""
            )

            st.markdown(
                f"""
                **{disease_en}**

                {disease_th}
                """
            )

            if pd.notna(disease_type):

                st.caption(
                    f"{disease_type} • ID: {row['disease_id']}"
                )

            else:

                st.caption(
                    f"ID: {row['disease_id']}"
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
            "No weeds recorded for this stage."
        )

    else:

        for _, row in stage_weeds.iterrows():

            weed_en = row.get(
                "weed_name_en",
                ""
            )

            weed_th = row.get(
                "weed_name_th",
                ""
            )

            st.markdown(
                f"""
                **{weed_en}**

                {weed_th}
                """
            )

            st.caption(
                f"ID: {row['weed_id']}"
            )

            st.divider()


# ============================================================
# ALL STAGES
# ============================================================

st.divider()

st.header("📋 Complete Crop Timeline")


# ============================================================
# LOOP THROUGH EVERY STAGE
# ============================================================

for _, stage_row in crop_timeline.iterrows():

    stage_id = stage_row["stage_id"]
    stage_name = stage_row["stage"]

    start_day = stage_row["start_day"]
    end_day = stage_row["end_day"]

    # --------------------------------------------------------
    # Filter pest / disease / weed
    # --------------------------------------------------------

    stage_pest_data = crop_pests[
        crop_pests["stage_id"] == stage_id
    ]

    stage_disease_data = crop_diseases[
        crop_diseases["stage_id"] == stage_id
    ]

    stage_weed_data = crop_weeds[
        crop_weeds["stage_id"] == stage_id
    ]


    # --------------------------------------------------------
    # Expander
    # --------------------------------------------------------

    with st.expander(
        f"{stage_id.upper()} | "
        f"{stage_name} | "
        f"Day {start_day:,.0f}–{end_day:,.0f}"
    ):

        col1, col2, col3 = st.columns(3)


        # ====================================================
        # PEST
        # ====================================================

        with col1:

            st.markdown("### 🐛 Pests")

            if stage_pest_data.empty:

                st.caption(
                    "No pest recorded"
                )

            else:

                for _, pest in stage_pest_data.iterrows():

                    st.markdown(
                        f"""
                        **{pest['pest_name_en']}**

                        {pest['pest_name_th']}
                        """
                    )


        # ====================================================
        # DISEASE
        # ====================================================

        with col2:

            st.markdown(
                "### 🍄 Diseases"
            )

            if stage_disease_data.empty:

                st.caption(
                    "No disease recorded"
                )

            else:

                for _, disease in stage_disease_data.iterrows():

                    disease_type = disease.get(
                        "type",
                        ""
                    )

                    st.markdown(
                        f"""
                        **{disease['disease_name_en']}**

                        {disease['disease_name_th']}
                        """
                    )

                    if pd.notna(disease_type):

                        st.caption(
                            disease_type
                        )


        # ====================================================
        # WEED
        # ====================================================

        with col3:

            st.markdown(
                "### 🌱 Weeds"
            )

            if stage_weed_data.empty:

                st.caption(
                    "No weed recorded"
                )

            else:

                for _, weed in stage_weed_data.iterrows():

                    st.markdown(
                        f"""
                        **{weed['weed_name_en']}**

                        {weed['weed_name_th']}
                        """
                    )


# ============================================================
# RAW DATA
# ============================================================

st.divider()

with st.expander(
    "📊 View Raw Data"
):

    tab1, tab2, tab3, tab4 = st.tabs(
        [
            "Timeline",
            "Pests",
            "Diseases",
            "Weeds"
        ]
    )

    with tab1:

        st.dataframe(
            crop_timeline,
            use_container_width=True,
            hide_index=True
        )

    with tab2:

        st.dataframe(
            crop_pests,
            use_container_width=True,
            hide_index=True
        )

    with tab3:

        st.dataframe(
            crop_diseases,
            use_container_width=True,
            hide_index=True
        )

    with tab4:

        st.dataframe(
            crop_weeds,
            use_container_width=True,
            hide_index=True
        )