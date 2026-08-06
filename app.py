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


elif timeline_category == "🍄 Diseases":

    chart_data = crop_diseases.copy()

    name_column = "disease_name_en"

    thai_column = "disease_name_th"

    category_name = "Disease"


else:

    chart_data = crop_weeds.copy()

    name_column = "weed_name_en"

    thai_column = "weed_name_th"

    category_name = "Weed"


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

            # Hover information
            hover_text = (
                f"<b>{name}</b><br>"
                f"{thai_name}<br><br>"
                f"Stage: {row['stage']}<br>"
                f"Day {start:,.0f} – {end:,.0f}"
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

        side="top",

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

        t=120,

        b=30
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


elif timeline_category == "🍄 Diseases":

    chart_data = crop_diseases.copy()

    name_column = "disease_name_en"

    thai_column = "disease_name_th"

    category_name = "Disease"


else:

    chart_data = crop_weeds.copy()

    name_column = "weed_name_en"

    thai_column = "weed_name_th"

    category_name = "Weed"


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

            # Hover information
            hover_text = (
                f"<b>{name}</b><br>"
                f"{thai_name}<br><br>"
                f"Stage: {row['stage']}<br>"
                f"Day {start:,.0f} – {end:,.0f}"
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

        side="top",

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

        t=120,

        b=30
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
