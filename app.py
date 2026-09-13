import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from pathlib import Path


# =========================================================
# PAGE CONFIG
# =========================================================

st.set_page_config(
    page_title="Gardenia Town | Inventory Dashboard",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded"
)


# =========================================================
# STYLE
# =========================================================

st.markdown("""
<style>

    .main {
        background-color: #f7f8fa;
    }

    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 2rem;
    }

    h1, h2, h3 {
        color: #17324d;
    }

    .dashboard-title {
        font-size: 32px;
        font-weight: 700;
        color: #17324d;
        margin-bottom: 0;
    }

    .dashboard-subtitle {
        color: #718096;
        margin-top: 5px;
        margin-bottom: 25px;
    }

    .kpi {
        background: white;
        padding: 20px;
        border-radius: 12px;
        border: 1px solid #e6e9ed;
        box-shadow: 0 2px 8px rgba(0,0,0,0.04);
    }

    .kpi-title {
        font-size: 13px;
        color: #718096;
        margin-bottom: 8px;
    }

    .kpi-value {
        font-size: 28px;
        font-weight: 700;
        color: #17324d;
    }

    .section-title {
        font-size: 20px;
        font-weight: 700;
        color: #17324d;
        margin-top: 25px;
        margin-bottom: 10px;
    }

    div[data-testid="stDataFrame"] {
        border-radius: 10px;
    }

</style>
""", unsafe_allow_html=True)


# =========================================================
# FIND EXCEL FILE
# =========================================================

BASE_DIR = Path(__file__).resolve().parent

possible_paths = [
    BASE_DIR / "data" / "All Inventory Project(1).xlsx",
    BASE_DIR / "All Inventory Project(1).xlsx",
    BASE_DIR / "data" / "All_Inventory_Project.xlsx",
    BASE_DIR / "All_Inventory_Project.xlsx",
]

EXCEL_PATH = None

for path in possible_paths:
    if path.exists():
        EXCEL_PATH = path
        break


# =========================================================
# LOAD DATA
# =========================================================

@st.cache_data
def load_data(file_path):

    df = pd.read_excel(
        file_path,
        sheet_name="Gardenia Town",
        header=4
    )

    # Remove completely empty columns
    df = df.dropna(axis=1, how="all")

    # Clean column names
    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
        .str.replace("\n", " ", regex=False)
    )

    return df


if EXCEL_PATH is None:

    st.error(
        "Excel file was not found. "
        "Make sure 'All Inventory Project(1).xlsx' is inside the "
        "data folder or the same folder as app.py."
    )

    st.stop()


try:

    df = load_data(EXCEL_PATH)

except Exception as e:

    st.error(f"Could not load the Excel file: {e}")
    st.stop()


# =========================================================
# CHECK REQUIRED COLUMNS
# =========================================================

required_columns = [
    "Phase",
    "Unit Code",
    "Unit Type",
    "Building",
    "Floor",
    "Type",
    "In/Area",
    "STATUS"
]

missing_columns = [
    col for col in required_columns
    if col not in df.columns
]

if missing_columns:

    st.error(
        "The following required columns are missing:\n\n"
        + ", ".join(missing_columns)
    )

    st.write("Columns detected in the Excel file:")
    st.write(list(df.columns))

    st.stop()


# =========================================================
# CLEAN DATA
# =========================================================

df["Phase"] = df["Phase"].astype(str).str.strip()
df["STATUS"] = df["STATUS"].astype(str).str.strip()
df["Type"] = df["Type"].astype(str).str.strip()
df["Unit Type"] = df["Unit Type"].astype(str).str.strip()
df["Building"] = df["Building"].astype(str).str.strip()

# Numeric fields
df["Floor"] = pd.to_numeric(df["Floor"], errors="coerce")
df["In/Area"] = pd.to_numeric(df["In/Area"], errors="coerce")

# Remove invalid phase rows
df = df[
    df["Phase"].isin(["ALBA", "ORCHID"])
].copy()


# =========================================================
# HEADER
# =========================================================

st.markdown(
    '<div class="dashboard-title">Gardenia Town Inventory Dashboard</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="dashboard-subtitle">'
    'ALBA vs ORCHID — Sales & Inventory Overview'
    '</div>',
    unsafe_allow_html=True
)


# =========================================================
# SIDEBAR FILTERS
# =========================================================

st.sidebar.title("Filters")

st.sidebar.markdown("### Inventory Selection")


# -------------------------
# Phase
# -------------------------

phase_options = sorted(
    df["Phase"].dropna().unique().tolist()
)

phase_filter = st.sidebar.multiselect(
    "Phase",
    options=phase_options,
    default=phase_options
)


# -------------------------
# STATUS
# -------------------------

status_options = sorted(
    df["STATUS"].dropna().unique().tolist()
)

status_filter = st.sidebar.multiselect(
    "Status",
    options=status_options,
    default=status_options
)


# -------------------------
# TYPE
# -------------------------

type_options = sorted(
    df["Type"].dropna().unique().tolist()
)

type_filter = st.sidebar.multiselect(
    "Type",
    options=type_options,
    default=type_options
)


# -------------------------
# UNIT TYPE
# -------------------------

unit_type_options = sorted(
    df["Unit Type"].dropna().unique().tolist()
)

unit_type_filter = st.sidebar.multiselect(
    "Unit Type",
    options=unit_type_options,
    default=unit_type_options
)


# -------------------------
# BUILDING
# -------------------------

building_options = sorted(
    df["Building"].dropna().unique().tolist()
)

building_filter = st.sidebar.multiselect(
    "Building",
    options=building_options,
    default=building_options
)


# =========================================================
# FLOOR FILTER
# =========================================================

valid_floors = df["Floor"].dropna()

if len(valid_floors) > 0:

    min_floor = int(valid_floors.min())
    max_floor = int(valid_floors.max())

    floor_range = st.sidebar.slider(
        "Floor",
        min_value=min_floor,
        max_value=max_floor,
        value=(min_floor, max_floor)
    )

else:

    floor_range = None


# =========================================================
# AREA FILTER
# =========================================================

valid_area = df["In/Area"].dropna()

if len(valid_area) > 0:

    min_area = float(valid_area.min())
    max_area = float(valid_area.max())

    area_range = st.sidebar.slider(
        "Area (m²)",
        min_value=float(min_area),
        max_value=float(max_area),
        value=(float(min_area), float(max_area)),
        step=1.0
    )

else:

    area_range = None


# =========================================================
# SEARCH
# =========================================================

search_text = st.sidebar.text_input(
    "Search Unit / Client",
    placeholder="Unit code or client name..."
)


# =========================================================
# APPLY FILTERS
# =========================================================

filtered_df = df.copy()


# Phase
if phase_filter:
    filtered_df = filtered_df[
        filtered_df["Phase"].isin(phase_filter)
    ]


# Status
if status_filter:
    filtered_df = filtered_df[
        filtered_df["STATUS"].isin(status_filter)
    ]


# Type
if type_filter:
    filtered_df = filtered_df[
        filtered_df["Type"].isin(type_filter)
    ]


# Unit Type
if unit_type_filter:
    filtered_df = filtered_df[
        filtered_df["Unit Type"].isin(unit_type_filter)
    ]


# Building
if building_filter:
    filtered_df = filtered_df[
        filtered_df["Building"].isin(building_filter)
    ]


# Floor
if floor_range is not None:

    filtered_df = filtered_df[
        filtered_df["Floor"].between(
            floor_range[0],
            floor_range[1],
            inclusive="both"
        )
        | filtered_df["Floor"].isna()
    ]


# Area
if area_range is not None:

    filtered_df = filtered_df[
        filtered_df["In/Area"].between(
            area_range[0],
            area_range[1],
            inclusive="both"
        )
        | filtered_df["In/Area"].isna()
    ]


# Search
if search_text.strip():

    search = search_text.strip().lower()

    mask = (
        filtered_df["Unit Code"]
        .astype(str)
        .str.lower()
        .str.contains(search, na=False)
        |
        filtered_df.get(
            "Name of client",
            pd.Series("", index=filtered_df.index)
        )
        .astype(str)
        .str.lower()
        .str.contains(search, na=False)
    )

    filtered_df = filtered_df[mask]


# =========================================================
# RESET FILTERS
# =========================================================

if st.sidebar.button("Reset Filters", use_container_width=True):

    st.cache_data.clear()
    st.rerun()


# =========================================================
# KPI CALCULATIONS
# =========================================================

total_units = len(filtered_df)

sold_units = len(
    filtered_df[
        filtered_df["STATUS"].str.upper() == "SOLD"
    ]
)

available_units = len(
    filtered_df[
        filtered_df["STATUS"].str.upper() == "AVAILABLE"
    ]
)

hold_units = len(
    filtered_df[
        filtered_df["STATUS"].str.upper() == "HOLD"
    ]
)

reserved_units = len(
    filtered_df[
        filtered_df["STATUS"].str.upper() == "RESERVED"
    ]
)

if total_units > 0:

    conversion = (
        sold_units / total_units
    ) * 100

else:

    conversion = 0


# =========================================================
# KPI CARDS
# =========================================================

c1, c2, c3, c4, c5 = st.columns(5)


with c1:

    st.markdown(
        f"""
        <div class="kpi">
            <div class="kpi-title">TOTAL UNITS</div>
            <div class="kpi-value">
                {total_units:,}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


with c2:

    st.markdown(
        f"""
        <div class="kpi">
            <div class="kpi-title">SOLD</div>
            <div class="kpi-value">
                {sold_units:,}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


with c3:

    st.markdown(
        f"""
        <div class="kpi">
            <div class="kpi-title">AVAILABLE</div>
            <div class="kpi-value">
                {available_units:,}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


with c4:

    st.markdown(
        f"""
        <div class="kpi">
            <div class="kpi-title">HOLD / RESERVED</div>
            <div class="kpi-value">
                {hold_units + reserved_units:,}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


with c5:

    st.markdown(
        f"""
        <div class="kpi">
            <div class="kpi-title">SALES CONVERSION</div>
            <div class="kpi-value">
                {conversion:.1f}%
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


# =========================================================
# NO DATA
# =========================================================

if filtered_df.empty:

    st.warning(
        "No units match the selected filters."
    )

    st.stop()


# =========================================================
# CHART 1 — INVENTORY STATUS
# =========================================================

st.markdown(
    '<div class="section-title">Inventory Status</div>',
    unsafe_allow_html=True
)

status_counts = (
    filtered_df["STATUS"]
    .value_counts()
    .reset_index()
)

status_counts.columns = [
    "Status",
    "Units"
]

fig_status = px.bar(
    status_counts,
    x="Status",
    y="Units",
    text="Units",
    title="Units by Status"
)

fig_status.update_traces(
    textposition="outside"
)

fig_status.update_layout(
    height=400,
    margin=dict(l=20, r=20, t=60, b=20),
    xaxis_title="",
    yaxis_title="Units",
    showlegend=False
)

st.plotly_chart(
    fig_status,
    use_container_width=True
)


# =========================================================
# CHART 2 — SALES CONVERSION BY PHASE
# =========================================================

st.markdown(
    '<div class="section-title">Sales Conversion by Phase</div>',
    unsafe_allow_html=True
)

phase_summary = (
    filtered_df
    .groupby("Phase")
    .agg(
        Total=("Phase", "size"),
        Sold=("STATUS", lambda x:
              (x.str.upper() == "SOLD").sum())
    )
    .reset_index()
)

phase_summary["Conversion"] = np.where(
    phase_summary["Total"] > 0,
    phase_summary["Sold"]
    / phase_summary["Total"]
    * 100,
    0
)

fig_conversion = px.bar(
    phase_summary,
    x="Phase",
    y="Conversion",
    text=phase_summary["Conversion"].round(1).astype(str) + "%",
    title="Sales Conversion by Phase"
)

fig_conversion.update_traces(
    textposition="outside"
)

fig_conversion.update_layout(
    height=400,
    yaxis_title="Conversion %",
    xaxis_title="",
    yaxis=dict(range=[0, 100]),
    margin=dict(l=20, r=20, t=60, b=20)
)

st.plotly_chart(
    fig_conversion,
    use_container_width=True
)


# =========================================================
# TWO COLUMN SECTION
# =========================================================

col1, col2 = st.columns(2)


# =========================================================
# AVAILABLE UNITS BY BUILDING
# =========================================================

with col1:

    st.markdown(
        '<div class="section-title">'
        'Available Units by Building'
        '</div>',
        unsafe_allow_html=True
    )

    building_available = (
        filtered_df[
            filtered_df["STATUS"].str.upper() == "AVAILABLE"
        ]
        .groupby("Building")
        .size()
        .reset_index(name="Available Units")
        .sort_values(
            "Available Units",
            ascending=False
        )
        .head(20)
    )

    if not building_available.empty:

        fig_building = px.bar(
            building_available,
            x="Building",
            y="Available Units",
            text="Available Units",
            title="Top Buildings by Available Inventory"
        )

        fig_building.update_traces(
            textposition="outside"
        )

        fig_building.update_layout(
            height=450,
            xaxis_title="Building",
            yaxis_title="Available Units",
            margin=dict(
                l=20,
                r=20,
                t=60,
                b=20
            )
        )

        st.plotly_chart(
            fig_building,
            use_container_width=True
        )

    else:

        st.info("No available units in the selected filters.")


# =========================================================
# UNIT TYPE MIX
# =========================================================

with col2:

    st.markdown(
        '<div class="section-title">Unit Type Mix</div>',
        unsafe_allow_html=True
    )

    unit_type_counts = (
        filtered_df["Unit Type"]
        .value_counts()
        .reset_index()
    )

    unit_type_counts.columns = [
        "Unit Type",
        "Units"
    ]

    fig_type = px.pie(
        unit_type_counts,
        names="Unit Type",
        values="Units",
        hole=0.45,
        title="Inventory by Unit Type"
    )

    fig_type.update_traces(
        textinfo="label+percent",
        hovertemplate=(
            "<b>%{label}</b><br>"
            "Units: %{value}<br>"
            "Share: %{percent}"
            "<extra></extra>"
        )
    )

    fig_type.update_layout(
        height=450,
        margin=dict(
            l=20,
            r=20,
            t=60,
            b=20
        )
    )

    st.plotly_chart(
        fig_type,
        use_container_width=True
    )


# =========================================================
# TYPE ANALYSIS
# =========================================================

st.markdown(
    '<div class="section-title">Sales Conversion by Type</div>',
    unsafe_allow_html=True
)

type_summary = (
    filtered_df
    .groupby("Type")
    .agg(
        Total=("Type", "size"),
        Sold=("STATUS", lambda x:
              (x.str.upper() == "SOLD").sum()),
        Available=("STATUS", lambda x:
                   (x.str.upper() == "AVAILABLE").sum())
    )
    .reset_index()
)

type_summary["Conversion"] = np.where(
    type_summary["Total"] > 0,
    type_summary["Sold"]
    / type_summary["Total"]
    * 100,
    0
)

type_summary = type_summary.sort_values(
    "Conversion",
    ascending=False
)

fig_type_conversion = px.bar(
    type_summary,
    x="Type",
    y="Conversion",
    text=type_summary["Conversion"].round(1).astype(str) + "%",
    hover_data=[
        "Total",
        "Sold",
        "Available"
    ],
    title="Sales Conversion by Type"
)

fig_type_conversion.update_traces(
    textposition="outside"
)

fig_type_conversion.update_layout(
    height=450,
    yaxis_title="Conversion %",
    xaxis_title="Type",
    yaxis=dict(range=[0, 100]),
    margin=dict(
        l=20,
        r=20,
        t=60,
        b=20
    )
)

st.plotly_chart(
    fig_type_conversion,
    use_container_width=True
)


# =========================================================
# PHASE SUMMARY TABLE
# =========================================================

st.markdown(
    '<div class="section-title">Phase Summary</div>',
    unsafe_allow_html=True
)

phase_table = (
    filtered_df
    .groupby("Phase")
    .agg(
        Total=("Phase", "size"),
        Sold=("STATUS", lambda x:
              (x.str.upper() == "SOLD").sum()),
        Available=("STATUS", lambda x:
                   (x.str.upper() == "AVAILABLE").sum()),
        Hold=("STATUS", lambda x:
              (x.str.upper() == "HOLD").sum()),
        Reserved=("STATUS", lambda x:
                  (x.str.upper() == "RESERVED").sum())
    )
    .reset_index()
)

phase_table["Conversion"] = (
    phase_table["Sold"]
    / phase_table["Total"]
    * 100
)

phase_table["Conversion"] = (
    phase_table["Conversion"]
    .round(1)
    .astype(str)
    + "%"
)

st.dataframe(
    phase_table,
    use_container_width=True,
    hide_index=True
)
#data
st.markdown(
    '<div class="section-title">Filtered Unit Details</div>',
    unsafe_allow_html=True
)

display_columns = [
    "Phase",
    "Unit Code",
    "Unit Type",
    "Building",
    "Floor",
    "Type",
    "In/Area",
    "STATUS"
]

# Add client if available
if "Name of client" in filtered_df.columns:
    display_columns.append("Name of client")

display_df = filtered_df[
    [col for col in display_columns if col in filtered_df.columns]
].copy()

st.dataframe(
    display_df,
    use_container_width=True,
    hide_index=True,
    height=500
)
