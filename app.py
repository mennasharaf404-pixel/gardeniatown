
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from pathlib import Path

st.set_page_config(
    page_title="Gardenia Town | ALBA vs ORCHID",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded",
)

DATA_FILE = Path(__file__).parent / "data" / "All Inventory Project(1).xlsx"
SHEET = "Gardenia Town"

# -----------------------------
# Style
# -----------------------------
st.markdown("""
<style>
.block-container {padding-top:1.2rem; padding-bottom:2rem;}
.main-title {font-size:2.2rem;font-weight:800;letter-spacing:.5px;}
.sub-title {color:#667085;margin-bottom:1rem;}
.section-title {font-size:1.15rem;font-weight:750;margin:1.1rem 0 .65rem;}
div[data-testid="stMetric"] {
    background:#fff;border:1px solid #e4e7ec;border-radius:14px;
    padding:14px 16px;box-shadow:0 2px 8px rgba(16,24,40,.05);
}
.filter-note {
    background:#f8fafc;border:1px solid #eaecf0;border-radius:10px;
    padding:9px 12px;color:#475467;font-size:.9rem;
}
</style>
""", unsafe_allow_html=True)

# -----------------------------
# Load data
# -----------------------------
@st.cache_data
def load_data(path):
    df = pd.read_excel(path, sheet_name=SHEET, header=4)
    df = df.dropna(how="all").copy()
    df.columns = [str(c).strip() for c in df.columns]

    text_cols = [
        "Phase", "Unit Code", "Unit Type", "Building", "Floor",
        "Apartment NO.", "Type", "STATUS", "Rooms Num", "Name of client"
    ]
    for col in text_cols:
        if col in df.columns:
            df[col] = df[col].astype("string").str.strip()

    if "Phase" in df.columns:
        df["Phase"] = df["Phase"].str.upper()
    if "STATUS" in df.columns:
        df["STATUS"] = df["STATUS"].str.strip()

    # Numeric columns used by filters / analysis
    for col in ["In/Area", "NEW M.PRICE", "M.PRICE", "Total Unit", "Rooms Num", "Floor"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    return df

try:
    df = load_data(DATA_FILE)
except Exception as e:
    st.error(f"Could not load the Excel file: {e}")
    st.stop()

required = ["Phase", "Unit Code", "Unit Type", "Building", "Floor", "STATUS"]
missing = [c for c in required if c not in df.columns]
if missing:
    st.error("Missing columns in Gardenia Town: " + ", ".join(missing))
    st.stop()

# -----------------------------
# Sidebar - stronger filters
# -----------------------------
st.sidebar.header("Dashboard Filters")

if st.sidebar.button("Reset filters", use_container_width=True):
    for key in list(st.session_state.keys()):
        if key.startswith("filter_"):
            del st.session_state[key]
    st.rerun()

phase_options = ["ALBA", "ORCHID"]
phase_options = [p for p in phase_options if p in set(df["Phase"].dropna())]
selected_phases = st.sidebar.multiselect(
    "Phase", phase_options, default=phase_options, key="filter_phase"
)

status_options = sorted(df["STATUS"].dropna().unique().tolist())
selected_statuses = st.sidebar.multiselect(
    "Status", status_options, default=status_options, key="filter_status"
)

building_options = sorted(df["Building"].dropna().unique().tolist())
selected_buildings = st.sidebar.multiselect(
    "Building", building_options, default=[], key="filter_building"
)

unit_type_options = sorted(df["Unit Type"].dropna().unique().tolist())
selected_unit_types = st.sidebar.multiselect(
    "Unit Type", unit_type_options, default=[], key="filter_unit_type"
)

floor_values = pd.to_numeric(df["Floor"], errors="coerce").dropna()
if not floor_values.empty:
    floor_min, floor_max = int(floor_values.min()), int(floor_values.max())
    if floor_min < floor_max:
        selected_floor = st.sidebar.slider(
            "Floor", floor_min, floor_max, (floor_min, floor_max),
            key="filter_floor"
        )
    else:
        selected_floor = (floor_min, floor_max)
else:
    selected_floor = None

rooms_values = pd.to_numeric(df["Rooms Num"], errors="coerce").dropna() if "Rooms Num" in df else pd.Series(dtype=float)
if not rooms_values.empty:
    room_options = sorted(rooms_values.astype(int).unique().tolist())
    selected_rooms = st.sidebar.multiselect(
        "Rooms", room_options, default=[], key="filter_rooms"
    )
else:
    selected_rooms = []

area_values = pd.to_numeric(df["In/Area"], errors="coerce").dropna()
if not area_values.empty and area_values.min() < area_values.max():
    area_min, area_max = float(area_values.min()), float(area_values.max())
    selected_area = st.sidebar.slider(
        "Area (m²)", min_value=area_min, max_value=area_max,
        value=(area_min, area_max), step=1.0, key="filter_area"
    )
else:
    selected_area = None

search = st.sidebar.text_input(
    "Search Unit / Client", "", key="filter_search",
    placeholder="e.g. A-101 or client name"
)

# Apply filters
filtered = df.copy()
filtered = filtered[filtered["Phase"].isin(selected_phases)]
filtered = filtered[filtered["STATUS"].isin(selected_statuses)]

if selected_buildings:
    filtered = filtered[filtered["Building"].isin(selected_buildings)]
if selected_unit_types:
    filtered = filtered[filtered["Unit Type"].isin(selected_unit_types)]
if selected_floor is not None:
    filtered = filtered[
        pd.to_numeric(filtered["Floor"], errors="coerce").between(
            selected_floor[0], selected_floor[1], inclusive="both"
        )
    ]
if selected_rooms and "Rooms Num" in filtered:
    filtered = filtered[
        pd.to_numeric(filtered["Rooms Num"], errors="coerce").isin(selected_rooms)
    ]
if selected_area is not None:
    filtered = filtered[
        pd.to_numeric(filtered["In/Area"], errors="coerce").between(
            selected_area[0], selected_area[1], inclusive="both"
        )
    ]

if search.strip():
    q = search.strip().lower()
    unit_match = filtered["Unit Code"].fillna("").astype(str).str.lower().str.contains(q, regex=False)
    client_match = filtered["Name of client"].fillna("").astype(str).str.lower().str.contains(q, regex=False)
    filtered = filtered[unit_match | client_match]

st.markdown('<div class="main-title">GARDENIA TOWN</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">ALBA vs ORCHID — Sales & Inventory Dashboard</div>', unsafe_allow_html=True)

st.markdown(
    f'<div class="filter-note">Showing <b>{len(filtered):,}</b> of '
    f'<b>{len(df):,}</b> Gardenia Town units based on the selected filters.</div>',
    unsafe_allow_html=True
)

# -----------------------------
# KPIs
# -----------------------------
total = len(filtered)
sold = int((filtered["STATUS"].astype(str).str.upper() == "SOLD").sum())
available = int((filtered["STATUS"].astype(str).str.lower() == "available").sum())
hold = int((filtered["STATUS"].astype(str).str.lower() == "hold").sum())
reserved = int((filtered["STATUS"].astype(str).str.lower() == "reserved").sum())
conversion = sold / total if total else 0

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("TOTAL UNITS", f"{total:,}")
c2.metric("SOLD", f"{sold:,}")
c3.metric("AVAILABLE", f"{available:,}")
c4.metric("HOLD", f"{hold:,}")
c5.metric("SALES CONVERSION", f"{conversion:.1%}")

st.divider()

# -----------------------------
# Helper for Plotly charts
# -----------------------------
def finish_chart(fig, height=390):
    fig.update_layout(
        height=height,
        margin=dict(l=20, r=20, t=55, b=45),
        hovermode="x unified",
        legend_title_text="",
    )
    return fig

# -----------------------------
# Interactive status chart
# -----------------------------
st.markdown('<div class="section-title">Inventory Status</div>', unsafe_allow_html=True)

status_rows = []
for phase in phase_options:
    x = filtered[filtered["Phase"] == phase]
    for status in ["SOLD", "Available", "Hold", "Reserved"]:
        count = int(
            (x["STATUS"].astype(str).str.upper() == status.upper()).sum()
        )
        status_rows.append({"Phase": phase, "Status": status, "Units": count})

status_df = pd.DataFrame(status_rows)

fig_status = px.bar(
    status_df,
    x="Status",
    y="Units",
    color="Phase",
    barmode="group",
    text="Units",
    category_orders={"Status": ["SOLD", "Available", "Hold", "Reserved"]},
    labels={"Status": "Status", "Units": "Units"},
)
fig_status.update_traces(textposition="outside", cliponaxis=False, hovertemplate="%{x}<br>Units: %{y}<extra></extra>")
fig_status.update_yaxes(rangemode="tozero")
st.plotly_chart(finish_chart(fig_status), use_container_width=True)

# -----------------------------
# Conversion + building
# -----------------------------
left, right = st.columns(2)

with left:
    st.markdown('<div class="section-title">Sales Conversion</div>', unsafe_allow_html=True)
    conversion_rows = []
    for phase in phase_options:
        x = filtered[filtered["Phase"] == phase]
        t = len(x)
        s = int((x["STATUS"].astype(str).str.upper() == "SOLD").sum())
        conversion_rows.append({"Phase": phase, "Conversion": round((s/t)*100, 1) if t else 0})
    conv_df = pd.DataFrame(conversion_rows)

    fig_conv = px.bar(
        conv_df, x="Phase", y="Conversion", text="Conversion",
        labels={"Conversion": "Sales Conversion (%)", "Phase": "Phase"}
    )
    fig_conv.update_traces(
        texttemplate="%{text:.1f}%", textposition="outside",
        cliponaxis=False,
        hovertemplate="%{x}<br>Conversion: %{y:.1f}%<extra></extra>"
    )
    fig_conv.update_yaxes(range=[0, max(100, float(conv_df["Conversion"].max()) + 10)])
    st.plotly_chart(finish_chart(fig_conv, 350), use_container_width=True)

with right:
    st.markdown('<div class="section-title">Available Units by Building</div>', unsafe_allow_html=True)
    available_df = filtered[
        filtered["STATUS"].astype(str).str.lower() == "available"
    ]
    building_df = (
        available_df.groupby(["Building", "Phase"])
        .size()
        .reset_index(name="Units")
    )
    fig_build = px.bar(
        building_df,
        x="Building", y="Units", color="Phase",
        barmode="group", text="Units",
        labels={"Building": "Building", "Units": "Available Units"}
    )
    fig_build.update_traces(
        textposition="outside", cliponaxis=False,
        hovertemplate="%{x}<br>Available: %{y}<extra></extra>"
    )
    fig_build.update_xaxes(tickangle=-45)
    st.plotly_chart(finish_chart(fig_build, 350), use_container_width=True)

# -----------------------------
# Unit type + phase summary
# -----------------------------
left, right = st.columns(2)

with left:
    st.markdown('<div class="section-title">Unit Type Mix</div>', unsafe_allow_html=True)
    type_df = filtered.groupby(["Unit Type", "Phase"]).size().reset_index(name="Units")
    fig_type = px.bar(
        type_df, x="Unit Type", y="Units", color="Phase",
        barmode="group", text="Units",
        labels={"Unit Type": "Unit Type", "Units": "Units"}
    )
    fig_type.update_traces(textposition="outside", cliponaxis=False)
    fig_type.update_xaxes(tickangle=-35)
    st.plotly_chart(finish_chart(fig_type, 390), use_container_width=True)

with right:
    st.markdown('<div class="section-title">Phase Summary</div>', unsafe_allow_html=True)
    phase_summary = []
    for phase in phase_options:
        x = filtered[filtered["Phase"] == phase]
        t = len(x)
        s = int((x["STATUS"].astype(str).str.upper() == "SOLD").sum())
        a = int((x["STATUS"].astype(str).str.lower() == "available").sum())
        phase_summary.append({
            "Phase": phase,
            "Total": t,
            "Sold": s,
            "Available": a,
            "Conversion": round((s/t)*100, 1) if t else 0
        })
    st.dataframe(
        pd.DataFrame(phase_summary),
        use_container_width=True,
        hide_index=True,
        column_config={
            "Conversion": st.column_config.NumberColumn("Conversion", format="%.1f%%")
        }
    )

# -----------------------------
# Detail data
# -----------------------------
st.markdown('<div class="section-title">Filtered Unit Details</div>', unsafe_allow_html=True)
display_cols = [
    c for c in [
        "Phase", "Unit Code", "Unit Type", "Building", "Floor",
        "Apartment NO.", "Type", "In/Area", "NEW M.PRICE",
        "STATUS", "Rooms Num", "Name of client"
    ] if c in filtered.columns
]

st.dataframe(
    filtered[display_cols].reset_index(drop=True),
    use_container_width=True,
    hide_index=True,
    height=430,
)

st.caption(
    "Charts are interactive: hover for details, click legend items to show/hide phases, "
    "and use the Plotly toolbar to zoom or reset the view."
)
