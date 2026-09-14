
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

type_options = sorted(df["Type"].dropna().unique().tolist())
selected_types = st.sidebar.multiselect(
    "Type", type_options, default=[], key="filter_type"
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

if selected_types:
    filtered = filtered[filtered["Type"].isin(selected_types)]

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

# ============================================================
# DATA + QUICK ACTIONS
# ============================================================
from io import BytesIO
from openpyxl import load_workbook

st.divider()
st.markdown('<div class="section-title">Unit Data</div>', unsafe_allow_html=True)

if "editable_df" not in st.session_state:
    st.session_state.editable_df = df.copy()

editable_df = st.session_state.editable_df

def save_to_excel(dataframe):
    if not DATA_FILE.exists():
        raise FileNotFoundError(f"Excel file not found: {DATA_FILE}")

    wb = load_workbook(DATA_FILE)

    if SHEET in wb.sheetnames:
        wb.remove(wb[SHEET])

    ws = wb.create_sheet(SHEET)

    # Keep the original header position: Excel row 5.
    for _ in range(4):
        ws.append([])

    for col_num, col_name in enumerate(dataframe.columns, start=1):
        ws.cell(row=5, column=col_num, value=str(col_name))

    for row_num, row in enumerate(dataframe.itertuples(index=False, name=None), start=6):
        for col_num, value in enumerate(row, start=1):
            if pd.isna(value):
                value = None
            elif isinstance(value, np.generic):
                value = value.item()
            ws.cell(row=row_num, column=col_num, value=value)

    wb.save(DATA_FILE)
    load_data.clear()

def clean_before_save(dataframe):
    result = dataframe.copy()

    for col in ["In/Area", "NEW M.PRICE", "M.PRICE", "Total Unit", "Rooms Num", "Floor"]:
        if col in result.columns:
            result[col] = pd.to_numeric(result[col], errors="coerce")

    for col in [
        "Phase", "Unit Code", "Unit Type", "Building", "Floor",
        "Apartment NO.", "Type", "STATUS", "Rooms Num", "Name of client"
    ]:
        if col in result.columns:
            result[col] = result[col].astype("string").str.strip()

    if "Phase" in result.columns:
        result["Phase"] = result["Phase"].str.upper()

    return result

# -----------------------------
# Edit dialog
# -----------------------------
@st.dialog("✏️ Edit Unit")
def edit_unit_dialog(row_id):
    data = st.session_state.editable_df
    row = data.loc[row_id]

    c1, c2 = st.columns(2)

    with c1:
        phase = st.selectbox(
            "Phase", ["ALBA", "ORCHID"],
            index=["ALBA", "ORCHID"].index(str(row.get("Phase", "")))
            if str(row.get("Phase", "")) in ["ALBA", "ORCHID"] else 0,
            key=f"edit_phase_{row_id}"
        )
        unit_code = st.text_input(
            "Unit Code", value="" if pd.isna(row.get("Unit Code")) else str(row.get("Unit Code")),
            key=f"edit_code_{row_id}"
        )
        unit_type = st.text_input(
            "Unit Type", value="" if pd.isna(row.get("Unit Type")) else str(row.get("Unit Type")),
            key=f"edit_unit_type_{row_id}"
        )
        building = st.text_input(
            "Building", value="" if pd.isna(row.get("Building")) else str(row.get("Building")),
            key=f"edit_building_{row_id}"
        )
        apartment = st.text_input(
            "Apartment NO.", value="" if pd.isna(row.get("Apartment NO.")) else str(row.get("Apartment NO.")),
            key=f"edit_apartment_{row_id}"
        )
        unit_area = st.number_input(
            "In/Area",
            value=float(row["In/Area"]) if pd.notna(row.get("In/Area")) else 0.0,
            key=f"edit_area_{row_id}"
        )

    with c2:
        floor = st.number_input(
            "Floor",
            value=int(row["Floor"]) if pd.notna(row.get("Floor")) else 0,
            step=1,
            key=f"edit_floor_{row_id}"
        )
        unit_type2 = st.text_input(
            "Type", value="" if pd.isna(row.get("Type")) else str(row.get("Type")),
            key=f"edit_type_{row_id}"
        )
        status_values = ["SOLD", "Available", "Hold", "Reserved"]
        current_status = str(row.get("STATUS", "Available"))
        status = st.selectbox(
            "STATUS", status_values,
            index=status_values.index(current_status) if current_status in status_values else 0,
            key=f"edit_status_{row_id}"
        )
        new_price = st.number_input(
            "NEW M.PRICE",
            value=float(row["NEW M.PRICE"]) if pd.notna(row.get("NEW M.PRICE")) else 0.0,
            key=f"edit_price_{row_id}"
        )
        rooms = st.number_input(
            "Rooms Num",
            value=int(row["Rooms Num"]) if pd.notna(row.get("Rooms Num")) else 0,
            min_value=0,
            step=1,
            key=f"edit_rooms_{row_id}"
        )
        client = st.text_input(
            "Name of client",
            value="" if pd.isna(row.get("Name of client")) else str(row.get("Name of client")),
            key=f"edit_client_{row_id}"
        )

    if st.button("💾 Save", type="primary", use_container_width=True):
        if not unit_code.strip():
            st.warning("Unit Code is required.")
            return

        updated = st.session_state.editable_df.copy()

        values = {
            "Phase": phase,
            "Unit Code": unit_code.strip(),
            "Unit Type": unit_type.strip(),
            "Building": building.strip(),
            "Floor": floor,
            "Apartment NO.": apartment.strip(),
            "Type": unit_type2.strip(),
            "In/Area": unit_area,
            "NEW M.PRICE": new_price,
            "STATUS": status,
            "Rooms Num": rooms,
            "Name of client": client.strip(),
        }

        for col, value in values.items():
            if col in updated.columns:
                updated.at[row_id, col] = value

        try:
            updated = clean_before_save(updated)
            save_to_excel(updated)
            st.session_state.editable_df = updated
            st.success("Unit updated successfully.")
            st.rerun()
        except Exception as e:
            st.error(f"Could not save changes: {e}")

# -----------------------------
# Add dialog
# -----------------------------
@st.dialog("➕ Add New Unit")
def add_unit_dialog():
    c1, c2 = st.columns(2)

    with c1:
        phase = st.selectbox("Phase", ["ALBA", "ORCHID"], key="quick_add_phase")
        unit_code = st.text_input("Unit Code", key="quick_add_code")
        unit_type = st.text_input("Unit Type", key="quick_add_unit_type")
        building = st.text_input("Building", key="quick_add_building")
        apartment = st.text_input("Apartment NO.", key="quick_add_apartment")
        area = st.number_input("In/Area", min_value=0.0, value=0.0, key="quick_add_area")

    with c2:
        floor = st.number_input("Floor", value=0, step=1, key="quick_add_floor")
        unit_type_value = st.text_input("Type", key="quick_add_type")
        status = st.selectbox(
            "STATUS", ["SOLD", "Available", "Hold", "Reserved"],
            index=1, key="quick_add_status"
        )
        new_price = st.number_input("NEW M.PRICE", min_value=0.0, value=0.0, key="quick_add_price")
        rooms = st.number_input("Rooms Num", min_value=0, value=0, step=1, key="quick_add_rooms")
        client = st.text_input("Name of client", key="quick_add_client")

    if st.button("➕ Add Unit", type="primary", use_container_width=True):
        if not unit_code.strip():
            st.warning("Unit Code is required.")
            return

        current = st.session_state.editable_df.copy()
        new_row = {col: pd.NA for col in current.columns}

        values = {
            "Phase": phase,
            "Unit Code": unit_code.strip(),
            "Unit Type": unit_type.strip(),
            "Building": building.strip(),
            "Floor": floor,
            "Apartment NO.": apartment.strip(),
            "Type": unit_type_value.strip(),
            "In/Area": area,
            "NEW M.PRICE": new_price,
            "STATUS": status,
            "Rooms Num": rooms,
            "Name of client": client.strip(),
        }

        for col, value in values.items():
            if col in new_row:
                new_row[col] = value

        updated = pd.concat(
            [current, pd.DataFrame([new_row])],
            ignore_index=True
        )

        try:
            updated = clean_before_save(updated)
            save_to_excel(updated)
            st.session_state.editable_df = updated
            st.success("Unit added successfully.")
            st.rerun()
        except Exception as e:
            st.error(f"Could not add the unit: {e}")

# -----------------------------
# Compact action bar
# -----------------------------
action1, action2, action3, action4 = st.columns([1, 1, 1, 1])

with action1:
    if st.button("➕", help="Add a new unit", key="add_icon", use_container_width=True):
        add_unit_dialog()

with action2:
    st.caption("Add")

with action3:
    st.download_button(
        "📥",
        data=BytesIO(),
        disabled=True,
        help="Use the Export button below to download the current filtered data.",
        key="export_placeholder"
    )

# -----------------------------
# Compact data table with row actions
# -----------------------------
st.caption("Use the small ✏️ and 🗑️ icons beside a unit to edit or delete it.")

display_cols = [
    c for c in [
        "Phase", "Unit Code", "Unit Type", "Building", "Floor",
        "Apartment NO.", "Type", "In/Area", "NEW M.PRICE",
        "STATUS", "Rooms Num", "Name of client"
    ] if c in filtered.columns
]

# Use the actual editable dataframe indexes so actions target the correct records.
visible = filtered[display_cols].copy()

if visible.empty:
    st.info("No units match the selected filters.")
else:
    header = st.columns([0.5] + [1.15] * len(display_cols) + [0.42, 0.42])

    header[0].markdown("**#**")
    for i, col in enumerate(display_cols, start=1):
        header[i].markdown(f"**{col}**")
    header[-2].markdown("**✏️**")
    header[-1].markdown("**🗑️**")

    for row_id, row in visible.iterrows():
        cells = st.columns([0.5] + [1.15] * len(display_cols) + [0.42, 0.42])

        cells[0].write(str(row_id + 1))

        for i, col in enumerate(display_cols, start=1):
            value = row[col]
            if pd.isna(value):
                value = "—"
            elif col in ["In/Area", "NEW M.PRICE"]:
                try:
                    value = f"{float(value):,.0f}"
                except Exception:
                    value = str(value)
            cells[i].write(str(value))

        with cells[-2]:
            if st.button(
                "✏️",
                key=f"edit_row_{row_id}",
                help=f"Edit {row.get('Unit Code', '')}",
            ):
                edit_unit_dialog(row_id)

        with cells[-1]:
            if st.button(
                "🗑️",
                key=f"delete_row_{row_id}",
                help=f"Delete {row.get('Unit Code', '')}",
            ):
                st.session_state.delete_confirm_id = row_id
                st.rerun()

# -----------------------------
# Delete confirmation
# -----------------------------
if "delete_confirm_id" in st.session_state:
    delete_id = st.session_state.delete_confirm_id

    if delete_id in st.session_state.editable_df.index:
        delete_code = st.session_state.editable_df.loc[delete_id, "Unit Code"]

        st.warning(f"Delete unit **{delete_code}**? This will permanently update the Excel file.")

        dc1, dc2 = st.columns(2)

        with dc1:
            if st.button("Delete", type="primary", use_container_width=True, key="confirm_delete"):
                try:
                    updated = (
                        st.session_state.editable_df
                        .drop(index=delete_id)
                        .reset_index(drop=True)
                    )
                    updated = clean_before_save(updated)
                    save_to_excel(updated)
                    st.session_state.editable_df = updated
                    del st.session_state.delete_confirm_id
                    st.success("Unit deleted successfully.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Could not delete the unit: {e}")

        with dc2:
            if st.button("Cancel", use_container_width=True, key="cancel_delete"):
                del st.session_state.delete_confirm_id
                st.rerun()

# -----------------------------
# Export filtered data
# -----------------------------
with st.expander("📥 Export", expanded=False):
    export_cols = [
        c for c in [
            "Phase", "Unit Code", "Unit Type", "Building", "Floor",
            "Apartment NO.", "Type", "In/Area", "NEW M.PRICE",
            "STATUS", "Rooms Num", "Name of client"
        ] if c in filtered.columns
    ]

    export_df = filtered[export_cols].copy()

    export_buffer = BytesIO()
    with pd.ExcelWriter(export_buffer, engine="openpyxl") as writer:
        export_df.to_excel(writer, index=False, sheet_name="Gardenia Town")

    export_buffer.seek(0)

    st.download_button(
        "📥 Download Filtered Data as Excel",
        data=export_buffer,
        file_name="Gardenia_Town_Export.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
