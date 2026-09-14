
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from pathlib import Path
from io import BytesIO
from openpyxl import load_workbook

st.set_page_config(
    page_title="Gardenia Town | ALBA vs ORCHID",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded",
)

DATA_FILE = Path(__file__).parent / "data" / "All Inventory Project(1).xlsx"
SHEET = "Gardenia Town"

# =========================================================
# STYLE
# =========================================================
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
.action-bar {
    background:#f8fafc;border:1px solid #eaecf0;border-radius:12px;
    padding:12px 14px;margin-bottom:10px;
}
</style>
""", unsafe_allow_html=True)


# =========================================================
# LOAD DATA
# =========================================================
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

    for col in [
        "In/Area", "NEW M.PRICE", "M.PRICE",
        "Total Unit", "Rooms Num", "Floor"
    ]:
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


# =========================================================
# FAST DATA MANAGEMENT HELPERS
# =========================================================
def save_dataframe_to_excel(new_df):
    """
    Update ONLY the Gardenia Town sheet values.
    Other workbook sheets are left untouched.
    The original workbook is not changed until the user
    explicitly clicks Save.
    """
    wb = load_workbook(DATA_FILE)
    if SHEET not in wb.sheetnames:
        raise ValueError(f"Sheet '{SHEET}' was not found.")

    ws = wb[SHEET]

    # Header is Excel row 5.
    header_row = 5
    start_row = header_row + 1

    # Preserve the existing workbook columns.
    existing_headers = [
        ws.cell(header_row, col).value
        for col in range(1, ws.max_column + 1)
    ]

    # Map normalized header names to worksheet columns.
    header_map = {}
    for col_idx, value in enumerate(existing_headers, start=1):
        if value is not None:
            header_map[str(value).strip()] = col_idx

    # Make sure every dataframe column exists in the workbook.
    for column in new_df.columns:
        if str(column).strip() not in header_map:
            raise ValueError(
                f"Column '{column}' is not present in the original Excel sheet."
            )

    # Write the dataframe values into the existing rows.
    for r_offset, (_, row) in enumerate(new_df.iterrows()):
        excel_row = start_row + r_offset
        for column in new_df.columns:
            col_idx = header_map[str(column).strip()]
            value = row[column]

            if pd.isna(value):
                value = None

            # Convert numpy scalar values to normal Python values.
            if isinstance(value, np.generic):
                value = value.item()

            ws.cell(excel_row, col_idx).value = value

    # Clear old rows that are now beyond the dataframe.
    old_last_data_row = max(start_row, ws.max_row)
    new_last_data_row = start_row + len(new_df) - 1

    for excel_row in range(new_last_data_row + 1, old_last_data_row + 1):
        for col_idx in range(1, ws.max_column + 1):
            ws.cell(excel_row, col_idx).value = None

    wb.save(DATA_FILE)
    load_data.clear()


def make_excel_download(dataframe):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        dataframe.to_excel(writer, index=False, sheet_name=SHEET)
    output.seek(0)
    return output.getvalue()


def clean_value(value):
    if pd.isna(value):
        return None
    if isinstance(value, np.generic):
        return value.item()
    return value


def show_edit_dialog(row_index):
    row = df.loc[row_index].copy()
    columns = list(df.columns)

    @st.dialog("Edit Unit")
    def dialog():
        st.caption("Edit the selected record. Nothing is written to Excel until you click Save changes.")

        edited = {}
        half = (len(columns) + 1) // 2
        col_groups = [columns[:half], columns[half:]]

        ui_cols = st.columns(2)

        for ui_col, group in zip(ui_cols, col_groups):
            with ui_col:
                for column in group:
                    value = clean_value(row[column])

                    if column in ["Floor", "Rooms Num", "In/Area", "NEW M.PRICE", "M.PRICE", "Total Unit"]:
                        try:
                            number_value = float(value) if value is not None else 0.0
                        except Exception:
                            number_value = 0.0

                        edited[column] = st.number_input(
                            column,
                            value=number_value,
                            key=f"edit_{row_index}_{column}",
                        )
                    else:
                        edited[column] = st.text_input(
                            column,
                            value="" if value is None else str(value),
                            key=f"edit_{row_index}_{column}",
                        )

        save_col, cancel_col = st.columns(2)

        with save_col:
            if st.button("Save changes", type="primary", use_container_width=True):
                for column in columns:
                    df.loc[row_index, column] = edited[column]

                try:
                    save_dataframe_to_excel(df)
                    st.success("Changes saved successfully.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Could not save changes: {e}")

        with cancel_col:
            if st.button("Cancel", use_container_width=True):
                st.rerun()

    dialog()


def show_add_dialog():
    columns = list(df.columns)

    @st.dialog("Add New Unit")
    def dialog():
        st.caption("Add one new record. Existing records are not changed.")

        new_values = {}
        half = (len(columns) + 1) // 2
        col_groups = [columns[:half], columns[half:]]

        ui_cols = st.columns(2)

        for ui_col, group in zip(ui_cols, col_groups):
            with ui_col:
                for column in group:
                    if column in ["Floor", "Rooms Num", "In/Area", "NEW M.PRICE", "M.PRICE", "Total Unit"]:
                        new_values[column] = st.number_input(
                            column,
                            value=0.0,
                            key=f"add_{column}",
                        )
                    else:
                        new_values[column] = st.text_input(
                            column,
                            value="",
                            key=f"add_{column}",
                        )

        save_col, cancel_col = st.columns(2)

        with save_col:
            if st.button("Add unit", type="primary", use_container_width=True):
                if not str(new_values.get("Unit Code", "")).strip():
                    st.error("Unit Code is required.")
                    return

                new_row = {}
                for column in columns:
                    value = new_values.get(column)

                    if isinstance(value, str) and not value.strip():
                        value = None

                    new_row[column] = value

                updated_df = pd.concat(
                    [df, pd.DataFrame([new_row], columns=columns)],
                    ignore_index=True,
                )

                try:
                    save_dataframe_to_excel(updated_df)
                    st.success("New unit added successfully.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Could not add the unit: {e}")

        with cancel_col:
            if st.button("Cancel", use_container_width=True):
                st.rerun()

    dialog()


def show_delete_dialog(row_index):
    unit_code = str(df.loc[row_index, "Unit Code"])
    phase = str(df.loc[row_index, "Phase"])
    building = str(df.loc[row_index, "Building"])

    @st.dialog("Delete Unit")
    def dialog():
        st.warning(
            f"Are you sure you want to delete Unit Code '{unit_code}' "
            f"from {phase} / Building {building}?"
        )
        st.caption("This action changes the Excel file only after you confirm.")

        yes_col, no_col = st.columns(2)

        with yes_col:
            if st.button("Delete", type="primary", use_container_width=True):
                updated_df = df.drop(index=row_index).reset_index(drop=True)

                try:
                    save_dataframe_to_excel(updated_df)
                    st.success("Unit deleted successfully.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Could not delete the unit: {e}")

        with no_col:
            if st.button("Cancel", use_container_width=True):
                st.rerun()

    dialog()


# =========================================================
# SIDEBAR FILTERS
# =========================================================
st.sidebar.header("Dashboard Filters")

if st.sidebar.button("Reset filters", use_container_width=True):
    for key in list(st.session_state.keys()):
        if key.startswith("filter_"):
            del st.session_state[key]
    st.rerun()

phase_options = ["ALBA", "ORCHID"]
phase_options = [
    p for p in phase_options
    if p in set(df["Phase"].dropna())
]

selected_phases = st.sidebar.multiselect(
    "Phase",
    phase_options,
    default=phase_options,
    key="filter_phase",
)

status_options = sorted(df["STATUS"].dropna().unique().tolist())
selected_statuses = st.sidebar.multiselect(
    "Status",
    status_options,
    default=status_options,
    key="filter_status",
)

type_options = sorted(df["Type"].dropna().unique().tolist())
selected_types = st.sidebar.multiselect(
    "Type",
    type_options,
    default=[],
    key="filter_type",
)

building_options = sorted(df["Building"].dropna().unique().tolist())
selected_buildings = st.sidebar.multiselect(
    "Building",
    building_options,
    default=[],
    key="filter_building",
)

unit_type_options = sorted(df["Unit Type"].dropna().unique().tolist())
selected_unit_types = st.sidebar.multiselect(
    "Unit Type",
    unit_type_options,
    default=[],
    key="filter_unit_type",
)

floor_values = pd.to_numeric(df["Floor"], errors="coerce").dropna()

if not floor_values.empty:
    floor_min, floor_max = int(floor_values.min()), int(floor_values.max())

    if floor_min < floor_max:
        selected_floor = st.sidebar.slider(
            "Floor",
            floor_min,
            floor_max,
            (floor_min, floor_max),
            key="filter_floor",
        )
    else:
        selected_floor = (floor_min, floor_max)
else:
    selected_floor = None

rooms_values = (
    pd.to_numeric(df["Rooms Num"], errors="coerce").dropna()
    if "Rooms Num" in df
    else pd.Series(dtype=float)
)

if not rooms_values.empty:
    room_options = sorted(rooms_values.astype(int).unique().tolist())
    selected_rooms = st.sidebar.multiselect(
        "Rooms",
        room_options,
        default=[],
        key="filter_rooms",
    )
else:
    selected_rooms = []

area_values = pd.to_numeric(df["In/Area"], errors="coerce").dropna()

if not area_values.empty and area_values.min() < area_values.max():
    area_min, area_max = float(area_values.min()), float(area_values.max())
    selected_area = st.sidebar.slider(
        "Area (m²)",
        min_value=area_min,
        max_value=area_max,
        value=(area_min, area_max),
        step=1.0,
        key="filter_area",
    )
else:
    selected_area = None

search = st.sidebar.text_input(
    "Search Unit / Client",
    "",
    key="filter_search",
    placeholder="e.g. A-101 or client name",
)


# =========================================================
# APPLY FILTERS
# =========================================================
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
            selected_floor[0],
            selected_floor[1],
            inclusive="both",
        )
    ]

if selected_rooms and "Rooms Num" in filtered:
    filtered = filtered[
        pd.to_numeric(filtered["Rooms Num"], errors="coerce").isin(selected_rooms)
    ]

if selected_area is not None:
    filtered = filtered[
        pd.to_numeric(filtered["In/Area"], errors="coerce").between(
            selected_area[0],
            selected_area[1],
            inclusive="both",
        )
    ]

if search.strip():
    q = search.strip().lower()

    unit_match = (
        filtered["Unit Code"]
        .fillna("")
        .astype(str)
        .str.lower()
        .str.contains(q, regex=False)
    )

    client_match = (
        filtered["Name of client"]
        .fillna("")
        .astype(str)
        .str.lower()
        .str.contains(q, regex=False)
    )

    filtered = filtered[unit_match | client_match]


# =========================================================
# HEADER
# =========================================================
st.markdown(
    '<div class="main-title">GARDENIA TOWN</div>',
    unsafe_allow_html=True,
)

st.markdown(
    '<div class="sub-title">ALBA vs ORCHID — Sales & Inventory Dashboard</div>',
    unsafe_allow_html=True,
)

st.markdown(
    f'<div class="filter-note">Showing <b>{len(filtered):,}</b> of '
    f'<b>{len(df):,}</b> Gardenia Town units based on the selected filters.</div>',
    unsafe_allow_html=True,
)


# =========================================================
# KPIs
# =========================================================
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


# =========================================================
# CHART HELPER
# =========================================================
def finish_chart(fig, height=390):
    fig.update_layout(
        height=height,
        margin=dict(l=20, r=20, t=55, b=45),
        hovermode="x unified",
        legend_title_text="",
    )
    return fig


# =========================================================
# INVENTORY STATUS
# =========================================================
st.markdown(
    '<div class="section-title">Inventory Status</div>',
    unsafe_allow_html=True,
)

status_rows = []

for phase in phase_options:
    x = filtered[filtered["Phase"] == phase]

    for status in ["SOLD", "Available", "Hold", "Reserved"]:
        count = int(
            (x["STATUS"].astype(str).str.upper() == status.upper()).sum()
        )
        status_rows.append(
            {"Phase": phase, "Status": status, "Units": count}
        )

status_df = pd.DataFrame(status_rows)

fig_status = px.bar(
    status_df,
    x="Status",
    y="Units",
    color="Phase",
    barmode="group",
    text="Units",
    category_orders={
        "Status": ["SOLD", "Available", "Hold", "Reserved"]
    },
    labels={"Status": "Status", "Units": "Units"},
)

fig_status.update_traces(
    textposition="outside",
    cliponaxis=False,
    hovertemplate="%{x}<br>Units: %{y}<extra></extra>",
)

fig_status.update_yaxes(rangemode="tozero")

st.plotly_chart(
    finish_chart(fig_status),
    use_container_width=True,
)


# =========================================================
# CONVERSION + BUILDING
# =========================================================
left, right = st.columns(2)

with left:
    st.markdown(
        '<div class="section-title">Sales Conversion</div>',
        unsafe_allow_html=True,
    )

    conversion_rows = []

    for phase in phase_options:
        x = filtered[filtered["Phase"] == phase]
        t = len(x)
        s = int(
            (x["STATUS"].astype(str).str.upper() == "SOLD").sum()
        )

        conversion_rows.append(
            {
                "Phase": phase,
                "Conversion": round((s / t) * 100, 1) if t else 0,
            }
        )

    conv_df = pd.DataFrame(conversion_rows)

    fig_conv = px.bar(
        conv_df,
        x="Phase",
        y="Conversion",
        text="Conversion",
        labels={
            "Conversion": "Sales Conversion (%)",
            "Phase": "Phase",
        },
    )

    fig_conv.update_traces(
        texttemplate="%{text:.1f}%",
        textposition="outside",
        cliponaxis=False,
        hovertemplate="%{x}<br>Conversion: %{y:.1f}%<extra></extra>",
    )

    fig_conv.update_yaxes(
        range=[
            0,
            max(100, float(conv_df["Conversion"].max()) + 10),
        ]
    )

    st.plotly_chart(
        finish_chart(fig_conv, 350),
        use_container_width=True,
    )


with right:
    st.markdown(
        '<div class="section-title">Available Units by Building</div>',
        unsafe_allow_html=True,
    )

    available_df = filtered[
        filtered["STATUS"].astype(str).str.lower() == "available"
    ]

    building_df = (
        available_df
        .groupby(["Building", "Phase"])
        .size()
        .reset_index(name="Units")
    )

    fig_build = px.bar(
        building_df,
        x="Building",
        y="Units",
        color="Phase",
        barmode="group",
        text="Units",
        labels={
            "Building": "Building",
            "Units": "Available Units",
        },
    )

    fig_build.update_traces(
        textposition="outside",
        cliponaxis=False,
        hovertemplate="%{x}<br>Available: %{y}<extra></extra>",
    )

    fig_build.update_xaxes(tickangle=-45)

    st.plotly_chart(
        finish_chart(fig_build, 350),
        use_container_width=True,
    )


# =========================================================
# UNIT TYPE + PHASE SUMMARY
# =========================================================
left, right = st.columns(2)

with left:
    st.markdown(
        '<div class="section-title">Unit Type Mix</div>',
        unsafe_allow_html=True,
    )

    type_df = (
        filtered
        .groupby(["Unit Type", "Phase"])
        .size()
        .reset_index(name="Units")
    )

    fig_type = px.bar(
        type_df,
        x="Unit Type",
        y="Units",
        color="Phase",
        barmode="group",
        text="Units",
        labels={
            "Unit Type": "Unit Type",
            "Units": "Units",
        },
    )

    fig_type.update_traces(
        textposition="outside",
        cliponaxis=False,
    )

    fig_type.update_xaxes(tickangle=-35)

    st.plotly_chart(
        finish_chart(fig_type, 390),
        use_container_width=True,
    )


with right:
    st.markdown(
        '<div class="section-title">Phase Summary</div>',
        unsafe_allow_html=True,
    )

    phase_summary = []

    for phase in phase_options:
        x = filtered[filtered["Phase"] == phase]
        t = len(x)

        s = int(
            (x["STATUS"].astype(str).str.upper() == "SOLD").sum()
        )

        a = int(
            (x["STATUS"].astype(str).str.lower() == "available").sum()
        )

        phase_summary.append(
            {
                "Phase": phase,
                "Total": t,
                "Sold": s,
                "Available": a,
                "Conversion": round((s / t) * 100, 1) if t else 0,
            }
        )

    st.dataframe(
        pd.DataFrame(phase_summary),
        use_container_width=True,
        hide_index=True,
        column_config={
            "Conversion": st.column_config.NumberColumn(
                "Conversion",
                format="%.1f%%",
            )
        },
    )


# =========================================================
# ADD / EDIT DIALOGS
# =========================================================
# The dialogs below use choices taken directly from the
# existing Excel data. This makes entry easier and avoids
# changing the original values or terminology.

def show_edit_dialog(row_index):
    row = df.loc[row_index].copy()
    columns = list(df.columns)

    @st.dialog("Edit Unit")
    def dialog():
        st.caption(
            "Update the selected unit. Changes are saved only when "
            "you click Save changes."
        )

        edited = {}

        # Use choices from the existing workbook wherever possible.
        choice_columns = {
            "Phase": sorted(df["Phase"].dropna().astype(str).unique().tolist()),
            "Type": sorted(df["Type"].dropna().astype(str).unique().tolist()),
            "STATUS": sorted(df["STATUS"].dropna().astype(str).unique().tolist()),
            "Unit Type": sorted(df["Unit Type"].dropna().astype(str).unique().tolist()),
            "Building": sorted(df["Building"].dropna().astype(str).unique().tolist()),
        }

        numeric_columns = {
            "Floor", "Rooms Num", "In/Area", "NEW M.PRICE",
            "M.PRICE", "Total Unit"
        }

        half = (len(columns) + 1) // 2
        ui_cols = st.columns(2)

        for ui_col, group in zip(ui_cols, [columns[:half], columns[half:]]):
            with ui_col:
                for column in group:
                    value = clean_value(row[column])

                    if column in choice_columns:
                        options = choice_columns[column]
                        current = "" if value is None else str(value)

                        if current and current not in options:
                            options = [current] + options

                        edited[column] = st.selectbox(
                            column,
                            options if options else [""],
                            index=options.index(current) if current in options else 0,
                            key=f"edit_choice_{row_index}_{column}",
                        )

                    elif column in numeric_columns:
                        try:
                            number_value = float(value) if value is not None else 0.0
                        except Exception:
                            number_value = 0.0

                        edited[column] = st.number_input(
                            column,
                            value=number_value,
                            key=f"edit_num_{row_index}_{column}",
                        )

                    else:
                        edited[column] = st.text_input(
                            column,
                            value="" if value is None else str(value),
                            key=f"edit_text_{row_index}_{column}",
                        )

        save_col, cancel_col = st.columns(2)

        with save_col:
            if st.button(
                "Save changes",
                type="primary",
                use_container_width=True,
            ):
                for column in columns:
                    df.loc[row_index, column] = edited[column]

                try:
                    save_dataframe_to_excel(df)
                    st.success("Changes saved successfully.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Could not save changes: {e}")

        with cancel_col:
            if st.button("Cancel", use_container_width=True):
                st.rerun()

    dialog()


def show_add_dialog():
    columns = list(df.columns)

    @st.dialog("Add New Unit")
    def dialog():
        st.caption(
            "Choose values from the existing data where possible. "
            "Only the new record will be added."
        )

        new_values = {}

        # Choices are generated from the real Excel data.
        # Nothing is invented or changed in the existing records.
        choice_columns = {
            "Phase": sorted(df["Phase"].dropna().astype(str).unique().tolist()),
            "Type": sorted(df["Type"].dropna().astype(str).unique().tolist()),
            "STATUS": sorted(df["STATUS"].dropna().astype(str).unique().tolist()),
            "Unit Type": sorted(df["Unit Type"].dropna().astype(str).unique().tolist()),
            "Building": sorted(df["Building"].dropna().astype(str).unique().tolist()),
        }

        numeric_columns = {
            "Floor", "Rooms Num", "In/Area", "NEW M.PRICE",
            "M.PRICE", "Total Unit"
        }

        half = (len(columns) + 1) // 2
        ui_cols = st.columns(2)

        for ui_col, group in zip(ui_cols, [columns[:half], columns[half:]]):
            with ui_col:
                for column in group:

                    if column in choice_columns:
                        options = choice_columns[column]

                        if options:
                            new_values[column] = st.selectbox(
                                column,
                                [""] + options,
                                index=0,
                                key=f"add_choice_{column}",
                            )
                        else:
                            new_values[column] = st.text_input(
                                column,
                                key=f"add_text_{column}",
                            )

                    elif column in numeric_columns:
                        new_values[column] = st.number_input(
                            column,
                            value=0.0,
                            key=f"add_num_{column}",
                        )

                    else:
                        new_values[column] = st.text_input(
                            column,
                            value="",
                            key=f"add_text_{column}",
                        )

        # Required fields
        required_add = ["Phase", "Unit Code", "Unit Type", "Building", "STATUS"]

        missing_required = [
            c for c in required_add
            if not str(new_values.get(c, "")).strip()
        ]

        save_col, cancel_col = st.columns(2)

        with save_col:
            if st.button(
                "Add unit",
                type="primary",
                use_container_width=True,
            ):
                if missing_required:
                    st.error(
                        "Please complete: "
                        + ", ".join(missing_required)
                    )
                    return

                new_row = {}

                for column in columns:
                    value = new_values.get(column)

                    if isinstance(value, str) and not value.strip():
                        value = None

                    new_row[column] = value

                # Prevent accidentally creating a duplicate Unit Code.
                existing_codes = (
                    df["Unit Code"]
                    .dropna()
                    .astype(str)
                    .str.strip()
                    .str.lower()
                )

                new_code = str(new_row["Unit Code"]).strip().lower()

                if new_code in set(existing_codes):
                    st.error(
                        "This Unit Code already exists. "
                        "Please enter a different Unit Code."
                    )
                    return

                updated_df = pd.concat(
                    [
                        df,
                        pd.DataFrame([new_row], columns=columns),
                    ],
                    ignore_index=True,
                )

                try:
                    save_dataframe_to_excel(updated_df)
                    st.success("New unit added successfully.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Could not add the unit: {e}")

        with cancel_col:
            if st.button("Cancel", use_container_width=True):
                st.rerun()

    dialog()


def show_delete_dialog(row_index):
    unit_code = str(df.loc[row_index, "Unit Code"])
    phase = str(df.loc[row_index, "Phase"])
    building = str(df.loc[row_index, "Building"])

    @st.dialog("Delete Unit")
    def dialog():
        st.warning(
            f"Are you sure you want to delete Unit Code '{unit_code}' "
            f"from {phase} / Building {building}?"
        )

        st.caption(
            "This removes only the selected record after confirmation."
        )

        yes_col, no_col = st.columns(2)

        with yes_col:
            if st.button(
                "Delete",
                type="primary",
                use_container_width=True,
            ):
                updated_df = df.drop(index=row_index).reset_index(drop=True)

                try:
                    save_dataframe_to_excel(updated_df)
                    st.success("Unit deleted successfully.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Could not delete the unit: {e}")

        with no_col:
            if st.button("Cancel", use_container_width=True):
                st.rerun()

    dialog()

# =========================================================
# DATA MANAGEMENT — FAST & EASY
# =========================================================
st.divider()

st.markdown(
    '<div class="section-title">Unit Data</div>',
    unsafe_allow_html=True,
)

st.caption(
    "Use the small search below to find a unit, then choose Edit or Delete."
)

display_cols = [
    c for c in [
        "Phase", "Unit Code", "Unit Type", "Building", "Floor",
        "Apartment NO.", "Type", "In/Area", "NEW M.PRICE",
        "STATUS", "Rooms Num", "Name of client"
    ]
    if c in filtered.columns
]

st.dataframe(
    filtered[display_cols],
    use_container_width=True,
    hide_index=True,
    height=430,
)

# ---------------------------------------------------------
# Compact search + actions
# ---------------------------------------------------------
st.markdown("**Unit actions**")

search_col, selected_col, edit_col, delete_col, add_col, export_col = st.columns(
    [1.35, 2.1, 0.55, 0.55, 0.55, 1.1]
)

with search_col:
    action_search = st.text_input(
        "Search",
        value="",
        placeholder="Unit code / client",
        key="action_search",
        label_visibility="collapsed",
    )

# Search from the already-filtered data so the dashboard filters
# continue to control what can be selected.
action_pool = filtered.copy()

if action_search.strip():
    q = action_search.strip().lower()

    unit_match = (
        action_pool["Unit Code"]
        .fillna("")
        .astype(str)
        .str.lower()
        .str.contains(q, regex=False)
    )

    client_match = (
        action_pool["Name of client"]
        .fillna("")
        .astype(str)
        .str.lower()
        .str.contains(q, regex=False)
    )

    action_pool = action_pool[unit_match | client_match]

with selected_col:
    if len(action_pool) > 0:

        def format_action_unit(idx):
            row = df.loc[idx]
            return (
                f"{row['Unit Code']} | "
                f"{row['Phase']} | "
                f"{row['Building']} | "
                f"{row['STATUS']}"
            )

        selected_row = st.selectbox(
            "Unit",
            action_pool.index.tolist(),
            format_func=format_action_unit,
            key="action_selected_unit",
            label_visibility="collapsed",
        )
    else:
        selected_row = None
        st.caption("No matching unit")

with edit_col:
    if st.button(
        "✏️",
        key="quick_edit",
        disabled=selected_row is None,
        use_container_width=True,
        help="Edit selected unit",
    ):
        show_edit_dialog(selected_row)

with delete_col:
    if st.button(
        "🗑️",
        key="quick_delete",
        disabled=selected_row is None,
        use_container_width=True,
        help="Delete selected unit",
    ):
        show_delete_dialog(selected_row)

with add_col:
    if st.button(
        "➕",
        key="quick_add",
        use_container_width=True,
        help="Add new unit",
    ):
        show_add_dialog()

with export_col:
    export_bytes = make_excel_download(filtered.copy())

    st.download_button(
        "📥 Export",
        data=export_bytes,
        file_name="Gardenia_Town_Filtered.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        help="Export the currently filtered data",
    )
