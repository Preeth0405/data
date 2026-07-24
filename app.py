import io
import re
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Huawei Inverter Data Combiner",
    page_icon="⚡",
    layout="wide"
)

st.title("⚡ Huawei Inverter Data Combiner")

st.write(
    "Upload all inverter Excel files. The app creates one row per timestamp "
    "and one column per inverter."
)

# ---------------------------------------------------------
# FILE UPLOAD
# ---------------------------------------------------------

files = st.file_uploader(
    "Upload Huawei inverter Excel files",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

date_order = st.radio(
    "Date/time order",
    ["Oldest → Newest", "Newest → Oldest"],
    horizontal=True
)


# ---------------------------------------------------------
# FUNCTIONS
# ---------------------------------------------------------

def tx_number(value):
    """
    Extract TX number from Site Name.

    Example:
    Rolleston TX 1 -> 1
    Rolleston TX 4 -> 4
    """

    match = re.search(
        r"\bTX\s*[-_]?\s*(\d+)\b",
        str(value),
        re.IGNORECASE
    )

    return int(match.group(1)) if match else 999


def inverter_number(value):
    """
    Extract inverter number for correct sorting.

    Example:
    Inverter(MBUS-22) -> 22
    MBUS-5 -> 5
    """

    text = str(value)

    match = re.search(
        r"MBUS[-_\s]*(\d+)",
        text,
        re.IGNORECASE
    )

    if match:
        return int(match.group(1))

    numbers = re.findall(r"\d+", text)

    return int(numbers[-1]) if numbers else 999999


def inverter_name(value):
    """
    Convert:
    Inverter(MBUS-22)

    to:
    MBUS-22
    """

    text = str(value)

    match = re.search(
        r"Inverter\(([^)]+)\)",
        text,
        re.IGNORECASE
    )

    return match.group(1) if match else text


def parse_time(value):

    if pd.isna(value):
        return pd.NaT

    text = str(value).strip()

    # Remove Huawei DST suffix
    text = re.sub(
        r"\s+DST\s*$",
        "",
        text,
        flags=re.IGNORECASE
    )

    return pd.to_datetime(
        text,
        errors="coerce"
    )


def read_huawei_file(file):

    # Huawei headings are on Excel row 4
    df = pd.read_excel(
        file,
        sheet_name=0,
        header=3
    )

    # Remove empty rows / columns
    df = df.dropna(
        axis=0,
        how="all"
    )

    df = df.dropna(
        axis=1,
        how="all"
    )

    required = [
        "Site Name",
        "ManageObject",
        "Start Time"
    ]

    missing = [
        col for col in required
        if col not in df.columns
    ]

    if missing:

        raise ValueError(
            "Missing columns: "
            + ", ".join(missing)
        )

    # Clean timestamp
    df["Start Time"] = (
        df["Start Time"]
        .apply(parse_time)
    )

    # Get inverter name
    df["Inverter"] = (
        df["ManageObject"]
        .apply(inverter_name)
    )

    # Get TX number
    df["TX"] = (
        df["Site Name"]
        .apply(tx_number)
    )

    # Inverter numerical order
    df["INV_ORDER"] = (
        df["Inverter"]
        .apply(inverter_number)
    )

    return df


# ---------------------------------------------------------
# PROCESS
# ---------------------------------------------------------

if files:

    st.info(
        f"{len(files)} inverter files selected."
    )

    if st.button(
        "Read & Combine Files",
        type="primary",
        use_container_width=True
    ):

        all_data = []
        errors = []

        progress = st.progress(0)

        for i, file in enumerate(files):

            try:

                df = read_huawei_file(file)

                all_data.append(df)

            except Exception as error:

                errors.append(
                    f"{file.name}: {error}"
                )

            progress.progress(
                (i + 1) / len(files)
            )

        # -------------------------------------------------
        # COMBINE
        # -------------------------------------------------

        if all_data:

            combined = pd.concat(
                all_data,
                ignore_index=True
            )

            # ---------------------------------------------
            # FIND AVAILABLE MEASUREMENTS
            # ---------------------------------------------

            exclude_columns = [
                "Site Name",
                "Management Domain",
                "ManageObject",
                "Start Time",
                "Inverter",
                "TX",
                "INV_ORDER"
            ]

            measurement_columns = [
                col
                for col in combined.columns
                if col not in exclude_columns
            ]

            # Save data into session state
            st.session_state["combined"] = combined
            st.session_state["measurements"] = measurement_columns

            st.success(
                f"{len(all_data)} files successfully loaded."
            )

        if errors:

            st.warning(
                "Some files could not be processed:"
            )

            for error in errors:

                st.write("• " + error)


# ---------------------------------------------------------
# MEASUREMENT SELECTION
# ---------------------------------------------------------

if "combined" in st.session_state:

    combined = st.session_state["combined"]

    measurements = (
        st.session_state["measurements"]
    )

    st.divider()

    st.subheader(
        "Select Measurement"
    )

    measurement = st.selectbox(
        "Data to export",
        measurements
    )

    # -----------------------------------------------------
    # INVERTER INFORMATION
    # -----------------------------------------------------

    inverter_info = (
        combined[
            [
                "Inverter",
                "TX",
                "INV_ORDER"
            ]
        ]
        .drop_duplicates()
        .sort_values(
            [
                "TX",
                "INV_ORDER"
            ]
        )
    )

    inverter_list = (
        inverter_info["Inverter"]
        .tolist()
    )

    # -----------------------------------------------------
    # CREATE TIME × INVERTER TABLE
    # -----------------------------------------------------

    selected = combined[
        [
            "Start Time",
            "Inverter",
            measurement
        ]
    ].copy()

    # Convert measurement to numeric
    selected[measurement] = (
        pd.to_numeric(
            selected[measurement],
            errors="coerce"
        )
    )

    # Pivot
    pivot = selected.pivot_table(
        index="Start Time",
        columns="Inverter",
        values=measurement,
        aggfunc="first"
    )

    # -----------------------------------------------------
    # ORDER INVERTERS TX1 → TX4
    # -----------------------------------------------------

    available_inverters = [
        inv
        for inv in inverter_list
        if inv in pivot.columns
    ]

    pivot = pivot.reindex(
        columns=available_inverters
    )

    # -----------------------------------------------------
    # RENAME COLUMNS INV1 → INV40
    # -----------------------------------------------------

    rename_map = {}

    for index, inverter in enumerate(
        available_inverters,
        start=1
    ):

        rename_map[inverter] = (
            f"INV {index}"
        )

    pivot = pivot.rename(
        columns=rename_map
    )

    # Start Time becomes normal column
    pivot = pivot.reset_index()

    # -----------------------------------------------------
    # DATE SORT
    # -----------------------------------------------------

    ascending = (
        date_order == "Oldest → Newest"
    )

    pivot = pivot.sort_values(
        "Start Time",
        ascending=ascending
    )

    # -----------------------------------------------------
    # SHOW INVERTER MAPPING
    # -----------------------------------------------------

    st.subheader(
        "Inverter Mapping"
    )

    mapping_data = []

    for index, row in (
        inverter_info
        .reset_index(drop=True)
        .iterrows()
    ):

        mapping_data.append(
            {
                "Output Column":
                    f"INV {index + 1}",

                "Actual Inverter":
                    row["Inverter"],

                "Transformer":
                    f"TX {int(row['TX'])}"
            }
        )

    mapping_df = pd.DataFrame(
        mapping_data
    )

    st.dataframe(
        mapping_df,
        use_container_width=True,
        hide_index=True
    )

    # -----------------------------------------------------
    # DATA PREVIEW
    # -----------------------------------------------------

    st.subheader(
        f"{measurement} Preview"
    )

    st.caption(
        f"{len(pivot):,} timestamps × "
        f"{len(available_inverters)} inverters"
    )

    st.dataframe(
        pivot.head(100),
        use_container_width=True,
        hide_index=True
    )

    # -----------------------------------------------------
    # CREATE EXCEL
    # -----------------------------------------------------

    output = io.BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        # Main data
        pivot.to_excel(
            writer,
            index=False,
            sheet_name="Inverter Data"
        )

        # Mapping
        mapping_df.to_excel(
            writer,
            index=False,
            sheet_name="Inverter Mapping"
        )

        # ---------------------------------------------
        # FORMAT MAIN SHEET
        # ---------------------------------------------

        worksheet = writer.book[
            "Inverter Data"
        ]

        worksheet.freeze_panes = "B2"

        worksheet.auto_filter.ref = (
            worksheet.dimensions
        )

        # Time column width
        worksheet.column_dimensions[
            "A"
        ].width = 22

        # Format timestamp
        for row in range(
            2,
            worksheet.max_row + 1
        ):

            worksheet.cell(
                row=row,
                column=1
            ).number_format = (
                "dd/mm/yyyy hh:mm"
            )

        # Inverter column widths
        for column in range(
            2,
            worksheet.max_column + 1
        ):

            worksheet.column_dimensions[
                worksheet.cell(
                    row=1,
                    column=column
                ).column_letter
            ].width = 14

    output.seek(0)

    # -----------------------------------------------------
    # DOWNLOAD
    # -----------------------------------------------------

    safe_measurement = re.sub(
        r"[^A-Za-z0-9]+",
        "_",
        measurement
    ).strip("_")

    st.download_button(
        "⬇️ Download Excel",
        data=output.getvalue(),
        file_name=(
            f"Inverter_{safe_measurement}.xlsx"
        ),
        mime=(
            "application/"
            "vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        type="primary",
        use_container_width=True
    )
