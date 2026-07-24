import io
import re

import pandas as pd
import streamlit as st


# =========================================================
# PAGE SETUP
# =========================================================

st.set_page_config(
    page_title="Huawei Inverter Data Combiner",
    page_icon="⚡",
    layout="wide",
)

st.title("⚡ Huawei Inverter Data Combiner")

st.write(
    "Upload all Huawei inverter Excel files. "
    "The app creates one row per timestamp and one column per inverter."
)


# =========================================================
# FILE UPLOAD
# =========================================================

files = st.file_uploader(
    "Upload Huawei inverter Excel files",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
)

date_order = st.radio(
    "Date/time order",
    ["Oldest → Newest", "Newest → Oldest"],
    horizontal=True,
)


# =========================================================
# HELPER FUNCTIONS
# =========================================================

def tx_number(value):
    """
    Extract transformer number.

    Example:
    Rolleston TX 1 -> 1
    Rolleston TX 2 -> 2
    """

    text = str(value)

    match = re.search(
        r"\bTX\s*[-_]?\s*(\d+)\b",
        text,
        re.IGNORECASE,
    )

    if match:
        return int(match.group(1))

    return 999999


def inverter_number(value):
    """
    Extract numerical inverter/MBUS number.

    Examples:
    MBUS-22 -> 22
    Inverter(MBUS-22) -> 22
    """

    text = str(value)

    match = re.search(
        r"MBUS[-_\s]*(\d+)",
        text,
        re.IGNORECASE,
    )

    if match:
        return int(match.group(1))

    numbers = re.findall(r"\d+", text)

    if numbers:
        return int(numbers[-1])

    return 999999


def inverter_name(value):
    """
    Convert:
    Inverter(MBUS-22)

    into:
    MBUS-22
    """

    text = str(value).strip()

    match = re.search(
        r"Inverter\(([^)]+)\)",
        text,
        re.IGNORECASE,
    )

    if match:
        return match.group(1).strip()

    return text


def parse_time(value):
    """
    Convert Huawei timestamp into pandas datetime.
    Also removes DST suffix.
    """

    if pd.isna(value):
        return pd.NaT

    text = str(value).strip()

    text = re.sub(
        r"\s+DST\s*$",
        "",
        text,
        flags=re.IGNORECASE,
    )

    return pd.to_datetime(
        text,
        errors="coerce",
    )


def read_huawei_file(file):
    """
    Read one Huawei inverter Excel export.
    Huawei headings are expected on Excel row 4.
    """

    df = pd.read_excel(
        file,
        sheet_name=0,
        header=3,
    )

    # Clean column names
    df.columns = [
        str(col).strip()
        for col in df.columns
    ]

    # Remove empty rows and columns
    df = df.dropna(
        axis=0,
        how="all",
    )

    df = df.dropna(
        axis=1,
        how="all",
    )

    required_columns = [
        "Site Name",
        "ManageObject",
        "Start Time",
    ]

    missing = [
        col
        for col in required_columns
        if col not in df.columns
    ]

    if missing:
        raise ValueError(
            "Missing required column(s): "
            + ", ".join(missing)
        )

    # Clean timestamp
    df["Start Time"] = (
        df["Start Time"]
        .apply(parse_time)
    )

    # Remove invalid timestamps
    df = df[
        df["Start Time"].notna()
    ].copy()

    # Actual inverter name
    df["Inverter"] = (
        df["ManageObject"]
        .apply(inverter_name)
    )

    # Transformer number
    df["TX"] = (
        df["Site Name"]
        .apply(tx_number)
    )

    # Inverter sorting number
    df["INV_ORDER"] = (
        df["Inverter"]
        .apply(inverter_number)
    )

    # Source file
    df["Source File"] = file.name

    # IMPORTANT:
    # Same MBUS number can potentially exist under another TX.
    # Therefore TX + inverter creates a unique inverter identity.
    df["Unique_Inverter"] = (
        "TX"
        + df["TX"].astype(str)
        + "_"
        + df["Inverter"].astype(str)
    )

    return df


# =========================================================
# READ FILES
# =========================================================

if files:

    st.info(
        f"{len(files)} files selected."
    )

    if st.button(
        "Read & Combine Files",
        type="primary",
        use_container_width=True,
    ):

        all_data = []
        errors = []

        progress = st.progress(0)

        for index, file in enumerate(files):

            try:

                df = read_huawei_file(file)

                all_data.append(df)

            except Exception as error:

                errors.append(
                    f"{file.name}: {error}"
                )

            progress.progress(
                (index + 1) / len(files)
            )

        # ---------------------------------------------
        # COMBINE ALL FILES
        # ---------------------------------------------

        if all_data:

            combined = pd.concat(
                all_data,
                ignore_index=True,
            )

            # -----------------------------------------
            # FIND MEASUREMENT COLUMNS
            # -----------------------------------------

            excluded_columns = {
                "Site Name",
                "Management Domain",
                "ManageObject",
                "Start Time",
                "Inverter",
                "TX",
                "INV_ORDER",
                "Source File",
                "Unique_Inverter",
            }

            measurement_columns = [
                col
                for col in combined.columns
                if col not in excluded_columns
            ]

            # Prefer columns containing numeric data
            usable_measurements = []

            for column in measurement_columns:

                numeric_test = pd.to_numeric(
                    combined[column],
                    errors="coerce",
                )

                if numeric_test.notna().any():
                    usable_measurements.append(
                        column
                    )

            st.session_state[
                "combined_data"
            ] = combined

            st.session_state[
                "measurements"
            ] = usable_measurements

            st.success(
                f"{len(all_data)} files loaded successfully."
            )

        # ---------------------------------------------
        # SHOW ERRORS
        # ---------------------------------------------

        if errors:

            st.warning(
                "Some files could not be processed:"
            )

            for error in errors:
                st.write(
                    "• " + error
                )


# =========================================================
# CREATE OUTPUT
# =========================================================

if "combined_data" in st.session_state:

    combined = st.session_state[
        "combined_data"
    ]

    measurements = st.session_state[
        "measurements"
    ]

    st.divider()

    # =====================================================
    # MEASUREMENT SELECTOR
    # =====================================================

    st.subheader(
        "1. Select Measurement"
    )

    if not measurements:

        st.error(
            "No numeric measurement columns were found."
        )

        st.stop()

    measurement = st.selectbox(
        "Measurement to export",
        measurements,
    )

    # =====================================================
    # CREATE UNIQUE INVERTER LIST
    # =====================================================

    inverter_info = (
        combined[
            [
                "Unique_Inverter",
                "Inverter",
                "TX",
                "INV_ORDER",
                "Site Name",
            ]
        ]
        .drop_duplicates(
            subset=["Unique_Inverter"]
        )
        .sort_values(
            by=[
                "TX",
                "INV_ORDER",
                "Unique_Inverter",
            ]
        )
        .reset_index(drop=True)
    )

    # Assign output names
    # INV 1, INV 2 ... INV 40

    inverter_info[
        "Output Column"
    ] = [
        f"INV {i}"
        for i in range(
            1,
            len(inverter_info) + 1,
        )
    ]

    # =====================================================
    # SELECT MEASUREMENT DATA
    # =====================================================

    selected = combined[
        [
            "Start Time",
            "Unique_Inverter",
            measurement,
        ]
    ].copy()

    # Convert measurement to numeric
    selected[measurement] = pd.to_numeric(
        selected[measurement],
        errors="coerce",
    )

    # =====================================================
    # PIVOT
    # =====================================================

    pivot = selected.pivot_table(
        index="Start Time",
        columns="Unique_Inverter",
        values=measurement,
        aggfunc="first",
    )

    # =====================================================
    # ORDER INVERTERS
    # TX1 -> TX2 -> TX3 -> TX4
    # =====================================================

    ordered_unique_inverters = (
        inverter_info[
            "Unique_Inverter"
        ]
        .tolist()
    )

    # Only use inverter columns actually present
    ordered_unique_inverters = [
        inverter
        for inverter
        in ordered_unique_inverters
        if inverter in pivot.columns
    ]

    pivot = pivot.reindex(
        columns=ordered_unique_inverters
    )

    # =====================================================
    # RENAME TO INV 1 -> INV 40
    # =====================================================

    rename_map = dict(
        zip(
            inverter_info[
                "Unique_Inverter"
            ],
            inverter_info[
                "Output Column"
            ],
        )
    )

    pivot = pivot.rename(
        columns=rename_map
    )

    # Remove pandas column index name
    pivot.columns.name = None

    # Convert time index back to column
    pivot = pivot.reset_index()

    # =====================================================
    # SAFETY CHECK - DUPLICATE COLUMNS
    # =====================================================

    if pivot.columns.duplicated().any():

        duplicate_names = (
            pivot.columns[
                pivot.columns.duplicated()
            ]
            .tolist()
        )

        st.error(
            "Duplicate output columns detected: "
            + ", ".join(
                map(str, duplicate_names)
            )
        )

        st.stop()

    # =====================================================
    # SORT DATE
    # =====================================================

    ascending = (
        date_order
        == "Oldest → Newest"
    )

    pivot = (
        pivot
        .sort_values(
            by="Start Time",
            ascending=ascending,
        )
        .reset_index(drop=True)
    )

    # =====================================================
    # INVERTER MAPPING
    # =====================================================

    mapping_df = inverter_info[
        [
            "Output Column",
            "Inverter",
            "TX",
            "Site Name",
        ]
    ].copy()

    mapping_df = mapping_df.rename(
        columns={
            "Inverter":
                "Actual Inverter",
            "TX":
                "Transformer",
        }
    )

    mapping_df[
        "Transformer"
    ] = mapping_df[
        "Transformer"
    ].apply(
        lambda x:
        f"TX {int(x)}"
        if x != 999999
        else "Unknown"
    )

    # =====================================================
    # SUMMARY
    # =====================================================

    st.subheader(
        "2. Inverter Mapping"
    )

    st.write(
        f"**{len(mapping_df)} inverters detected**"
    )

    st.dataframe(
        mapping_df,
        use_container_width=True,
        hide_index=True,
    )

    # =====================================================
    # OUTPUT PREVIEW
    # =====================================================

    st.subheader(
        "3. Output Preview"
    )

    st.write(
        f"**Measurement:** {measurement}"
    )

    st.write(
        f"**{len(pivot):,} timestamps × "
        f"{len(pivot.columns) - 1} inverters**"
    )

    st.dataframe(
        pivot.head(100),
        use_container_width=True,
        hide_index=True,
    )

    # =====================================================
    # CREATE EXCEL
    # =====================================================

    output = io.BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl",
    ) as writer:

        # ---------------------------------------------
        # MAIN DATA
        # ---------------------------------------------

        pivot.to_excel(
            writer,
            index=False,
            sheet_name="Inverter Data",
        )

        # ---------------------------------------------
        # MAPPING SHEET
        # ---------------------------------------------

        mapping_df.to_excel(
            writer,
            index=False,
            sheet_name="Inverter Mapping",
        )

        # ---------------------------------------------
        # FORMAT MAIN SHEET
        # ---------------------------------------------

        worksheet = writer.book[
            "Inverter Data"
        ]

        # Freeze time column and header
        worksheet.freeze_panes = "B2"

        # Excel filter
        worksheet.auto_filter.ref = (
            worksheet.dimensions
        )

        # Time column width
        worksheet.column_dimensions[
            "A"
        ].width = 22

        # Date formatting
        for row in range(
            2,
            worksheet.max_row + 1,
        ):

            worksheet.cell(
                row=row,
                column=1,
            ).number_format = (
                "dd/mm/yyyy hh:mm"
            )

        # Inverter column widths
        for column_number in range(
            2,
            worksheet.max_column + 1,
        ):

            column_letter = (
                worksheet.cell(
                    row=1,
                    column=column_number,
                ).column_letter
            )

            worksheet.column_dimensions[
                column_letter
            ].width = 13

        # ---------------------------------------------
        # FORMAT MAPPING SHEET
        # ---------------------------------------------

        mapping_sheet = writer.book[
            "Inverter Mapping"
        ]

        mapping_sheet.freeze_panes = "A2"

        mapping_sheet.auto_filter.ref = (
            mapping_sheet.dimensions
        )

        mapping_sheet.column_dimensions[
            "A"
        ].width = 16

        mapping_sheet.column_dimensions[
            "B"
        ].width = 22

        mapping_sheet.column_dimensions[
            "C"
        ].width = 16

        mapping_sheet.column_dimensions[
            "D"
        ].width = 30

    output.seek(0)

    # =====================================================
    # DOWNLOAD
    # =====================================================

    safe_measurement = re.sub(
        r"[^A-Za-z0-9]+",
        "_",
        str(measurement),
    ).strip("_")

    st.subheader(
        "4. Download"
    )

    st.download_button(
        label="⬇️ Download Combined Excel",
        data=output.getvalue(),
        file_name=(
            f"Inverter_{safe_measurement}.xlsx"
        ),
        mime=(
            "application/vnd.openxmlformats-"
            "officedocument.spreadsheetml.sheet"
        ),
        type="primary",
        use_container_width=True,
    )


# =========================================================
# NO FILES
# =========================================================

else:

    if not files:

        st.info(
            "Upload your inverter Excel files to begin."
        )
