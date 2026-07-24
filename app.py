import io
import re
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Huawei Inverter Combiner",
    page_icon="⚡",
    layout="wide"
)

st.title("⚡ Huawei Inverter Data Combiner")

st.write(
    "Upload all Huawei inverter Excel files. "
    "The output will contain all data in one sheet, ordered TX 1 → TX 4."
)

# ---------------------------------------------------------
# FILE UPLOAD
# ---------------------------------------------------------

files = st.file_uploader(
    "Upload inverter Excel files",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

date_order = st.radio(
    "Date/time order within each transformer",
    ["Newest → Oldest", "Oldest → Newest"],
    horizontal=True
)


# ---------------------------------------------------------
# FUNCTIONS
# ---------------------------------------------------------

def tx_number(value):
    """
    Extract TX number from Site Name.

    Examples:
    Rolleston TX 1 -> 1
    Rolleston TX 2 -> 2
    Rolleston TX 3 -> 3
    Rolleston TX 4 -> 4
    """

    match = re.search(
        r"\bTX\s*[-_]?\s*(\d+)\b",
        str(value),
        re.IGNORECASE
    )

    if match:
        return int(match.group(1))

    return 999999


def parse_time(value):
    """
    Convert Huawei Start Time into datetime.
    Removes DST text if present.
    """

    if pd.isna(value):
        return pd.NaT

    value = str(value).strip()

    value = re.sub(
        r"\s+DST\s*$",
        "",
        value,
        flags=re.IGNORECASE
    )

    return pd.to_datetime(
        value,
        errors="coerce"
    )


def get_inverter(value):
    """
    Extract inverter ID.

    Example:
    Inverter(MBUS-22) -> MBUS-22
    """

    text = str(value)

    match = re.search(
        r"Inverter\(([^)]+)\)",
        text,
        re.IGNORECASE
    )

    if match:
        return match.group(1)

    return text


def read_file(file):

    # Huawei Excel export has headings on Excel row 4
    df = pd.read_excel(
        file,
        sheet_name=0,
        header=3
    )

    # Remove completely empty rows and columns
    df = df.dropna(
        axis=0,
        how="all"
    )

    df = df.dropna(
        axis=1,
        how="all"
    )

    # Required Huawei columns
    required_columns = [
        "Site Name",
        "ManageObject",
        "Start Time"
    ]

    missing_columns = [
        column
        for column in required_columns
        if column not in df.columns
    ]

    if missing_columns:
        raise ValueError(
            "Missing column(s): "
            + ", ".join(missing_columns)
        )

    # Extract inverter name
    df["Inverter"] = df["ManageObject"].apply(
        get_inverter
    )

    # Transformer number used for sorting
    df["_TX"] = df["Site Name"].apply(
        tx_number
    )

    # Datetime used for sorting
    df["_TIME"] = df["Start Time"].apply(
        parse_time
    )

    # Keep original filename for reference
    df["Source File"] = file.name

    return df


# ---------------------------------------------------------
# PROCESS FILES
# ---------------------------------------------------------

if files:

    st.write(
        f"**{len(files)} files selected**"
    )

    if st.button(
        "Combine Files",
        type="primary",
        use_container_width=True
    ):

        data = []
        errors = []

        progress = st.progress(0)

        for index, file in enumerate(files):

            try:

                df = read_file(file)

                data.append(df)

            except Exception as error:

                errors.append(
                    f"{file.name}: {error}"
                )

            progress.progress(
                (index + 1) / len(files)
            )

        # -------------------------------------------------
        # COMBINE DATA
        # -------------------------------------------------

        if data:

            combined = pd.concat(
                data,
                ignore_index=True
            )

            # User selected time sorting
            time_ascending = (
                date_order == "Oldest → Newest"
            )

            # Sort:
            # TX1
            # TX2
            # TX3
            # TX4
            #
            # Then sort date/time within each TX

            combined = combined.sort_values(
                by=[
                    "_TX",
                    "_TIME",
                    "Inverter"
                ],
                ascending=[
                    True,
                    time_ascending,
                    True
                ],
                na_position="last"
            )

            combined = combined.reset_index(
                drop=True
            )

            # Replace original Start Time with
            # clean Excel datetime
            combined["Start Time"] = combined["_TIME"]

            # -------------------------------------------------
            # COLUMN ORDER
            # -------------------------------------------------

            first_columns = [
                "Site Name",
                "Inverter",
                "ManageObject",
                "Start Time"
            ]

            remaining_columns = [
                column
                for column in combined.columns
                if column not in first_columns
                and column not in [
                    "_TX",
                    "_TIME"
                ]
            ]

            output_df = combined[
                first_columns
                + remaining_columns
            ]

            # -------------------------------------------------
            # CREATE EXCEL FILE
            # -------------------------------------------------

            output = io.BytesIO()

            with pd.ExcelWriter(
                output,
                engine="openpyxl"
            ) as writer:

                output_df.to_excel(
                    writer,
                    index=False,
                    sheet_name="Combined Data"
                )

                worksheet = writer.book[
                    "Combined Data"
                ]

                # Freeze top row
                worksheet.freeze_panes = "A2"

                # Enable Excel filters
                worksheet.auto_filter.ref = (
                    worksheet.dimensions
                )

                # Date format
                start_time_column = (
                    output_df.columns.get_loc(
                        "Start Time"
                    )
                    + 1
                )

                for row in range(
                    2,
                    worksheet.max_row + 1
                ):

                    worksheet.cell(
                        row=row,
                        column=start_time_column
                    ).number_format = (
                        "dd/mm/yyyy hh:mm"
                    )

            output.seek(0)

            # -------------------------------------------------
            # RESULTS
            # -------------------------------------------------

            valid_transformers = sorted(
                [
                    tx
                    for tx in combined["_TX"]
                    .dropna()
                    .unique()
                    if tx != 999999
                ]
            )

            transformer_text = ", ".join(
                f"TX {int(tx)}"
                for tx in valid_transformers
            )

            st.success(
                f"Done! "
                f"{len(data)} files combined | "
                f"{len(output_df):,} rows | "
                f"{transformer_text}"
            )

            # -------------------------------------------------
            # SUMMARY
            # -------------------------------------------------

            st.subheader(
                "Transformer Summary"
            )

            summary = (
                output_df
                .groupby("Site Name")
                .agg(
                    Rows=(
                        "Site Name",
                        "size"
                    ),
                    Inverters=(
                        "Inverter",
                        "nunique"
                    ),
                    First_Time=(
                        "Start Time",
                        "min"
                    ),
                    Last_Time=(
                        "Start Time",
                        "max"
                    )
                )
                .reset_index()
            )

            summary["_TX"] = (
                summary["Site Name"]
                .apply(tx_number)
            )

            summary = (
                summary
                .sort_values("_TX")
                .drop(
                    columns="_TX"
                )
            )

            st.dataframe(
                summary,
                use_container_width=True,
                hide_index=True
            )

            # -------------------------------------------------
            # PREVIEW
            # -------------------------------------------------

            st.subheader(
                "Combined Data Preview"
            )

            st.dataframe(
                output_df.head(100),
                use_container_width=True,
                hide_index=True
            )

            # -------------------------------------------------
            # DOWNLOAD
            # -------------------------------------------------

            st.download_button(
                label="⬇️ Download Combined Excel",
                data=output.getvalue(),
                file_name=(
                    "Combined_Inverter_Data_"
                    "TX1_to_TX4.xlsx"
                ),
                mime=(
                    "application/vnd.openxmlformats-"
                    "officedocument.spreadsheetml.sheet"
                ),
                use_container_width=True
            )

        # -------------------------------------------------
        # ERRORS
        # -------------------------------------------------

        if errors:

            st.warning(
                "Some files could not be processed:"
            )

            for error in errors:

                st.write(
                    "• " + error
                )

else:

    st.info(
        "Upload your inverter Excel files to begin."
    )
