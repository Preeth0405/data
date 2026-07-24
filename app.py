from pathlib import Path

app_code = r'''import io
import re
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Huawei Inverter Data Combiner", page_icon="⚡", layout="wide")

st.title("⚡ Huawei Inverter Data Combiner")
st.write(
    "Upload all inverter Excel exports. The app combines them into one Excel sheet, "
    "groups them TX 1 → TX 4, and sorts the timestamps within each transformer."
)

uploaded_files = st.file_uploader(
    "Upload inverter Excel files",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

sort_direction = st.radio(
    "Date order within each transformer",
    ["Newest → Oldest", "Oldest → Newest"],
    horizontal=True
)

def tx_number(site_name):
    """Extract transformer number from values such as 'Rolleston TX 2'."""
    match = re.search(r"\bTX\s*[-_]?\s*(\d+)\b", str(site_name), flags=re.IGNORECASE)
    return int(match.group(1)) if match else 999999

def clean_time(value):
    """Convert Huawei timestamp such as '2026-07-08 00:00:00 DST' to datetime."""
    if pd.isna(value):
        return pd.NaT
    text = str(value).strip()
    text = re.sub(r"\s+DST\s*$", "", text, flags=re.IGNORECASE)
    return pd.to_datetime(text, errors="coerce")

def inverter_name(value):
    """Extract Inverter(MBUS-22) / MBUS-22 style identifier."""
    text = str(value)
    match = re.search(r"Inverter\(([^)]+)\)", text, flags=re.IGNORECASE)
    return match.group(1) if match else text

def read_huawei_file(uploaded_file):
    # Huawei export has metadata in rows 1-3 and column names on row 4.
    df = pd.read_excel(uploaded_file, sheet_name=0, header=3)

    # Remove completely empty columns/rows.
    df = df.dropna(axis=1, how="all").dropna(axis=0, how="all")

    required = {"Site Name", "ManageObject", "Start Time"}
    missing = required.difference(df.columns)
    if missing:
        raise ValueError(
            "Could not find expected columns: " + ", ".join(sorted(missing))
        )

    df["Source File"] = uploaded_file.name
    df["Inverter"] = df["ManageObject"].apply(inverter_name)
    df["_TX_Order"] = df["Site Name"].apply(tx_number)
    df["_Start_Time_Sort"] = df["Start Time"].apply(clean_time)

    return df

if uploaded_files:
    st.info(f"{len(uploaded_files)} file(s) selected.")

    if st.button("Combine Files", type="primary", use_container_width=True):
        frames = []
        errors = []

        progress = st.progress(0)

        for i, uploaded_file in enumerate(uploaded_files):
            try:
                frames.append(read_huawei_file(uploaded_file))
            except Exception as exc:
                errors.append(f"{uploaded_file.name}: {exc}")

            progress.progress((i + 1) / len(uploaded_files))

        if frames:
            combined = pd.concat(frames, ignore_index=True)

            ascending_time = sort_direction == "Oldest → Newest"

            # Always TX 1, TX 2, TX 3, TX 4...
            # Timestamp sorting is selected by the user.
            combined = combined.sort_values(
                by=["_TX_Order", "_Start_Time_Sort", "Inverter"],
                ascending=[True, ascending_time, True],
                na_position="last"
            ).reset_index(drop=True)

            # Replace Huawei DST strings with proper Excel datetimes.
            combined["Start Time"] = combined["_Start_Time_Sort"]

            # Put the most useful identification columns first.
            preferred = ["Site Name", "Inverter", "ManageObject", "Start Time"]
            remaining = [
                c for c in combined.columns
                if c not in preferred + ["_TX_Order", "_Start_Time_Sort"]
            ]
            output_df = combined[preferred + remaining]

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                output_df.to_excel(
                    writer,
                    index=False,
                    sheet_name="Combined Inverter Data"
                )

                ws = writer.book["Combined Inverter Data"]
                ws.freeze_panes = "A2"
                ws.auto_filter.ref = ws.dimensions

                # Sensible widths without making huge data columns excessively wide.
                for column_cells in ws.columns:
                    letter = column_cells[0].column_letter
                    max_length = 0
                    for cell in column_cells[:200]:
                        if cell.value is not None:
                            max_length = max(max_length, len(str(cell.value)))
                    ws.column_dimensions[letter].width = min(max(max_length + 2, 12), 35)

            output.seek(0)

            tx_summary = (
                output_df.groupby("Site Name", dropna=False)
                .agg(
                    Rows=("Site Name", "size"),
                    Inverters=("Inverter", "nunique"),
                    First_Time=("Start Time", "min"),
                    Last_Time=("Start Time", "max"),
                )
                .reset_index()
            )
            tx_summary["_order"] = tx_summary["Site Name"].apply(tx_number)
            tx_summary = tx_summary.sort_values("_order").drop(columns="_order")

            st.success(
                f"Combined {len(frames)} file(s) into {len(output_df):,} rows."
            )
            st.dataframe(tx_summary, use_container_width=True, hide_index=True)

            st.subheader("Preview")
            st.dataframe(output_df.head(100), use_container_width=True)

            st.download_button(
                "⬇️ Download Combined Excel",
                data=output.getvalue(),
                file_name="Combined_Inverter_Data_TX1_to_TX4.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        if errors:
            st.warning("Some files could not be processed:")
            for error in errors:
                st.write("•", error)

else:
    st.caption(
        "Expected Huawei format: row 4 contains Site Name, Management Domain, "
        "ManageObject, Start Time, Active power, Daily energy, etc."
    )
'''

requirements = """streamlit
pandas
openpyxl
xlrd
"""

app_path = Path("/mnt/data/huawei_inverter_combiner.py")
req_path = Path("/mnt/data/requirements.txt")
app_path.write_text(app_code, encoding="utf-8")
req_path.write_text(requirements, encoding="utf-8")

print(f"Created {app_path.name}")
print(f"Created {req_path.name}")
