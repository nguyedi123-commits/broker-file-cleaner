from __future__ import annotations

from datetime import datetime

import pandas as pd
import streamlit as st

from broker_file_cleaning.process_templates import process_uploaded_files

MAX_FILES = 50
MAX_TOTAL_MB = 300
MAX_SINGLE_MB = 100
ALLOWED_TYPES = ["csv", "xlsx", "xlsb", "xls"]


def _mb(num_bytes: int) -> float:
    return num_bytes / (1024 * 1024)


def _validate_uploads(files: list[object]) -> list[str]:
    errors: list[str] = []
    if not files:
        return ["Upload at least one file to begin."]

    if len(files) > MAX_FILES:
        errors.append(f"Too many files: uploaded {len(files)} files, limit is {MAX_FILES}.")

    total_bytes = sum(getattr(file, "size", 0) for file in files)
    if total_bytes > MAX_TOTAL_MB * 1024 * 1024:
        errors.append(
            f"Batch too large: {_mb(total_bytes):.1f} MB uploaded, limit is {MAX_TOTAL_MB} MB total."
        )

    oversized = [
        f"{file.name} ({_mb(file.size):.1f} MB)"
        for file in files
        if getattr(file, "size", 0) > MAX_SINGLE_MB * 1024 * 1024
    ]
    if oversized:
        errors.append(
            f"These files exceed the {MAX_SINGLE_MB} MB per-file limit: {', '.join(oversized)}."
        )

    return errors


st.set_page_config(page_title="Broker File Cleaner", layout="wide")

st.title("Broker File Cleaner")
st.caption(
    "Upload raw broker meter files, clean them into one standard CSV, and download the result."
)
st.caption(
    "Energy-only rule: this cleaner only processes interval data in kWh or MWh. "
    "Reactive or non-energy units such as kVArh are ignored."
)

with st.sidebar:
    st.subheader("Limits")
    st.write(f"Files per upload: `{MAX_FILES}`")
    st.write(f"Total batch size: `{MAX_TOTAL_MB} MB`")
    st.write(f"Single file size: `{MAX_SINGLE_MB} MB`")
    st.write("Supported types: `.csv`, `.xlsx`, `.xlsb`, `.xls`")
    st.write("Supported energy units: `kWh`, `MWh`")
    st.write("Ignored units: `kVArh` and other non-energy units")

uploaded_files = st.file_uploader(
    "Upload raw files",
    type=ALLOWED_TYPES,
    accept_multiple_files=True,
    help="You can upload up to 50 files in one batch.",
)

if uploaded_files:
    total_bytes = sum(file.size for file in uploaded_files)
    st.info(
        f"{len(uploaded_files)} file(s) selected, {_mb(total_bytes):.1f} MB total."
    )

validation_errors = _validate_uploads(uploaded_files or [])
for error in validation_errors:
    st.error(error)

process_clicked = st.button("Process files", type="primary", use_container_width=True)

if process_clicked and not validation_errors:
    with st.spinner("Cleaning uploaded files..."):
        result = process_uploaded_files(uploaded_files)

    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("Uploaded", len(uploaded_files))
    col2.metric("Processed", len(result.processed_files))
    col3.metric("Failed", len(result.failed_files))
    col4.metric("Rows", f"{result.total_rows:,}")
    col5.metric("Unique NMIs", f"{result.unique_nmis:,}")

    if result.warnings:
        for warning in result.warnings:
            st.warning(warning)

    if result.output_csv_bytes and result.output_filename:
        st.success("Cleaning complete. Download the merged CSV below.")
        st.download_button(
            "Download cleaned CSV",
            data=result.output_csv_bytes,
            file_name=result.output_filename,
            mime="text/csv",
            use_container_width=True,
        )
    else:
        st.error("No cleanable data was produced from this batch.")

    if result.processed_files:
        st.subheader("Processed files")
        processed_df = pd.DataFrame(
            [
                {
                    "filename": item.filename,
                    "rows": item.rows,
                    "format": item.format_id,
                    "nmi_preview": item.nmi_preview,
                    "aggregated_to_30min": item.aggregated_to_30min,
                }
                for item in result.processed_files
            ]
        )
        st.dataframe(processed_df, use_container_width=True, hide_index=True)

    if result.failed_files:
        st.subheader("Failed files")
        failed_df = pd.DataFrame(
            [
                {
                    "filename": item.filename,
                    "status": item.status,
                    "reason": item.reason,
                }
                for item in result.failed_files
            ]
        )
        st.dataframe(failed_df, use_container_width=True, hide_index=True)

    st.caption(f"Run completed at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}.")
