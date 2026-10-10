import io
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Payroll Reconciliation Checker",
    page_icon="💷",
    layout="wide",
)

DEFAULT_FIELDS = [
    "Gross Pay",
    "PAYE",
    "Employee NI",
    "Pension",
    "Other Deductions",
    "Net Pay",
]

REQUIRED_KEYS = ["Employee ID", "Employee Name"]


def clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def read_upload(uploaded_file) -> pd.DataFrame:
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(uploaded_file)
    if name.endswith((".xlsx", ".xls")):
        return pd.read_excel(uploaded_file)
    raise ValueError("Please upload a CSV or Excel file.")


def validate(df: pd.DataFrame, label: str, fields: list[str]) -> list[str]:
    errors = []
    required = REQUIRED_KEYS + fields
    missing = [c for c in required if c not in df.columns]
    if missing:
        errors.append(f"{label}: missing columns: {', '.join(missing)}")

    if "Employee ID" in df.columns:
        duplicate_ids = df["Employee ID"].astype(str).duplicated(keep=False)
        if duplicate_ids.any():
            count = duplicate_ids.sum()
            errors.append(f"{label}: {count} rows contain duplicate Employee IDs.")
    return errors


def numeric(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(0.0)


def reconcile(expected: pd.DataFrame, processed: pd.DataFrame, fields: list[str], tolerance: float):
    expected = expected.copy()
    processed = processed.copy()

    expected["Employee ID"] = expected["Employee ID"].astype(str).str.strip()
    processed["Employee ID"] = processed["Employee ID"].astype(str).str.strip()

    keep_expected = ["Employee ID", "Employee Name"] + fields
    keep_processed = ["Employee ID", "Employee Name"] + fields

    expected = expected[keep_expected]
    processed = processed[keep_processed]

    expected = expected.rename(
        columns={c: f"{c} Expected" for c in expected.columns if c != "Employee ID"}
    )
    processed = processed.rename(
        columns={c: f"{c} Processed" for c in processed.columns if c != "Employee ID"}
    )

    merged = expected.merge(processed, on="Employee ID", how="outer", indicator=True)

    merged["Employee Name"] = (
        merged.get("Employee Name Expected", pd.Series(index=merged.index, dtype="object"))
        .fillna(merged.get("Employee Name Processed", pd.Series(index=merged.index, dtype="object")))
    )

    status_cols = []
    diff_cols = []

    for field in fields:
        exp_col = f"{field} Expected"
        proc_col = f"{field} Processed"
        diff_col = f"{field} Difference"
        status_col = f"{field} Status"

        merged[exp_col] = numeric(merged[exp_col])
        merged[proc_col] = numeric(merged[proc_col])
        merged[diff_col] = (merged[proc_col] - merged[exp_col]).round(2)
        merged[status_col] = merged[diff_col].abs().le(tolerance).map(
            {True: "OK", False: "Mismatch"}
        )
        diff_cols.append(diff_col)
        status_cols.append(status_col)

    merged["Record Status"] = merged["_merge"].map(
        {
            "left_only": "Missing from processed payroll",
            "right_only": "Unexpected employee in processed payroll",
            "both": "Present in both",
        }
    )

    merged["Overall Status"] = "OK"
    merged.loc[merged["_merge"] != "both", "Overall Status"] = "Exception"

    any_field_mismatch = merged[status_cols].eq("Mismatch").any(axis=1)
    merged.loc[any_field_mismatch, "Overall Status"] = "Exception"

    ordered = (
        ["Employee ID", "Employee Name", "Record Status", "Overall Status"]
        + [col for field in fields for col in (
            f"{field} Expected",
            f"{field} Processed",
            f"{field} Difference",
            f"{field} Status",
        )]
    )

    return merged[ordered], diff_cols


def to_excel_bytes(full_df: pd.DataFrame, exceptions_df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        full_df.to_excel(writer, sheet_name="Reconciliation", index=False)
        exceptions_df.to_excel(writer, sheet_name="Exceptions", index=False)
    output.seek(0)
    return output.getvalue()


st.title("Payroll Reconciliation Checker")
st.caption(
    "Compare expected payroll against processed payroll, identify exceptions, "
    "and download a reconciliation report."
)

with st.expander("Required file format"):
    st.write(
        "Both files should contain **Employee ID**, **Employee Name**, and the payroll fields "
        "you want to compare. The included sample files use: Gross Pay, PAYE, Employee NI, "
        "Pension, Other Deductions, and Net Pay."
    )

left, right = st.columns(2)
with left:
    expected_file = st.file_uploader(
        "1. Expected payroll file",
        type=["csv", "xlsx", "xls"],
        key="expected",
    )
with right:
    processed_file = st.file_uploader(
        "2. Processed payroll file",
        type=["csv", "xlsx", "xls"],
        key="processed",
    )

tolerance = st.number_input(
    "Allowed difference (£)",
    min_value=0.0,
    value=0.01,
    step=0.01,
    help="A difference within this amount will be treated as OK.",
)

if expected_file and processed_file:
    try:
        expected = clean_columns(read_upload(expected_file))
        processed = clean_columns(read_upload(processed_file))
    except Exception as exc:
        st.error(f"Could not read the files: {exc}")
        st.stop()

    common_fields = [
        field for field in DEFAULT_FIELDS
        if field in expected.columns and field in processed.columns
    ]

    fields = st.multiselect(
        "Payroll fields to reconcile",
        options=common_fields,
        default=common_fields,
    )

    if not fields:
        st.warning("Select at least one payroll field.")
        st.stop()

    errors = validate(expected, "Expected payroll", fields) + validate(
        processed, "Processed payroll", fields
    )

    if errors:
        for error in errors:
            st.error(error)
        st.stop()

    result, diff_cols = reconcile(expected, processed, fields, tolerance)
    exceptions = result[result["Overall Status"] == "Exception"].copy()

    total = len(result)
    exception_count = len(exceptions)
    ok_count = total - exception_count
    total_abs_difference = result[diff_cols].abs().sum().sum() if diff_cols else 0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Employees checked", total)
    c2.metric("Matched / OK", ok_count)
    c3.metric("Exceptions", exception_count)
    c4.metric("Total absolute variance", f"£{total_abs_difference:,.2f}")

    st.subheader("Reconciliation results")
    st.dataframe(result, use_container_width=True, hide_index=True)

    st.subheader("Exceptions")
    if exceptions.empty:
        st.success("No reconciliation exceptions found.")
    else:
        st.dataframe(exceptions, use_container_width=True, hide_index=True)

    report_bytes = to_excel_bytes(result, exceptions)

    d1, d2 = st.columns(2)
    with d1:
        st.download_button(
            "Download Excel reconciliation report",
            data=report_bytes,
            file_name="payroll_reconciliation_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with d2:
        st.download_button(
            "Download exceptions CSV",
            data=exceptions.to_csv(index=False).encode("utf-8"),
            file_name="payroll_exceptions.csv",
            mime="text/csv",
        )
else:
    st.info("Upload both payroll files to begin.")
