
import io
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Payroll Reconciliation Checker", page_icon="favicon.svg", layout="wide")
st.markdown("""
<svg width="100" height="100" viewBox="0 0 64 64"
     xmlns="http://www.w3.org/2000/svg"
     style="display:block; margin-bottom:20px;">

    <defs>
        <linearGradient id="payrollGradient" x1="0" y1="0" x2="1" y2="1">
            <stop offset="0%" stop-color="#35C7FF"/>
            <stop offset="50%" stop-color="#7C4DFF"/>
            <stop offset="100%" stop-color="#FF2E9F"/>
        </linearGradient>
    </defs>

    <rect x="4" y="4" width="56" height="56"
          rx="15" fill="url(#payrollGradient)"/>

    <rect x="14" y="13" width="27" height="34"
          rx="4" fill="none" stroke="white" stroke-width="4"/>

    <path d="M20 21h15 M20 29h4 M31 29h4 M20 37h4 M31 37h4"
          stroke="white" stroke-width="3" stroke-linecap="round"/>

    <ellipse cx="46" cy="40" rx="8" ry="4"
             fill="none" stroke="white" stroke-width="3"/>

    <path d="M38 40v8c0 2.2 3.6 4 8 4s8-1.8 8-4v-8"
          fill="none" stroke="white" stroke-width="3"/>

</svg>
""", unsafe_allow_html=True)

st.markdown("""
<style>

.block-container {
    padding-top: 2rem;
    padding-bottom: 4rem;
    max-width: 1200px;
}

h1 {
    color: #0A214A;
    font-weight: 800;
}

h2, h3 {
    color: #173A63;
    font-weight: 700;
}

div[data-testid="stAlert"] {
    border-radius: 14px;
}

div[data-testid="stFileUploader"] {
    background: #F8FAFC;
    border-radius: 14px;
    padding: 10px;
}

.stButton > button,
.stDownloadButton > button {
    border-radius: 10px;
    font-weight: 600;
    min-height: 44px;
}

div[data-testid="stDataFrame"] {
    border: 1px solid #E4E7EC;
    border-radius: 12px;
    overflow: hidden;
}

div[data-testid="stExpander"] {
    border-radius: 12px;
    border: 1px solid #E4E7EC;
}

</style>
""", unsafe_allow_html=True)

DEFAULT_FIELDS = ["Gross Pay", "PAYE", "Employee NI", "Pension", "Other Deductions", "Net Pay"]
REQUIRED_KEYS = ["Employee ID", "Employee Name"]

def clean_columns(df):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df

def read_upload(uploaded_file):
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(uploaded_file)
    if name.endswith((".xlsx", ".xls")):
        return pd.read_excel(uploaded_file)
    raise ValueError("Please upload a CSV or Excel file.")

def validate(df, label, fields):
    errors = []
    required = REQUIRED_KEYS + fields
    missing = [c for c in required if c not in df.columns]
    if missing:
        errors.append(f"{label}: missing columns: {', '.join(missing)}")
    if "Employee ID" in df.columns:
        dup = df["Employee ID"].astype(str).duplicated(keep=False)
        if dup.any():
            errors.append(f"{label}: {int(dup.sum())} rows contain duplicate Employee IDs.")
    return errors

def numeric(series):
    return pd.to_numeric(series, errors="coerce").fillna(0.0)

def build_exception_reason(row, fields):
    reasons = []
    if row["Record Status"] == "Missing from processed payroll":
        reasons.append("Missing from processed payroll")
    elif row["Record Status"] == "Unexpected employee in processed payroll":
        reasons.append("Unexpected employee in processed payroll")
    for field in fields:
        if row.get(f"{field} Status") == "Mismatch":
            diff = row.get(f"{field} Difference", 0)
            reasons.append(f"{field} mismatch ({diff:+,.2f})")
    return "; ".join(reasons) if reasons else "No exception"

def reconcile(expected, processed, fields, tolerance):
    expected = expected.copy()
    processed = processed.copy()
    expected["Employee ID"] = expected["Employee ID"].astype(str).str.strip()
    processed["Employee ID"] = processed["Employee ID"].astype(str).str.strip()

    expected = expected[["Employee ID", "Employee Name"] + fields]
    processed = processed[["Employee ID", "Employee Name"] + fields]

    expected = expected.rename(columns={c: f"{c} Expected" for c in expected.columns if c != "Employee ID"})
    processed = processed.rename(columns={c: f"{c} Processed" for c in processed.columns if c != "Employee ID"})

    merged = expected.merge(processed, on="Employee ID", how="outer", indicator=True)
    merged["Employee Name"] = merged["Employee Name Expected"].fillna(merged["Employee Name Processed"])

    status_cols = []
    for field in fields:
        exp_col = f"{field} Expected"
        proc_col = f"{field} Processed"
        diff_col = f"{field} Difference"
        status_col = f"{field} Status"
        merged[exp_col] = numeric(merged[exp_col])
        merged[proc_col] = numeric(merged[proc_col])
        merged[diff_col] = (merged[proc_col] - merged[exp_col]).round(2)
        merged[status_col] = merged[diff_col].abs().le(tolerance).map({True: "OK", False: "Mismatch"})
        status_cols.append(status_col)

    merged["Record Status"] = merged["_merge"].map({
        "left_only": "Missing from processed payroll",
        "right_only": "Unexpected employee in processed payroll",
        "both": "Present in both",
    })
    merged["Overall Status"] = "OK"
    merged.loc[merged["_merge"] != "both", "Overall Status"] = "Exception"
    merged.loc[merged[status_cols].eq("Mismatch").any(axis=1), "Overall Status"] = "Exception"
    merged["Exception Reason"] = merged.apply(lambda r: build_exception_reason(r, fields), axis=1)
    merged["Investigation Status"] = merged["Overall Status"].map({"OK": "Not required", "Exception": "Open"})
    merged["Comments"] = ""
    merged["Action Required"] = ""

    ordered = [
        "Employee ID", "Employee Name", "Record Status", "Overall Status",
        "Exception Reason", "Investigation Status", "Comments", "Action Required"
    ] + [
        col for field in fields for col in (
            f"{field} Expected", f"{field} Processed",
            f"{field} Difference", f"{field} Status"
        )
    ]
    return merged[ordered]

def control_totals(result, fields):
    rows = []
    for field in fields:
        exp = result[f"{field} Expected"].sum()
        proc = result[f"{field} Processed"].sum()
        rows.append({
            "Payroll Control": field,
            "Expected Total": round(exp, 2),
            "Processed Total": round(proc, 2),
            "Variance": round(proc - exp, 2),
        })
    return pd.DataFrame(rows)

def to_excel_bytes(full_df, exceptions_df, controls_df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        controls_df.to_excel(writer, sheet_name="Control Totals", index=False)
        full_df.to_excel(writer, sheet_name="Reconciliation", index=False)
        exceptions_df.to_excel(writer, sheet_name="Exceptions", index=False)
    return output.getvalue()

st.title("Payroll Reconciliation Checker")

st.markdown("**Developed by Sindu | Payroll Reconciliation Checker**")

st.markdown("""
### Smarter payroll comparison and exception review

The **Payroll Reconciliation Checker** compares an **Expected Payroll** file
with a **Processed Payroll** file and presents the values side by side.

It helps identify differences across key payroll fields such as:

- Gross Pay
- PAYE
- Employee NI
- Pension
- Other Deductions
- Net Pay

The tool is designed to make payroll reconciliation easier to review,
highlight exceptions quickly, and support investigation before payroll
results are finalised.
""")

st.info(
    "💡 How it works: Download the demo files → Upload Expected Payroll and "
    "Processed Payroll → Review the comparison → Investigate exceptions."
)

st.warning(
    "⚠️ Demo use only: Please use fictional or sample payroll data. "
    "Do not upload real employee, banking, tax, confidential, or company payroll information."
)

with st.expander("Required file format"):
    st.write("Both files should contain Employee ID, Employee Name, and the payroll fields you want to compare.")

payroll_period = st.text_input("Payroll period", placeholder="Example: September 2026")

st.subheader("🧪 Try the Demo")

st.write(
    "Download the two fictional sample payroll files below, "
    "then upload them into the Expected and Processed payroll sections."
)

demo1, demo2 = st.columns(2)

with demo1:
    with open("expected_payroll_demo.xlsx", "rb") as file:
        st.download_button(
            "📥 Download Expected Demo",
            file,
            file_name="expected_payroll_demo.xlsx"
        )

with demo2:
    with open("processed_payroll_demo.xlsx", "rb") as file:
        st.download_button(
            "📥 Download Processed Demo",
            file,
            file_name="processed_payroll_demo.xlsx"
        )

st.info(
    "Demo data is fictional. Do not upload real employee or payroll information."
)
with st.expander("📘 How to use this app"):
    st.write("""
    1. Download the Expected Payroll Demo file.
    2. Download the Processed Payroll Demo file.
    3. Upload both files below.
    4. Choose the allowed difference.
    5. Run the reconciliation.
    6. Review matched records and payroll exceptions.
    7. Download the reconciliation results.
    """)
    
c1, c2 = st.columns(2)
with c1:
    expected_file = st.file_uploader("1. Expected payroll file", type=["csv", "xlsx", "xls"], key="expected")
with c2:
    processed_file = st.file_uploader("2. Processed payroll file", type=["csv", "xlsx", "xls"], key="processed")

tolerance = st.number_input("Allowed difference (£)", min_value=0.0, value=0.01, step=0.01)

if expected_file and processed_file:
    try:
        expected = clean_columns(read_upload(expected_file))
        processed = clean_columns(read_upload(processed_file))
    except Exception as exc:
        st.error(f"Could not read the files: {exc}")
        st.stop()

    common_fields = [f for f in DEFAULT_FIELDS if f in expected.columns and f in processed.columns]
    fields = st.multiselect("Payroll fields to reconcile", common_fields, default=common_fields)

    if not fields:
        st.warning("Select at least one payroll field.")
        st.stop()

    errors = validate(expected, "Expected payroll", fields) + validate(processed, "Processed payroll", fields)
    if errors:
        for error in errors:
            st.error(error)
        st.stop()

    result = reconcile(expected, processed, fields, tolerance)
    controls = control_totals(result, fields)

    total = len(result)
    exceptions_count = int((result["Overall Status"] == "Exception").sum())
    ok_count = total - exceptions_count
    exception_pct = (exceptions_count / total * 100) if total else 0
    gross_variance = result["Gross Pay Difference"].sum() if "Gross Pay Difference" in result.columns else 0
    net_variance = result["Net Pay Difference"].sum() if "Net Pay Difference" in result.columns else 0

    if payroll_period:
        st.subheader(f"Payroll Period: {payroll_period}")

    m1, m2, m3, m4, m5, m6 = st.columns(6)
    m1.metric("Employees checked", total)
    m2.metric("Matched / OK", ok_count)
    m3.metric("Exceptions", exceptions_count)
    m4.metric("Exception %", f"{exception_pct:.1f}%")
    m5.metric("Gross variance", f"£{gross_variance:,.2f}")
    m6.metric("Net pay variance", f"£{net_variance:,.2f}")

    st.subheader("Payroll control totals")
    shown = controls.copy()
    for col in ["Expected Total", "Processed Total", "Variance"]:
        shown[col] = shown[col].map(lambda x: f"£{x:,.2f}")
    st.dataframe(shown, use_container_width=True, hide_index=True)

    st.subheader("Reconciliation results")
    st.caption("Expected and processed values remain side by side.")
    st.dataframe(result, use_container_width=True, hide_index=True)
  
    st.subheader("Exceptions and investigation")
    exceptions = result[result["Overall Status"] == "Exception"].copy()

    if exceptions.empty:
        st.success("No reconciliation exceptions found.")
        edited_exceptions = exceptions
    else:
        edited_exceptions = st.data_editor(
            exceptions,
            use_container_width=True,
            hide_index=True,
            disabled=[c for c in exceptions.columns if c not in ["Investigation Status", "Comments", "Action Required"]],
            column_config={
                "Investigation Status": st.column_config.SelectboxColumn(
                    "Investigation Status",
                    options=["Open", "Investigating", "Resolved"],
                    required=True,
                ),
                "Comments": st.column_config.TextColumn("Comments"),
                "Action Required": st.column_config.TextColumn("Action Required"),
            },
            key="exceptions_editor",
        )

    final_result = result.copy()
    if not edited_exceptions.empty:
        editable_cols = ["Investigation Status", "Comments", "Action Required"]
        final_result = final_result.set_index("Employee ID")
        edited = edited_exceptions.set_index("Employee ID")
        for col in editable_cols:
            final_result.loc[edited.index, col] = edited[col]
        final_result = final_result.reset_index()

    final_exceptions = final_result[final_result["Overall Status"] == "Exception"].copy()
    report_bytes = to_excel_bytes(final_result, final_exceptions, controls)

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
            data=final_exceptions.to_csv(index=False).encode("utf-8"),
            file_name="payroll_exceptions.csv",
            mime="text/csv",
        )
else:
    st.info("Upload both payroll files to begin.")
