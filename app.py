
import io
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Payroll Reconciliation Checker", page_icon="favicon.svg", layout="wide")

st.markdown("""
<style>
.block-container {
    padding-top: 1.5rem;
    padding-bottom: 3rem;
    max-width: 1200px;
}

.hero-box {
    background: linear-gradient(135deg, #fff7fb 0%, #f6f0ff 45%, #eef4ff 100%);
    border-radius: 24px;
    padding: 40px 36px;
    margin-bottom: 30px;
    box-shadow: 0 10px 30px rgba(0,0,0,0.08);
}

.hero-tag {
    color: #c25b8e;
    font-size: 13px;
    font-weight: 700;
    letter-spacing: 2px;
    text-transform: uppercase;
    margin-bottom: 18px;
}

.hero-title {
    font-size: 56px;
    font-weight: 800;
    line-height: 1.05;
    color: #2c2c2c;
    margin-bottom: 20px;
}

.hero-text {
    font-size: 19px;
    color: #555;
    line-height: 1.7;
    margin-bottom: 24px;
}

.info-card {
    background: white;
    border-radius: 18px;
    padding: 22px;
    box-shadow: 0 8px 24px rgba(0,0,0,0.08);
    border: 1px solid #f0e7f5;
}

.info-card h4 {
    margin-top: 0px;
    color: #333;
    font-size: 22px;
}

.small-box {
    background: #faf7ff;
    border: 1px solid #eee3f8;
    border-radius: 14px;
    padding: 14px;
    margin-bottom: 12px;
    font-size: 15px;
}

.icon-logo {
    width: 90px;
    margin-bottom: 18px;
}

.developed-by {
    font-size: 16px;
    color: #555;
    margin-top: 8px;
    margin-bottom: 20px;
    font-weight: 500;
}

.section-note {
    background: #fff8df;
    border-left: 5px solid #f0c14b;
    padding: 12px 16px;
    border-radius: 10px;
    color: #555;
    margin-top: 15px;
    margin-bottom: 20px;
}
</style>
""", unsafe_allow_html=True)

col1, col2 = st.columns([1.25, 1], gap="large")

with col1:
    st.image("favicon.svg", width=90)

    st.markdown('<div class="hero-tag">Payroll • Reconciliation • Validation</div>', unsafe_allow_html=True)
    st.markdown('<div class="hero-title">Make payroll reconciliation easier to review.</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="hero-text">'
        'Compare expected and processed payroll data, identify differences quickly, '
        'investigate exceptions, and review results in a clear workflow.'
        '</div>',
        unsafe_allow_html=True
    )
    st.markdown(
        '<div class="developed-by"><b>Developed by Sindu</b> | Payroll Reconciliation Checker</div>',
        unsafe_allow_html=True
    )

with col2:
with col2:
    st.markdown("""
    <div class="info-card">
        <h4>Live reconciliation overview</h4>

        <div class="small-box">
            <b>Expected Payroll</b><br>
            Approved payroll values ready for comparison
        </div>

        <div class="small-box">
            <b>Processed Payroll</b><br>
            Payroll output to validate against expected results
        </div>

        <div class="small-box">
            <b>Review Output</b><br>
            Compare values side by side and highlight exceptions
        </div>
    </div>
    """, unsafe_allow_html=True)

st.markdown(
    '<div class="section-note">⚠️ Demo use only: Please use fictional or sample payroll data only. '
    'Do not upload real employee, banking, tax, confidential, or company payroll information.</div>',
    unsafe_allow_html=True
)

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
