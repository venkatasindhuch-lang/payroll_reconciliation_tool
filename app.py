import io
from datetime import date

import pandas as pd
import streamlit as st

st.set_page_config(page_title="CARRR Payroll", page_icon="✅", layout="wide")

# -----------------------------
# CARRR Payroll configuration
# -----------------------------
RECON_FIELDS = [
    "Gross Pay", "Salary", "PAYE", "Employee NI", "Pension",
    "Other Deductions", "Net Pay", "Overtime", "Overtime Pay", "Bonus"
]
REQUIRED_KEYS = ["Employee ID", "Employee Name"]

# CARRR Control Centre styling
st.markdown(
    """
    <style>
    * {box-sizing:border-box}
    .stApp{background:linear-gradient(180deg,#f8fbff 0%,#ffffff 38%,#fbf9ff 100%);color:#17213d}
    .block-container{padding-top:2.1rem;padding-bottom:3rem;max-width:1180px}
    h1,h2,h3{letter-spacing:-.025em}
    .carrr-head{display:flex;justify-content:space-between;gap:1rem;align-items:flex-start;padding:1.15rem 1.3rem;border:1px solid #e9e8f4;border-radius:22px;background:rgba(255,255,255,.92);box-shadow:0 12px 32px rgba(41,52,94,.06);margin-bottom:1rem}
    .carrr-brand{font-size:1.9rem;font-weight:850;color:#141d4c}.carrr-tag{font-size:.95rem;font-weight:750;color:#5a63a8;margin-top:.15rem}.carrr-sub{color:#71758b;font-size:.88rem;margin-top:.25rem}
    .phasebar{display:grid;grid-template-columns:1fr .7fr 1fr .8fr;gap:.5rem;margin:.8rem 0 1.2rem}
    .phase{padding:.72rem .8rem;border-radius:15px;border:1px solid #e5e6f2;background:#fff;text-align:center;font-size:.78rem;font-weight:800;color:#3d4666}
    .phase.before{background:#edf5ff}.phase.run{background:#f6f5f8}.phase.after{background:#f2efff}.phase.gate{background:#effaf4}
    .summary-card{border:1px solid #e8e9f3;background:#fff;border-radius:18px;padding:1rem 1.05rem;min-height:112px}.summary-label{font-size:.75rem;font-weight:800;color:#74788f;text-transform:uppercase;letter-spacing:.06em}.summary-value{font-size:1.75rem;font-weight:850;margin-top:.24rem;color:#17213d}.summary-note{font-size:.8rem;color:#7a7e92;margin-top:.2rem}
    .before-box,.after-box,.gate-box{border-radius:22px;padding:1.2rem 1.25rem;border:1px solid #e7e8f3;background:#fff;margin:.7rem 0 1rem}.before-box{border-left:5px solid #3b82f6}.after-box{border-left:5px solid #7958e8}.gate-box{border-left:5px solid #10a15b}
    .kicker{font-size:.74rem;letter-spacing:.14em;font-weight:850;color:#6e75a6;text-transform:uppercase}.bigline{font-size:1.45rem;font-weight:850;color:#182044;margin:.15rem 0}.muted{color:#73788d;font-size:.9rem}
    .status-pass{color:#087f4a;font-weight:800}.status-warn{color:#a56600;font-weight:800}.status-fail{color:#c72f3c;font-weight:800}
    .exception-head{padding:.9rem 1rem;border-radius:16px;background:#fff7ef;border:1px solid #ffe2bd;margin:.7rem 0}.release-ready{padding:1.1rem 1.2rem;border-radius:18px;background:#edfbf3;border:1px solid #bfead0;color:#086c40;font-size:1.4rem;font-weight:900;text-align:center}.release-hold{padding:1.1rem 1.2rem;border-radius:18px;background:#fff1f2;border:1px solid #ffcdd3;color:#a51f31;font-size:1.4rem;font-weight:900;text-align:center}
    .privacy{font-size:.78rem;color:#777b8f;padding:.55rem .7rem;border-radius:12px;background:#fffdf4;border:1px solid #f1ead0;margin:.5rem 0 1rem}
    div[data-testid="stMetric"]{background:white;border:1px solid #ececf4;padding:.75rem .85rem;border-radius:16px}
    @media(max-width:800px){.phasebar{grid-template-columns:1fr 1fr}.carrr-head{display:block}.block-container{padding-left:1rem;padding-right:1rem}.carrr-brand{font-size:1.6rem}}
    </style>
    """, unsafe_allow_html=True
)

st.markdown("""
<div class="carrr-head">
 <div>
  <div class="carrr-brand">CARRR Payroll</div>
  <div class="carrr-tag">Before payroll. After payroll. Before release.</div>
  <div class="carrr-sub">Check · Analyse · Reconcile · Review · Release</div>
 </div>
 <div class="muted">Payroll control centre</div>
</div>
<div class="phasebar">
 <div class="phase before">BEFORE PAYROLL<br>Check + Analyse</div>
 <div class="phase run">PAYROLL RUN<br>Your payroll system</div>
 <div class="phase after">AFTER PAYROLL<br>Reconcile + Review</div>
 <div class="phase gate">RELEASE GATE<br>Ready / Hold</div>
</div>
<div class="privacy"><b>Public trial:</b> use fictional or anonymised payroll data. Real payroll data should only be handled in an appropriately secured private deployment.</div>
""", unsafe_allow_html=True)

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


def numeric(series):
    return pd.to_numeric(series, errors="coerce").fillna(0.0)


def safe_text(series):
    return series.fillna("").astype(str).str.strip()


def employee_name(df):
    if "Employee Name" in df.columns:
        return safe_text(df["Employee Name"])
    return pd.Series([""] * len(df), index=df.index)


def add_issue(rows, severity, check_name, detail, emp_id="", emp_name="", source="CHECK"):
    rows.append({
        "Source": source,
        "Status": severity,
        "Employee ID": str(emp_id) if emp_id is not None else "",
        "Employee Name": str(emp_name) if emp_name is not None else "",
        "Check / Issue": check_name,
        "Detail": detail,
    })


def check_payroll(df, label):
    rows = []
    names = employee_name(df)

    # Required columns
    for col in REQUIRED_KEYS:
        if col not in df.columns:
            add_issue(rows, "Fail", f"{label}: required column", f"Missing required column: {col}")

    if "Employee ID" not in df.columns:
        return pd.DataFrame(rows)

    ids = safe_text(df["Employee ID"])

    # Blank key data
    for idx in df.index[ids.eq("")]:
        add_issue(rows, "Fail", "Missing Employee ID", "Employee ID is blank.", "", names.loc[idx], "CHECK")

    if "Employee Name" in df.columns:
        blank_names = safe_text(df["Employee Name"]).eq("")
        for idx in df.index[blank_names]:
            add_issue(rows, "Fail", "Missing Employee Name", "Employee Name is blank.", ids.loc[idx], "", "CHECK")

    # Duplicate IDs
    dup = ids.ne("") & ids.duplicated(keep=False)
    for idx in df.index[dup]:
        add_issue(rows, "Fail", "Duplicate employee", "Duplicate Employee ID detected.", ids.loc[idx], names.loc[idx], "CHECK")

    # Optional payroll master-data checks: only enforce when the column exists.
    optional_fields = {
        "Tax Code": "Tax code is blank.",
        "NI Number": "NI number is blank.",
        "Pension": "Pension value/status is blank.",
        "Bank Details": "Bank details are blank.",
        "Sort Code": "Sort code is blank.",
        "Account Number": "Account number is blank.",
    }
    for col, detail in optional_fields.items():
        if col in df.columns:
            blank = safe_text(df[col]).eq("")
            for idx in df.index[blank]:
                add_issue(rows, "Warning", f"Missing {col}", detail, ids.loc[idx], names.loc[idx], "CHECK")

    # Light NI format check for demo validation (not a substitute for HMRC validation).
    if "NI Number" in df.columns:
        ni = safe_text(df["NI Number"]).str.upper().str.replace(" ", "", regex=False)
        valid = ni.eq("") | ni.str.match(r"^[A-CEGHJ-PR-TW-Z]{2}\d{6}[A-D]$")
        for idx in df.index[~valid]:
            add_issue(rows, "Warning", "NI number format", "NI number format looks unusual; verify against source data.", ids.loc[idx], names.loc[idx], "CHECK")

    # Numeric sanity checks for available pay fields.
    for col in ["Gross Pay", "Salary", "PAYE", "Employee NI", "Pension", "Other Deductions", "Net Pay", "Overtime", "Overtime Pay", "Overtime Hours", "Bonus"]:
        if col in df.columns:
            raw = pd.to_numeric(df[col], errors="coerce")
            bad = raw.isna() & safe_text(df[col]).ne("")
            for idx in df.index[bad]:
                add_issue(rows, "Fail", f"Invalid {col}", f"{col} is not numeric.", ids.loc[idx], names.loc[idx], "CHECK")

    if not rows:
        add_issue(rows, "Pass", f"{label}: validation", "No validation issues detected in the supplied columns.", source="CHECK")

    return pd.DataFrame(rows)


def analyse_payroll(expected, processed, overtime_limit, pay_change_pct, overtime_hours_limit):
    rows = []
    if "Employee ID" not in processed.columns:
        add_issue(rows, "Fail", "Analysis unavailable", "Processed payroll has no Employee ID column.", source="ANALYSE")
        return pd.DataFrame(rows)

    proc = processed.copy()
    proc["Employee ID"] = safe_text(proc["Employee ID"])
    proc_names = employee_name(proc)

    # Duplicate employee IDs
    dup = proc["Employee ID"].ne("") & proc["Employee ID"].duplicated(keep=False)
    for idx in proc.index[dup]:
        add_issue(rows, "Fail", "Duplicate employee", "Duplicate employee appears in processed payroll.", proc.loc[idx, "Employee ID"], proc_names.loc[idx], "ANALYSE")

    # Negative net pay
    if "Net Pay" in proc.columns:
        net = numeric(proc["Net Pay"])
        for idx in proc.index[net.lt(0)]:
            add_issue(rows, "Fail", "Negative net pay", f"Net pay is £{net.loc[idx]:,.2f}.", proc.loc[idx, "Employee ID"], proc_names.loc[idx], "ANALYSE")

    # High overtime amount.
    overtime_col = None
    if "Overtime" in proc.columns:
        overtime_col = "Overtime"
    elif "Overtime Pay" in proc.columns:
        overtime_col = "Overtime Pay"
    if overtime_col:
        ot = numeric(proc[overtime_col])
        for idx in proc.index[ot.gt(overtime_limit)]:
            add_issue(rows, "Warning", "Unusually high overtime", f"{overtime_col} is {ot.loc[idx]:,.2f}, above the review threshold of {overtime_limit:,.2f}.", proc.loc[idx, "Employee ID"], proc_names.loc[idx], "ANALYSE")

    # High overtime hours when hours are supplied separately.
    if "Overtime Hours" in proc.columns:
        ot_hours = numeric(proc["Overtime Hours"])
        for idx in proc.index[ot_hours.gt(overtime_hours_limit)]:
            add_issue(rows, "Warning", "Unusually high overtime hours", f"Overtime Hours are {ot_hours.loc[idx]:,.2f}, above the review threshold of {overtime_hours_limit:,.2f} hours.", proc.loc[idx, "Employee ID"], proc_names.loc[idx], "ANALYSE")

    # Large pay movements between expected and processed.
    pay_col = "Gross Pay" if "Gross Pay" in proc.columns and "Gross Pay" in expected.columns else ("Salary" if "Salary" in proc.columns and "Salary" in expected.columns else None)
    if pay_col and "Employee ID" in expected.columns:
        exp = expected[["Employee ID", pay_col]].copy()
        exp["Employee ID"] = safe_text(exp["Employee ID"])
        exp[f"{pay_col} Expected"] = numeric(exp[pay_col])
        merged = proc[["Employee ID", pay_col]].copy().merge(
            exp[["Employee ID", f"{pay_col} Expected"]], on="Employee ID", how="left"
        )
        merged[f"{pay_col} Processed"] = numeric(merged[pay_col])
        base = merged[f"{pay_col} Expected"].abs()
        merged["Change %"] = ((merged[f"{pay_col} Processed"] - merged[f"{pay_col} Expected"]).abs() / base.replace(0, pd.NA) * 100).fillna(0)
        big = merged["Change %"].gt(pay_change_pct)
        name_lookup = proc.set_index("Employee ID").get("Employee Name", pd.Series(dtype=object)) if "Employee Name" in proc.columns else pd.Series(dtype=object)
        for idx in merged.index[big]:
            emp_id = merged.loc[idx, "Employee ID"]
            emp_name = ""
            if not name_lookup.empty and emp_id in name_lookup.index:
                value = name_lookup.loc[emp_id]
                emp_name = value.iloc[0] if isinstance(value, pd.Series) else value
            add_issue(rows, "Warning", "Large pay change", f"{pay_col} changed by {merged.loc[idx, 'Change %']:.1f}% against expected payroll.", emp_id, emp_name, "ANALYSE")

    # Unexpected processed employees
    if "Employee ID" in expected.columns:
        expected_ids = set(safe_text(expected["Employee ID"]))
        unexpected = ~proc["Employee ID"].isin(expected_ids)
        for idx in proc.index[unexpected]:
            add_issue(rows, "Fail", "Unexpected payment / employee", "Employee appears in processed payroll but not in expected payroll.", proc.loc[idx, "Employee ID"], proc_names.loc[idx], "ANALYSE")

    # Leaver payment check when a status field exists.
    status_col = next((c for c in ["Employment Status", "Employee Status", "Starter/Leaver", "Status"] if c in proc.columns), None)
    if status_col and "Net Pay" in proc.columns:
        status = safe_text(proc[status_col]).str.lower()
        net = numeric(proc["Net Pay"])
        leaver_paid = status.str.contains("leaver|left|terminated|inactive", regex=True) & net.gt(0)
        for idx in proc.index[leaver_paid]:
            add_issue(rows, "Warning", "Leaver payment", f"{status_col} indicates a leaver/inactive employee but net pay is £{net.loc[idx]:,.2f}.", proc.loc[idx, "Employee ID"], proc_names.loc[idx], "ANALYSE")

    if not rows:
        add_issue(rows, "Pass", "Automated analysis", "No analytical exceptions detected using the selected thresholds.", source="ANALYSE")

    return pd.DataFrame(rows)


def validate_reconciliation_columns(df, label, fields):
    errors = []
    required = REQUIRED_KEYS + fields
    missing = [c for c in required if c not in df.columns]
    if missing:
        errors.append(f"{label}: missing columns: {', '.join(missing)}")
    if "Employee ID" in df.columns:
        dup = safe_text(df["Employee ID"]).duplicated(keep=False)
        if dup.any():
            errors.append(f"{label}: {int(dup.sum())} rows contain duplicate Employee IDs.")
    return errors


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
    expected["Employee ID"] = safe_text(expected["Employee ID"])
    processed["Employee ID"] = safe_text(processed["Employee ID"])

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

    ordered = [
        "Employee ID", "Employee Name", "Record Status", "Overall Status", "Exception Reason"
    ] + [
        col for field in fields for col in (
            f"{field} Expected", f"{field} Processed", f"{field} Difference", f"{field} Status"
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


def reconciliation_review_items(result):
    rows = []
    exceptions = result[result["Overall Status"].eq("Exception")]
    for _, r in exceptions.iterrows():
        rows.append({
            "Source": "RECONCILE",
            "Status": "Fail",
            "Employee ID": r.get("Employee ID", ""),
            "Employee Name": r.get("Employee Name", ""),
            "Check / Issue": "Reconciliation exception",
            "Detail": r.get("Exception Reason", ""),
        })
    if not rows:
        add_issue(rows, "Pass", "Reconciliation", "Expected and processed payroll match within tolerance.", source="RECONCILE")
    return pd.DataFrame(rows)



def control_coverage(expected, processed):
    controls = [
        ("Employee ID", ["Employee ID"]),
        ("Employee Name", ["Employee Name"]),
        ("Tax Code", ["Tax Code"]),
        ("NI Number", ["NI Number"]),
        ("Bank Details", ["Bank Details", "Sort Code", "Account Number"]),
        ("Pension", ["Pension"]),
        ("Starter / Leaver Status", ["Employment Status", "Employee Status", "Starter/Leaver", "Status"]),
        ("Salary / Gross Pay", ["Salary", "Gross Pay"]),
        ("Overtime", ["Overtime", "Overtime Pay", "Overtime Hours"]),
        ("Bonus", ["Bonus"]),
        ("Deductions", ["PAYE", "Employee NI", "Other Deductions"]),
        ("Net Pay", ["Net Pay"]),
    ]
    rows = []
    for label, candidates in controls:
        exp_cols = [c for c in candidates if c in expected.columns]
        proc_cols = [c for c in candidates if c in processed.columns]
        rows.append({
            "Control": label,
            "Expected File": ", ".join(exp_cols) if exp_cols else "Not supplied",
            "Processed File": ", ".join(proc_cols) if proc_cols else "Not supplied",
            "Coverage": "Available" if exp_cols and proc_cols else "Partial / Missing",
        })
    return pd.DataFrame(rows)

def build_review_queue(check_df, analysis_df, recon_df):
    queue = pd.concat([check_df, analysis_df, recon_df], ignore_index=True)
    queue.insert(0, "Exception ID", [f"CARRR-{i+1:04d}" for i in range(len(queue))])
    queue["Reviewed"] = queue["Status"].eq("Pass")
    queue["Review Status"] = queue["Status"].map({"Pass": "Not required", "Warning": "Open", "Fail": "Open"})
    queue["Reviewer Note"] = ""
    return queue


def to_excel_bytes(summary_df, check_df, analysis_df, controls_df, recon_df, review_df, release_df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="Summary", index=False)
        check_df.to_excel(writer, sheet_name="Check", index=False)
        analysis_df.to_excel(writer, sheet_name="Analyse", index=False)
        controls_df.to_excel(writer, sheet_name="Control Totals", index=False)
        recon_df.to_excel(writer, sheet_name="Reconciliation", index=False)
        review_df.to_excel(writer, sheet_name="Review", index=False)
        release_df.to_excel(writer, sheet_name="Release", index=False)
    return output.getvalue()




def analyse_input_payroll(df, overtime_limit=500.0, overtime_hours_limit=20.0, bonus_limit=1000.0):
    rows=[]
    if "Employee ID" not in df.columns:
        add_issue(rows,"Fail","Analysis unavailable","Payroll input has no Employee ID column.",source="ANALYSE")
        return pd.DataFrame(rows)
    ids=safe_text(df["Employee ID"]); names=employee_name(df)
    dup=ids.ne("") & ids.duplicated(keep=False)
    for idx in df.index[dup]:
        add_issue(rows,"Fail","Duplicate employee","Employee ID appears more than once in payroll input.",ids.loc[idx],names.loc[idx],"ANALYSE")
    for col,label,limit in [("Overtime Pay","High overtime pay",overtime_limit),("Overtime","High overtime",overtime_limit),("Overtime Hours","High overtime hours",overtime_hours_limit),("Bonus","High bonus",bonus_limit)]:
        if col in df.columns:
            vals=numeric(df[col])
            for idx in df.index[vals.gt(limit)]:
                add_issue(rows,"Warning",label,f"{col} is {vals.loc[idx]:,.2f}, above the review threshold of {limit:,.2f}.",ids.loc[idx],names.loc[idx],"ANALYSE")
    if not rows:
        add_issue(rows,"Pass","Before-payroll analysis","No input exceptions detected using the selected thresholds.",source="ANALYSE")
    return pd.DataFrame(rows)


def attention_only(df):
    if df is None or df.empty or "Status" not in df.columns:
        return pd.DataFrame(columns=["Status","Employee ID","Employee Name","Check / Issue","Detail"])
    return df[df["Status"].isin(["Warning","Fail"])].copy()


# -----------------------------
# Simple CARRR control-centre workflow
# -----------------------------
with st.expander("What files does CARRR need?", expanded=False):
    st.markdown("""
**Before payroll — Payroll Input / Expected Data**  
Export or prepare the approved payroll inputs/expected values (Excel or CSV).

**After payroll — Pre-Release Payroll Output**  
After your payroll system calculates payroll, export the pre-release payroll report (Excel or CSV) and upload it here.

CARRR does not send payments. It checks and controls payroll before the final release in your payroll system.
""")

payroll_period=st.text_input("Payroll period",placeholder="Example: October 2026")

st.markdown('<div class="before-box"><div class="kicker">Before payroll</div><div class="bigline">1. Upload the payroll input</div><div class="muted">CARRR checks the input population and flags only the records that need attention.</div></div>',unsafe_allow_html=True)
input_file=st.file_uploader("Payroll Input / Expected Data",type=["csv","xlsx","xls"],key="carrr_input")

with st.expander("Before-payroll thresholds", expanded=False):
    t1,t2,t3=st.columns(3)
    overtime_limit=t1.number_input("High overtime pay (£)",min_value=0.0,value=500.0,step=50.0)
    overtime_hours_limit=t2.number_input("High overtime hours",min_value=0.0,value=20.0,step=1.0)
    bonus_limit=t3.number_input("High bonus (£)",min_value=0.0,value=1000.0,step=100.0)

if not input_file:
    st.info("Start here: upload the Payroll Input / Expected Data. CARRR will run the before-payroll checks first.")
    st.stop()

try:
    expected=clean_columns(read_upload(input_file))
except Exception as exc:
    st.error(f"Could not read the payroll input: {exc}")
    st.stop()

check_input=check_payroll(expected,"Payroll input")
analysis_input=analyse_input_payroll(expected,overtime_limit,overtime_hours_limit,bonus_limit)
before_items=pd.concat([attention_only(check_input),attention_only(analysis_input)],ignore_index=True)

employee_count=len(expected)
before_fails=int(before_items["Status"].eq("Fail").sum()) if not before_items.empty else 0
before_warn=int(before_items["Status"].eq("Warning").sum()) if not before_items.empty else 0
before_attention=len(before_items)

c1,c2,c3,c4=st.columns(4)
c1.metric("Employees",employee_count)
c2.metric("Before-payroll issues",before_attention)
c3.metric("Warnings",before_warn)
c4.metric("Fails",before_fails)

if before_attention==0:
    st.success("Before payroll: no input exceptions detected. You can proceed to your payroll system.")
else:
    st.markdown(f'<div class="exception-head"><b>{before_attention} item(s) need attention before payroll.</b> You do not need to inspect every employee — start with these exceptions.</div>',unsafe_allow_html=True)
    st.dataframe(before_items[[c for c in ["Status","Employee ID","Employee Name","Check / Issue","Detail"] if c in before_items.columns]],use_container_width=True,hide_index=True)

with st.expander("Show full before-payroll validation", expanded=False):
    st.dataframe(check_input,use_container_width=True,hide_index=True)
    st.dataframe(analysis_input,use_container_width=True,hide_index=True)

st.markdown('<div class="after-box"><div class="kicker">After payroll</div><div class="bigline">2. Upload the pre-release output</div><div class="muted">Once your payroll system has calculated payroll, upload its pre-release output. CARRR will reconcile it against the input automatically.</div></div>',unsafe_allow_html=True)
output_file=st.file_uploader("Pre-Release Payroll Output",type=["csv","xlsx","xls"],key="carrr_output")

if not output_file:
    st.info("Before-payroll checks are complete. Run payroll in your payroll system, then return here and upload the Pre-Release Payroll Output.")
    st.stop()

try:
    processed=clean_columns(read_upload(output_file))
except Exception as exc:
    st.error(f"Could not read the pre-release output: {exc}")
    st.stop()

common_fields=[f for f in RECON_FIELDS if f in expected.columns and f in processed.columns]
if not common_fields:
    st.error("The two files do not share a supported payroll value such as Gross Pay, PAYE, Pension or Net Pay.")
    st.stop()

with st.expander("Reconciliation settings",expanded=False):
    s1,s2=st.columns(2)
    fields=s1.multiselect("Fields to compare",common_fields,default=common_fields)
    tolerance=s2.number_input("Allowed difference (£)",min_value=0.0,value=0.01,step=0.01)

if not fields:
    st.warning("Select at least one field to compare.")
    st.stop()

errors=validate_reconciliation_columns(expected,"Payroll input",fields)+validate_reconciliation_columns(processed,"Pre-release output",fields)
if errors:
    st.error("CARRR cannot reconcile until these file-structure issues are fixed:")
    for e in errors: st.write("- "+e)
    st.stop()

check_output=check_payroll(processed,"Pre-release output")
reconciliation=reconcile(expected,processed,fields,tolerance)
controls=control_totals(reconciliation,fields)
recon_items=reconciliation_review_items(reconciliation)
after_items=pd.concat([attention_only(check_output),attention_only(recon_items)],ignore_index=True)
all_items=pd.concat([before_items,after_items],ignore_index=True)

exceptions_count=int(reconciliation["Overall Status"].eq("Exception").sum())
matched_count=len(reconciliation)-exceptions_count
net_variance=float(reconciliation["Net Pay Difference"].sum()) if "Net Pay Difference" in reconciliation.columns else 0.0

st.markdown("### CARRR result")
r1,r2,r3,r4=st.columns(4)
r1.metric("Employees checked",len(reconciliation))
r2.metric("Matched / OK",matched_count)
r3.metric("Employees with exceptions",exceptions_count)
r4.metric("Net pay variance",f"£{net_variance:,.2f}")

# Split true payroll mismatches from threshold/master-data warnings so users can fix the important records first.
mismatch_rows = reconciliation[reconciliation["Overall Status"].eq("Exception")].copy()
if not mismatch_rows.empty:
    def mismatch_type(row):
        if row.get("Record Status") == "Missing from processed payroll":
            return "Missing employee/output"
        if row.get("Record Status") == "Unexpected employee in processed payroll":
            return "Unexpected employee/payment"
        return "Payroll value mismatch"

    mismatch_summary = pd.DataFrame({
        "Employee ID": mismatch_rows["Employee ID"],
        "Employee Name": mismatch_rows["Employee Name"],
        "Mismatch Type": mismatch_rows.apply(mismatch_type, axis=1),
        "What changed": mismatch_rows["Exception Reason"],
        "Net Pay Difference": mismatch_rows["Net Pay Difference"] if "Net Pay Difference" in mismatch_rows.columns else 0.0,
    })
    mismatch_summary["Action"] = mismatch_summary["Mismatch Type"].map({
        "Missing employee/output": "Check payroll output / missing employee",
        "Unexpected employee/payment": "Verify employee and payment before release",
        "Payroll value mismatch": "Check source input and payroll calculation",
    })
else:
    mismatch_summary = pd.DataFrame(columns=["Employee ID","Employee Name","Mismatch Type","What changed","Net Pay Difference","Action"])

other_after_warnings = attention_only(check_output)

st.markdown("### Fix First — Payroll Mismatches")
if mismatch_summary.empty:
    st.success("No payroll value mismatches found.")
else:
    st.markdown(
        f'<div class="exception-head"><b>{len(mismatch_summary)} employee(s) have payroll mismatches.</b> '
        'These are the records to investigate first. They are separated from general warnings so you do not need to search through the full review queue.</div>',
        unsafe_allow_html=True,
    )
    show_mismatch = mismatch_summary.copy()
    if "Net Pay Difference" in show_mismatch.columns:
        show_mismatch["Net Pay Difference"] = show_mismatch["Net Pay Difference"].map(lambda x: f"£{x:,.2f}")
    st.dataframe(show_mismatch, use_container_width=True, hide_index=True)

if all_items.empty:
    st.success("No review exceptions detected.")
else:
    st.markdown(
        f'<div class="exception-head"><b>Review Centre: {len(all_items)} total review item(s).</b> '
        f'Payroll mismatches: {len(mismatch_summary)} employee(s). Before-payroll warnings: {len(before_items)} item(s). '
        'Use the separate sections below instead of searching one long table.</div>',
        unsafe_allow_html=True,
    )

# Build review queue from attention items only; no pass rows cluttering the screen.
review_queue=all_items.copy()
if review_queue.empty:
    review_queue=pd.DataFrame(columns=["Source","Status","Employee ID","Employee Name","Check / Issue","Detail"])
review_queue.insert(0,"Exception ID",[f"CARRR-{i+1:04d}" for i in range(len(review_queue))])
review_queue["Reviewed"]=False
review_queue["Review Status"]="Open"
review_queue["Reviewer Note"]=""

main_tab,mismatch_tab,warning_tab,detail_tab,review_tab,release_tab=st.tabs(["Overview","Fix First: Mismatches","Warnings","Reconciliation Detail","Review & Notes","Release Gate"])

with main_tab:
    st.markdown("#### What needs attention")
    o1,o2,o3=st.columns(3)
    o1.metric("Payroll mismatches", len(mismatch_summary))
    o2.metric("Before-payroll warnings", len(before_items))
    o3.metric("Other output warnings", len(other_after_warnings))
    st.caption("Start with Payroll mismatches. Warnings are secondary checks and can be reviewed afterwards.")
    with st.expander("Payroll control totals",expanded=False):
        shown=controls.copy()
        for col in ["Expected Total","Processed Total","Variance"]:
            if col in shown.columns: shown[col]=shown[col].map(lambda x:f"£{x:,.2f}")
        st.dataframe(shown,use_container_width=True,hide_index=True)

with mismatch_tab:
    st.markdown("#### Payroll mismatches — investigate these first")
    if mismatch_summary.empty:
        st.success("No payroll mismatches found.")
    else:
        mm=mismatch_summary.copy()
        mm["Net Pay Difference"] = mm["Net Pay Difference"].map(lambda x:f"£{x:,.2f}")
        st.dataframe(mm,use_container_width=True,hide_index=True)
        st.info("A mismatch shows that the pre-release payroll output differs from the expected/input data. Check the source input and the payroll-system calculation before release.")

with warning_tab:
    st.markdown("#### Warnings — review after payroll mismatches")
    st.caption("Warnings can include high overtime, missing master data, leaver payments or other threshold checks. They are separated from true reconciliation mismatches.")
    if before_items.empty and other_after_warnings.empty:
        st.success("No warnings found.")
    else:
        if not before_items.empty:
            st.markdown(f"**Before-payroll warnings ({len(before_items)})**")
            st.dataframe(before_items[[c for c in ["Status","Employee ID","Employee Name","Check / Issue","Detail"] if c in before_items.columns]],use_container_width=True,hide_index=True)
        if not other_after_warnings.empty:
            st.markdown(f"**Pre-release output warnings ({len(other_after_warnings)})**")
            st.dataframe(other_after_warnings[[c for c in ["Status","Employee ID","Employee Name","Check / Issue","Detail"] if c in other_after_warnings.columns]],use_container_width=True,hide_index=True)

with detail_tab:
    st.caption("Use this detailed table only when you need to investigate an individual employee or field difference.")
    st.dataframe(reconciliation,use_container_width=True,hide_index=True)

with review_tab:
    st.markdown("#### Record review decisions and notes")
    if review_queue.empty:
        review_editor=review_queue.copy()
        st.success("No exceptions require human review.")
    else:
        review_editor=st.data_editor(
            review_queue,use_container_width=True,hide_index=True,
            disabled=["Exception ID","Source","Status","Employee ID","Employee Name","Check / Issue","Detail"],
            column_config={
                "Reviewed":st.column_config.CheckboxColumn("Reviewed"),
                "Review Status":st.column_config.SelectboxColumn("Outcome",options=["Open","Resolved","Accepted"],required=True),
                "Reviewer Note":st.column_config.TextColumn("Reviewer Note"),
            },key="carrr_v2_review")
        open_items=review_editor[(~review_editor["Reviewed"])|review_editor["Review Status"].eq("Open")]
        st.metric("Total review items still open",len(open_items))

with release_tab:
    st.markdown("#### CARRR Release Gate")
    if review_queue.empty:
        final_review=review_queue.copy(); open_items=review_queue.copy()
    else:
        final_review=review_editor.copy()
        open_items=final_review[(~final_review["Reviewed"])|final_review["Review Status"].eq("Open")]
    p1,p2=st.columns(2)
    prepared_by=p1.text_input("Prepared by")
    reviewed_by=p2.text_input("Reviewed by")
    review_date=st.date_input("Review date",value=date.today())
    approval=st.selectbox("Approval status",["Pending","Approved","Hold"])
    c1=st.checkbox("Payroll inputs checked")
    c2=st.checkbox("Exceptions reviewed")
    c3=st.checkbox("Reconciliation reviewed")
    ready=bool(prepared_by.strip()) and bool(reviewed_by.strip()) and approval=="Approved" and c1 and c2 and c3 and open_items.empty
    if ready:
        st.markdown('<div class="release-ready">✓ READY TO RELEASE</div>',unsafe_allow_html=True)
        decision="READY TO RELEASE"
    else:
        st.markdown('<div class="release-hold">HOLD PAYROLL</div>',unsafe_allow_html=True)
        decision="HOLD PAYROLL"
        reasons=[]
        if not prepared_by.strip() or not reviewed_by.strip(): reasons.append("reviewer names incomplete")
        if approval!="Approved": reasons.append("approval not recorded")
        if not (c1 and c2 and c3): reasons.append("release checklist incomplete")
        if not open_items.empty: reasons.append(f"{len(open_items)} exception(s) still open")
        if reasons: st.caption("Hold reason: "+"; ".join(reasons)+".")

    summary_df=pd.DataFrame([
        {"Metric":"Payroll Period","Value":payroll_period or "Not specified"},
        {"Metric":"Employees Checked","Value":len(reconciliation)},
        {"Metric":"Matched / OK","Value":matched_count},
        {"Metric":"Reconciliation Exceptions","Value":exceptions_count},
        {"Metric":"Open Review Items","Value":len(open_items)},
        {"Metric":"Net Pay Variance","Value":round(net_variance,2)},
        {"Metric":"Final Decision","Value":decision},
    ])
    release_df=pd.DataFrame([{"Payroll Period":payroll_period,"Prepared By":prepared_by,"Reviewed By":reviewed_by,"Review Date":str(review_date),"Approval Status":approval,"Open Review Items":len(open_items),"Decision":decision}])
    report=to_excel_bytes(summary_df,check_input,analysis_input,controls,reconciliation,final_review,release_df)
    fp=(payroll_period.strip().replace(" ","_") if payroll_period.strip() else "payroll")
    st.download_button("Download CARRR audit report",report,file_name=f"CARRR_{fp}_audit_report.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
