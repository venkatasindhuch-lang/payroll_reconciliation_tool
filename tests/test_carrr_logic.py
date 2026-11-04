"""Automated checks for CARRR's reconciliation, footing, de-duplication and bulk-review logic.

The first group runs the real 500-employee sample files (tests/data) and compares CARRR's answer with an
independent plain-pandas calculation, so a future change cannot quietly stop catching a difference.
The rest are small synthetic cases for each rule.
"""
import pandas as pd
import pytest

from conftest import DATA

NUMERIC = ["Salary", "Overtime Hours", "Overtime Pay", "Bonus", "Gross Pay", "PAYE", "Employee NI", "Pension",
           "Other Deductions", "Net Pay"]

# ------------------------------------------------------------------ 500-employee sample files


def test_population_differences_are_found(app, expected500, output500):
    sv = app.source_validation_summary(expected500, output500, "output_500.xlsx")
    assert sv["Missing ID List"] == ["E0499"]
    assert sv["Unexpected ID List"] == ["E9999"]
    assert sv["Population Aligned"] is False


def test_exact_exception_set(recon500):
    rec, _, _ = recon500
    got = set(rec.loc[rec["Overall Status"] == "Exception", "Employee ID"])
    value_changes = {"E0033", "E0052", "E0128", "E0145", "E0176", "E0215", "E0278", "E0301", "E0367", "E0419",
                     "E0444", "E0455"}
    assert got == value_changes | {"E0231", "E0349"} | {"E0499", "E9999"}
    assert len(rec) == 501 and (rec["Overall Status"] == "OK").sum() == 485


def test_every_independent_difference_is_flagged(app, recon500):
    """Ground truth from the raw spreadsheets with plain pandas; CARRR must flag exactly these, no more, no less."""
    rec, fields, _ = recon500
    e = pd.read_excel(DATA / "expected_500.xlsx").set_index("Employee ID")
    o = pd.read_excel(DATA / "output_500.xlsx").set_index("Employee ID")
    common = e.index.intersection(o.index)
    truth = set()
    for col in NUMERIC:
        diff = (o.loc[common, col].fillna(0) - e.loc[common, col].fillna(0)).round(2)
        truth |= {(i, col) for i in diff.index[diff.abs() > 0.01]}
    both = rec[rec["Record Status"] == "Present in both"]
    flagged = {(r["Employee ID"], f) for _, r in both.iterrows() for f in fields if r[f"{f} Status"] == "Mismatch"}
    assert flagged == truth
    assert len(truth) == 49  # 47 pay-value differences + 2 overtime-hours differences


def test_net_pay_variance_matches_independent_total(recon500):
    rec, _, _ = recon500
    assert round(rec["Net Pay Difference"].sum(), 2) == -1017.53


def test_master_data_changes_show_old_and_new(recon500):
    rec, _, _ = recon500
    reason = rec.set_index("Employee ID")["Exception Reason"]
    assert reason["E0231"] == "Tax Code changed: S1257L → (blank)"
    assert reason["E0349"] == "NI Number changed: AA100349B → (blank)"


def test_missing_and_unexpected_employees_have_one_clear_reason(recon500):
    rec, _, _ = recon500
    reason = rec.set_index("Employee ID")["Exception Reason"]
    assert reason["E0499"] == "Missing from processed payroll"
    assert reason["E9999"] == "Unexpected employee in processed payroll"
    counts = rec.set_index("Employee ID")[["Value Mismatches", "Master Data Changes"]]
    assert counts.loc["E0499"].sum() == 0 and counts.loc["E9999"].sum() == 0


def test_without_new_options_result_matches_the_previous_behaviour(app, expected500, output500):
    """Old call signature (no master data, no overtime hours) still works and gives the old 14 exceptions."""
    old_fields = [f for f in app.RECON_FIELDS if f != "Overtime Hours" and f in expected500.columns]
    rec = app.reconcile(expected500, output500, old_fields, 0.01)
    assert int((rec["Overall Status"] == "Exception").sum()) == 14


def test_footing_flags_exactly_the_five_broken_rows(app, expected500, output500):
    out = app.check_footing(output500)
    assert sorted(out["Employee ID"]) == ["E0033", "E0128", "E0215", "E0367", "E0444"]
    assert set(out["Check / Issue"]) == {"Gross pay does not add up"}
    assert app.check_footing(expected500).empty


def test_duplicate_warnings_are_separated(app, expected500, output500, recon500):
    rec, _, _ = recon500
    check_in = app.attention_only(app.check_payroll(expected500, "in"))
    check_out = app.attention_only(app.check_payroll(output500, "out"))
    new, repeated, covered = app.separate_new_output_items(check_out, check_in, rec)
    assert sorted(new["Employee ID"]) == ["E0033", "E0128", "E0215", "E0367", "E0444"]
    assert sorted(repeated["Employee ID"]) == ["E0014", "E0077", "E0163", "E0284", "E0340"]
    assert sorted(covered["Employee ID"]) == ["E0231", "E0349"]
    assert len(new) + len(repeated) + len(covered) == len(check_out)  # nothing is lost


# ------------------------------------------------------------------ master data comparison


def _pair(**overrides):
    base = {"Employee ID": ["A1", "A2"], "Employee Name": ["Ann One", "Bob Two"], "Net Pay": [100.0, 200.0],
            "Tax Code": ["1257L", "BR"], "NI Number": ["AB123456C", "CD123456A"],
            "Bank Details": ["12-34-56 12345678", "11-11-11 87654321"]}
    exp = pd.DataFrame(base)
    out = exp.copy()
    for col, values in overrides.items():
        out[col] = values
    return exp, out


def _rec(app, exp, out, text=("Tax Code", "NI Number", "Bank Details")):
    return app.reconcile(exp, out, ["Net Pay"], 0.01, text_fields=list(text)).set_index("Employee ID")


def test_changed_tax_code_is_caught_even_when_not_blank(app):
    exp, out = _pair(**{"Tax Code": ["1257L", "0T"]})
    rec = _rec(app, exp, out)
    assert rec.loc["A2", "Overall Status"] == "Exception"
    assert rec.loc["A2", "Exception Reason"] == "Tax Code changed: BR → 0T"
    assert rec.loc["A1", "Overall Status"] == "OK"


def test_case_spacing_and_hyphens_are_not_treated_as_changes(app):
    exp, out = _pair(**{"Tax Code": ["1257l", "BR "], "NI Number": ["ab 12 34 56 c", "CD123456A"],
                        "Bank Details": ["123456 12345678", "11-11-11  87654321"]})
    rec = _rec(app, exp, out)
    assert (rec["Overall Status"] == "OK").all()


def test_excel_float_account_number_is_not_a_change(app):
    exp = pd.DataFrame({"Employee ID": ["A1"], "Employee Name": ["Ann"], "Net Pay": [1.0], "Account Number": ["12345678"]})
    out = exp.copy()
    out["Account Number"] = [12345678.0]
    rec = app.reconcile(exp, out, ["Net Pay"], 0.01, text_fields=["Account Number"])
    assert (rec["Overall Status"] == "OK").all()


def test_bank_details_are_masked_in_the_reason_text(app):
    exp, out = _pair(**{"Bank Details": ["12-34-56 12345678", "11-11-11 99999999"]})
    reason = _rec(app, exp, out).loc["A2", "Exception Reason"]
    assert reason == "Bank Details changed: …4321 → …9999"
    assert "87654321" not in reason and "99999999" not in reason


def test_name_change_is_flagged_but_case_is_not(app):
    exp, out = _pair(**{"Employee Name": ["ann one", "Bob Twoo"]})
    rec = _rec(app, exp, out)
    assert rec.loc["A1", "Overall Status"] == "OK"
    assert rec.loc["A2", "Exception Reason"] == "Employee Name changed: Bob Two → Bob Twoo"


def test_value_and_master_changes_are_both_reported(app):
    exp, out = _pair(**{"Net Pay": [100.0, 250.0], "Tax Code": ["1257L", "0T"]})
    rec = _rec(app, exp, out)
    assert rec.loc["A2", "Value Mismatches"] == 1 and rec.loc["A2", "Master Data Changes"] == 1
    assert "Net Pay mismatch (+50.00)" in rec.loc["A2", "Exception Reason"]
    assert "Tax Code changed: BR → 0T" in rec.loc["A2", "Exception Reason"]


def test_review_item_label_for_master_data_only_change(app):
    exp, out = _pair(**{"Tax Code": ["1257L", "0T"]})
    items = app.reconciliation_review_items(app.reconcile(exp, out, ["Net Pay"], 0.01, text_fields=["Tax Code"]))
    assert items["Check / Issue"].tolist() == ["Master data changed"]


def test_columns_missing_from_one_file_are_skipped_not_crashed(app):
    exp, out = _pair()
    out = out.drop(columns=["Bank Details"])
    rec = app.reconcile(exp, out, ["Net Pay"], 0.01, text_fields=["Bank Details", "Tax Code"])
    assert "Bank Details Status" not in rec.columns and "Tax Code Status" in rec.columns


# ------------------------------------------------------------------ footing checks


def _pay(**row):
    base = {"Employee ID": "E1", "Employee Name": "Ann", "Salary": 3000.0, "Overtime Pay": 200.0, "Bonus": 100.0,
            "Gross Pay": 3300.0, "PAYE": 500.0, "Employee NI": 300.0, "Pension": 150.0, "Other Deductions": 50.0,
            "Net Pay": 2300.0}
    base.update(row)
    return pd.DataFrame([base])


def test_clean_row_has_no_footing_warnings(app):
    assert app.check_footing(_pay()).empty


def test_gross_that_does_not_add_up_is_warned_with_the_sums(app):
    out = app.check_footing(_pay(**{"Gross Pay": 3400.0, "Net Pay": 2400.0}))
    assert out["Check / Issue"].tolist() == ["Gross pay does not add up"]
    assert "= 3,300.00" in out.iloc[0]["Detail"] and "+100.00" in out.iloc[0]["Detail"]
    assert out.iloc[0]["Status"] == "Warning"


def test_net_that_does_not_add_up_is_warned(app):
    out = app.check_footing(_pay(**{"Net Pay": 2250.0}))
    assert out["Check / Issue"].tolist() == ["Net pay does not add up"]


def test_penny_rounding_is_tolerated(app):
    assert app.check_footing(_pay(**{"Gross Pay": 3300.01, "Net Pay": 2300.01})).empty
    assert not app.check_footing(_pay(**{"Gross Pay": 3300.05, "Net Pay": 2300.05})).empty


def test_footing_needs_the_columns_it_uses(app):
    no_salary = _pay(**{"Gross Pay": 9999.0, "Net Pay": 9000.0}).drop(columns=["Salary"])
    assert "Gross pay does not add up" not in set(app.check_footing(no_salary)["Check / Issue"])  # no Salary: cannot run
    no_deductions = _pay(**{"Net Pay": 1.0}).drop(columns=["PAYE"])
    assert "Net pay does not add up" not in set(app.check_footing(no_deductions)["Check / Issue"])


def test_footing_can_be_switched_off(app):
    bad = _pay(**{"Gross Pay": 9999.0})
    assert "Gross pay does not add up" in set(app.check_payroll(bad, "x")["Check / Issue"])
    assert "Gross pay does not add up" not in set(app.check_payroll(bad, "x", footing=False)["Check / Issue"])


# ------------------------------------------------------------------ de-duplication


def _items(rows):
    return pd.DataFrame(rows, columns=["Source", "Status", "Employee ID", "Employee Name", "Check / Issue", "Detail"])


def test_identical_warning_on_input_is_not_repeated_but_new_one_is_kept(app):
    inp = _items([["CHECK", "Warning", "E1", "A", "Missing Tax Code", "Tax code is blank."]])
    out = _items([["CHECK", "Warning", "E1", "A", "Missing Tax Code", "Tax code is blank."],
                  ["CHECK", "Warning", "E2", "B", "Missing Tax Code", "Tax code is blank."]])
    new, repeated, covered = app.separate_new_output_items(out, inp, None)
    assert new["Employee ID"].tolist() == ["E2"] and repeated["Employee ID"].tolist() == ["E1"] and covered.empty


def test_blank_on_both_sides_is_a_repeat_not_a_change(app):
    """Blank in the input and blank in the output is not a master-data change, so the warning is the repeat."""
    exp, out = _pair(**{"Tax Code": ["1257L", "BR"]})
    exp.loc[0, "Tax Code"] = None
    out.loc[0, "Tax Code"] = None
    rec = app.reconcile(exp, out, ["Net Pay"], 0.01, text_fields=["Tax Code"])
    assert (rec["Overall Status"] == "OK").all()
    inp = _items([["CHECK", "Warning", "A1", "Ann One", "Missing Tax Code", "Tax code is blank."]])
    new, repeated, covered = app.separate_new_output_items(inp.copy(), inp, rec)
    assert len(repeated) == 1 and new.empty and covered.empty


def test_missing_field_warning_stays_when_it_is_not_shown_as_a_change(app):
    """If the reconciliation does not report the blank as a change, the warning must not be hidden."""
    out = _items([["CHECK", "Warning", "E5", "A", "Missing Tax Code", "Tax code is blank."]])
    rec = pd.DataFrame({"Employee ID": ["E5"], "Tax Code Status": ["OK"]})
    new, repeated, covered = app.separate_new_output_items(out, _items([]), rec)
    assert len(new) == 1 and repeated.empty and covered.empty


def test_empty_output_items(app):
    new, repeated, covered = app.separate_new_output_items(_items([]), _items([]), None)
    assert new.empty and repeated.empty and covered.empty


# ------------------------------------------------------------------ bulk review


def _queue(app):
    rows = [["CHECK", "Warning", "E1", "A", "High overtime pay", "d1"],
            ["CHECK", "Warning", "E2", "B", "High overtime pay", "d2"],
            ["CHECK", "Warning", "E3", "C", "Missing Tax Code", "d3"],
            ["CHECK", "Fail", "E4", "D", "Duplicate employee", "d4"],
            ["RECONCILE", "Fail", "E5", "E", "Reconciliation exception", "d5"]]
    q = _items(rows)
    q.insert(0, "Exception ID", [f"CARRR-{i+1:04d}" for i in range(len(q))])
    q["Reviewed"] = False
    q["Review Status"] = "Open"
    q["Reviewer Note"] = ""
    return q


def test_bulk_accept_marks_selected_warnings_with_the_note(app):
    q = _queue(app)
    out = app.apply_bulk_decision(q, q["Check / Issue"].eq("High overtime pay"), "Accepted", "Checked against timesheets")
    assert out.loc[out["Employee ID"].isin(["E1", "E2"]), "Reviewed"].all()
    assert (out.loc[out["Employee ID"].isin(["E1", "E2"]), "Review Status"] == "Accepted").all()
    assert (out.loc[out["Employee ID"].isin(["E1", "E2"]), "Reviewer Note"] == "Checked against timesheets").all()
    assert not out.loc[out["Employee ID"].isin(["E3", "E4", "E5"]), "Reviewed"].any()
    assert q["Reviewed"].sum() == 0  # the original table is untouched


@pytest.mark.parametrize("note", ["", "   ", None])
def test_bulk_decision_requires_a_note(app, note):
    q = _queue(app)
    with pytest.raises(ValueError, match="note"):
        app.apply_bulk_decision(q, q["Status"].eq("Warning"), "Accepted", note)


def test_bulk_decision_refuses_fails_and_payroll_mismatches(app):
    q = _queue(app)
    with pytest.raises(ValueError, match="individually"):
        app.apply_bulk_decision(q, q["Employee ID"].eq("E4"), "Resolved", "ok")
    with pytest.raises(ValueError, match="individually"):
        app.apply_bulk_decision(q, q["Employee ID"].eq("E5"), "Resolved", "ok")
    with pytest.raises(ValueError, match="individually"):
        app.apply_bulk_decision(q, q["Status"].isin(["Warning", "Fail"]), "Resolved", "ok")


def test_bulk_decision_rejects_empty_selection_and_bad_outcome(app):
    q = _queue(app)
    with pytest.raises(ValueError, match="No open items"):
        app.apply_bulk_decision(q, q["Employee ID"].eq("nobody"), "Accepted", "ok")
    with pytest.raises(ValueError, match="Resolved or Accepted"):
        app.apply_bulk_decision(q, q["Employee ID"].eq("E1"), "Open", "ok")


def test_bulk_note_is_appended_to_an_existing_note(app):
    q = _queue(app)
    q.loc[0, "Reviewer Note"] = "Looked at this one first"
    out = app.apply_bulk_decision(q, q["Employee ID"].eq("E1"), "Accepted", "bulk")
    assert out.loc[0, "Reviewer Note"] == "Looked at this one first | bulk"


def test_decisions_survive_a_rerun_and_do_not_leak_to_other_items(app):
    q = _queue(app)
    decided = app.apply_bulk_decision(q, q["Employee ID"].eq("E1"), "Accepted", "ok")
    saved = app.snapshot_decisions(decided)
    assert list(saved) and len(saved) == 1
    rerun = _queue(app).iloc[::-1].reset_index(drop=True)  # different order, same items
    restored = app.apply_saved_decisions(rerun, saved)
    assert restored.loc[restored["Employee ID"] == "E1", "Review Status"].item() == "Accepted"
    assert (restored.loc[restored["Employee ID"] != "E1", "Review Status"] == "Open").all()


def test_a_changed_item_is_not_pre_approved(app):
    """Same employee and issue but a different detail (the numbers changed) must come back as Open."""
    q = _queue(app)
    saved = app.snapshot_decisions(app.apply_bulk_decision(q, q["Employee ID"].eq("E1"), "Accepted", "ok"))
    changed = _queue(app)
    changed.loc[changed["Employee ID"] == "E1", "Detail"] = "d1 but different amount"
    restored = app.apply_saved_decisions(changed, saved)
    assert (restored["Review Status"] == "Open").all() and not restored["Reviewed"].any()
