"""Smoke test: run the whole app.py script top to bottom on the 500-employee sample files.

This does not draw anything. A small stand-in for the `streamlit` module records what the app would show
(metrics, tables, messages) and answers widgets with their defaults, so wiring mistakes in the screens
(wrong variable names, bad columns, crashes in the Review/Release tabs) are caught without a browser.
It cannot judge layout or colours: check those by eye on a phone.
"""
import io
import sys
import types
from contextlib import contextmanager
from pathlib import Path

import pandas as pd
import pytest

from conftest import DATA, ROOT


class StopApp(Exception):
    pass


class FakeUpload(io.BytesIO):
    def __init__(self, path):
        super().__init__(Path(path).read_bytes())
        self.name = Path(path).name
        self.size = len(self.getvalue())


class Recorder:
    def __init__(self):
        self.metrics, self.tables, self.messages, self.downloads = {}, [], [], []
        self.session_state = {}


class FakeStreamlit(types.ModuleType):
    """Just enough of streamlit for app.py; widgets answer with their default values."""

    def __init__(self, rec, input_path, output_path, button_presses=(), answers=None):
        super().__init__("streamlit")
        self.rec, self._in, self._out = rec, input_path, output_path
        self._answers = answers or {}
        self.session_state = rec.session_state
        self.column_config = types.SimpleNamespace(
            CheckboxColumn=lambda *a, **k: None, SelectboxColumn=lambda *a, **k: None, TextColumn=lambda *a, **k: None)
        self._pressed = set(button_presses)

    # --- layout
    def _ctx(self):
        @contextmanager
        def cm(*a, **k):
            yield self
        return cm()

    def columns(self, spec, **k):
        return [self for _ in range(spec if isinstance(spec, int) else len(spec))]

    def tabs(self, names):
        return [self._ctx() for _ in names]

    def expander(self, *a, **k):
        return self._ctx()

    # --- display (recorded)
    def metric(self, label, value, *a, **k):
        self.rec.metrics[label] = value

    def dataframe(self, df, *a, **k):
        self.rec.tables.append(df)

    def _msg(kind):
        def f(self, text="", *a, **k):
            self.rec.messages.append((kind, str(text)))
        return f

    success, info, warning, error, caption, markdown, write, text = (
        _msg("success"), _msg("info"), _msg("warning"), _msg("error"), _msg("caption"), _msg("markdown"),
        _msg("write"), _msg("text"))

    def set_page_config(self, *a, **k): pass
    def divider(self): pass
    def download_button(self, label, data, *a, **k): self.rec.downloads.append((label, data, k.get("file_name")))

    # --- inputs (default answers)
    def file_uploader(self, label, *a, accept_multiple_files=False, **k):
        if accept_multiple_files:
            return [FakeUpload(self._in)]
        return FakeUpload(self._out)

    def number_input(self, label, min_value=None, value=0, *a, **k): return value
    def text_input(self, label, *a, **k): return self._answers.get(label, "")
    def date_input(self, label, value=None, *a, **k): return value
    def selectbox(self, label, options, index=0, *a, **k): return list(options)[index]
    def multiselect(self, label, options, default=None, *a, **k):
        return list(self._answers.get(label, default or []))
    def checkbox(self, label, value=False, *a, **k): return value
    def data_editor(self, df, *a, **k): return df
    def button(self, label, *a, **k): return any(p in label for p in self._pressed)

    def stop(self): raise StopApp()
    def rerun(self): raise StopApp()


def run_app(monkeypatch, expected_path, output_path, session_state=None, answers=None, presses=()):
    """One run of the script. Pass the same session_state dict to simulate the page reloading in one session."""
    rec = Recorder()
    if session_state is not None:
        rec.session_state = session_state
    fake = FakeStreamlit(rec, expected_path, output_path, button_presses=presses, answers=answers)
    monkeypatch.setitem(sys.modules, "streamlit", fake)
    source = (ROOT / "app.py").read_text()
    try:
        exec(compile(source, str(ROOT / "app.py"), "exec"), {"__name__": "__main__"})
    except StopApp:
        pass
    return rec


def test_whole_app_runs_on_the_500_employee_files(monkeypatch):
    rec = run_app(monkeypatch, DATA / "expected_500.xlsx", DATA / "output_500.xlsx")
    errors = [m for m in rec.messages if m[0] == "error"]
    assert not errors, errors

    m = rec.metrics
    assert m["Employees checked"] == 501
    assert m["Matched / OK"] == 485
    assert m["Employees with exceptions"] == 16
    assert m["Net pay variance"] == "£-1,017.53"
    assert m["Missing IDs"] == 1 and m["Unexpected IDs"] == 1
    assert m["Payroll mismatches"] == 16
    # 80 before-payroll + 5 new footing warnings on the output + 16 reconciliation exceptions
    assert m["Total review items still open"] == 101
    # The release gate cannot be cleared with nothing reviewed
    assert any("HOLD PAYROLL" in text for kind, text in rec.messages if kind == "markdown")
    assert m["Open mismatches"] == 16 and m["Open fails"] == 0 and m["Open warnings"] == 85

    # The note about removed duplicates is shown
    assert any("5 output warning(s) were already raised on the input" in t and "2 are shown as master-data" in t
               for _, t in rec.messages)

    # Audit report is produced and has the new columns/sheets with sensible content
    label, data, filename = rec.downloads[0]
    sheets = pd.read_excel(io.BytesIO(data), sheet_name=None)
    assert {"Summary", "Check", "Analyse", "Reconciliation", "Review", "Release"} <= set(sheets)
    recon = sheets["Reconciliation"]
    assert {"Value Mismatches", "Master Data Changes", "Tax Code Status", "Overtime Hours Status"} <= set(recon.columns)
    summary = sheets["Summary"].set_index("Metric")["Value"]
    assert summary["Output Warnings Already Raised On Input"] == 5
    assert summary["Final Decision"] == "HOLD PAYROLL"


def test_app_runs_when_files_have_no_master_data_columns(monkeypatch, tmp_path):
    """A minimal file pair (ID, name, a few pay columns) must not crash any of the new code paths."""
    exp = pd.DataFrame({"Employee ID": ["A1", "A2"], "Employee Name": ["Ann", "Bob"],
                        "Gross Pay": [1000.0, 2000.0], "Net Pay": [800.0, 1600.0]})
    out = exp.copy()
    out.loc[1, "Net Pay"] = 1500.0
    pe, po = tmp_path / "exp.xlsx", tmp_path / "out.xlsx"
    exp.to_excel(pe, index=False)
    out.to_excel(po, index=False)
    rec = run_app(monkeypatch, pe, po)
    assert not [m for m in rec.messages if m[0] == "error"]
    assert rec.metrics["Employees checked"] == 2 and rec.metrics["Employees with exceptions"] == 1
    assert rec.downloads  # reached the end of the Release tab


def test_review_decisions_reset_when_the_files_change(monkeypatch, tmp_path):
    """Saved decisions belong to one pair of files; a different upload must start clean."""
    rec1 = run_app(monkeypatch, DATA / "expected_500.xlsx", DATA / "output_500.xlsx")
    fingerprint = rec1.session_state["carrr_review_fp"]
    assert rec1.session_state["carrr_review_decisions"] == {}
    rec2 = Recorder()
    rec2.session_state.update({"carrr_review_fp": "some-other-upload", "carrr_review_decisions": {"x": {}},
                               "carrr_review_version": 3})
    fake = FakeStreamlit(rec2, DATA / "expected_500.xlsx", DATA / "output_500.xlsx")
    monkeypatch.setitem(sys.modules, "streamlit", fake)
    try:
        exec(compile((ROOT / "app.py").read_text(), str(ROOT / "app.py"), "exec"), {"__name__": "__main__"})
    except StopApp:
        pass
    assert rec2.session_state["carrr_review_fp"] == fingerprint
    assert rec2.session_state["carrr_review_decisions"] == {} and rec2.session_state["carrr_review_version"] == 0


def test_bulk_accept_end_to_end_and_survives_a_reload(monkeypatch):
    state = {}
    answers = {"Warning types": ["High overtime pay"], "Reviewer note (required)": "Checked against timesheets"}
    run_app(monkeypatch, DATA / "expected_500.xlsx", DATA / "output_500.xlsx", state, answers, presses=("Apply to",))
    decisions = state["carrr_review_decisions"]
    assert len(decisions) == 72 and state["carrr_review_version"] == 1
    assert all(d["Review Status"] == "Accepted" and d["Reviewer Note"] == "Checked against timesheets"
               for d in decisions.values())

    rec = run_app(monkeypatch, DATA / "expected_500.xlsx", DATA / "output_500.xlsx", state)  # page reloads
    m = rec.metrics
    assert m["Total review items still open"] == 101 - 72
    assert m["Open warnings"] == 85 - 72 and m["Open mismatches"] == 16
    assert m["Accepted without note"] == 0
    summary = pd.read_excel(io.BytesIO(rec.downloads[0][1]), sheet_name="Summary").set_index("Metric")["Value"]
    assert summary["Accepted Review Items"] == 72 and summary["Final Decision"] == "HOLD PAYROLL"  # mismatches still block


def test_bulk_accept_without_a_note_changes_nothing(monkeypatch):
    state = {}
    answers = {"Warning types": ["High overtime pay"], "Reviewer note (required)": ""}
    rec = run_app(monkeypatch, DATA / "expected_500.xlsx", DATA / "output_500.xlsx", state, answers, presses=("Apply to",))
    assert state.get("carrr_review_decisions") == {}
    assert any(kind == "error" and "note is required" in text for kind, text in rec.messages)
