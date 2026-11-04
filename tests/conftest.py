"""Load CARRR's pure functions from app.py without running the Streamlit UI.

app.py is a single Streamlit script, so importing it would execute the whole page. Instead we parse it,
keep only imports, constants and function definitions, and execute those. Anything that needs the
`streamlit` package is replaced by a stub, so the tests run with just pandas/openpyxl/pytest.
"""
import ast
import warnings
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock

import pandas as pd
import pytest

ROOT = Path(__file__).resolve().parent.parent
DATA = Path(__file__).resolve().parent / "data"


def _load_app_namespace():
    src = (ROOT / "app.py").read_text()
    tree = ast.parse(src)
    ns = {"__name__": "carrr_app_under_test"}
    for node in tree.body:
        wanted = False
        if isinstance(node, (ast.FunctionDef, ast.Import, ast.ImportFrom)):
            wanted = True
        elif isinstance(node, ast.Try):
            # Only optional-import guards (try: import x / except ImportError: x = None), never UI code.
            wanted = all(isinstance(n, (ast.Import, ast.ImportFrom)) for n in node.body)
        elif isinstance(node, ast.Assign):
            wanted = all(isinstance(t, ast.Name) and t.id.lstrip("_").replace("_", "").isupper() for t in node.targets)
        if not wanted:
            continue
        code = compile(ast.Module(body=[node], type_ignores=[]), str(ROOT / "app.py"), "exec")
        try:
            exec(code, ns)
        except ImportError:
            if isinstance(node, (ast.Import, ast.ImportFrom)):
                for alias in node.names:
                    ns[alias.asname or alias.name.split(".")[0]] = MagicMock()
    return ns


@pytest.fixture(scope="session")
def app():
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        return SimpleNamespace(**_load_app_namespace())


def _prepare(app, path):
    raw = pd.read_excel(path, header=None, dtype=object)  # same as read_upload() for .xlsx
    df, _ = app.dataframe_from_raw(raw)
    df = app.clean_columns(df)
    mapping, _ = app.build_auto_mapping(df)
    return app.apply_mapping(df, mapping)


@pytest.fixture(scope="session")
def expected500(app):
    return _prepare(app, DATA / "expected_500.xlsx")


@pytest.fixture(scope="session")
def output500(app):
    return _prepare(app, DATA / "output_500.xlsx")


@pytest.fixture(scope="session")
def recon500(app, expected500, output500):
    fields = [f for f in app.RECON_FIELDS if f in expected500.columns and f in output500.columns]
    text = [f for f in app.MASTER_FIELDS if f in expected500.columns and f in output500.columns]
    return app.reconcile(expected500, output500, fields, 0.01, text_fields=text), fields, text
