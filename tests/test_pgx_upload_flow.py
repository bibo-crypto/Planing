"""The "Upload PG-X Partita file" flow of the shared-orders window, run end to end.

The window itself is replaced by a tiny stub, Tk's after() runs callbacks at once and background
threads run inline, so what is exercised is the real _upload_pgx_gg_file logic against real
workbooks: which rows get a raw batch, what the user is asked, and which rows are removed from
the uploaded file afterwards.
"""
import logging
import types
from pathlib import Path

import openpyxl
import pandas as pd
import pytest

pytest.importorskip("tkinter")

from exporters.biglietti_exporter import OrderRecord, append_create_excel, load_create_excel_records  # noqa: E402
from gui.tabs import shared_order_actions as actions  # noqa: E402


def _record(partita_col, article="C130027S"):
    return OrderRecord(
        customer_code="3009", customer_name="ELVY", article=article, description="", additional_raw="",
        color_code="5305", color_name="EL-281311", order_no="7777", order_row="1",
        colored_batch=partita_col, raw_batch="PG-X", quantity_cones=32, raw_weight=30,
        dispo="D-00505450-001", bagno="S940",
    )


class Host(actions.SharedOrdersActionsMixin):
    """The little the mixin needs from the window that normally owns it."""

    def __init__(self, shared_excel, summary):
        self.shared_excel_path = Path(shared_excel)
        self._summary = summary
        self._logger = logging.getLogger("test.pgx")
        self.pending = []
        self.callback_errors = []

    def after(self, _ms, callback=None, *args):
        # Like Tk: the callback does not run inside the caller; it runs later from the event loop,
        # and an exception in it is reported there instead of reaching the code that scheduled it.
        if callback:
            self.pending.append((callback, args))

    def drain(self):
        while self.pending:
            callback, args = self.pending.pop(0)
            try:
                callback(*args)
            except Exception as exc:  # noqa: BLE001 -- what report_callback_exception would swallow
                self.callback_errors.append(exc)

    def _load_common_sources(self):
        return {}, {}, {}, {}, self._summary


@pytest.fixture
def pgx(tmp_path, monkeypatch):
    class InlineThread:
        def __init__(self, target=None, args=(), kwargs=None, daemon=None):
            self._target, self._args, self._kwargs = target, args, kwargs or {}

        def start(self):
            self._target(*self._args, **self._kwargs)

    monkeypatch.setattr(actions, "threading", types.SimpleNamespace(Thread=InlineThread))
    shown = {"info": [], "error": [], "asked": []}
    answers = {"missing_batches": True}

    def ask(title, text, **_kw):
        shown["asked"].append(text)
        if isinstance(answers["missing_batches"], Exception):
            raise answers["missing_batches"]
        return answers["missing_batches"]

    monkeypatch.setattr(actions.messagebox, "showinfo", lambda title, text, **k: shown["info"].append(text))
    monkeypatch.setattr(actions.messagebox, "showerror", lambda title, text, **k: shown["error"].append(text))
    monkeypatch.setattr(actions.messagebox, "askyesno", ask)

    shared = tmp_path / "shared.xlsx"
    append_create_excel(shared, [_record("157890"), _record("157891"), _record("157892")], "ELVY")
    summary = pd.DataFrame([
        {"articolo": "G130027S", "partita": "158694", "mag_rocche": 32},     # right raw article for C130027S
        {"articolo": "G999999S", "partita": "158700", "mag_rocche": 10},     # exists, but for another article
    ])
    upload = tmp_path / "upload.xlsx"

    def write_upload(rows):
        wb = openpyxl.Workbook()
        wb.active.append(["Partita Col", "Partita GG"])
        for row in rows:
            wb.active.append(list(row))
        wb.save(upload)

    def run(rows):
        write_upload(rows)
        monkeypatch.setattr(actions.filedialog, "askopenfilename", lambda **k: str(upload))
        refreshed = []
        host.pgx_refreshed = refreshed
        host._upload_pgx_gg_file(object(), {}, lambda datasets: refreshed.append(datasets), lambda: None, None)
        host.drain()

    def upload_rows():
        wb = openpyxl.load_workbook(upload, data_only=True)
        try:
            return [tuple(row) for row in wb.active.values][1:]
        finally:
            wb.close()

    def raw_batches():
        return {r.colored_batch: r.raw_batch for r in load_create_excel_records(shared, sheet_name="PG-X")}

    host = Host(shared, summary)
    return types.SimpleNamespace(host=host, shown=shown, answers=answers, run=run, upload_rows=upload_rows,
                                 raw_batches=raw_batches, shared=shared)


def test_valid_rows_are_assigned_and_only_those_leave_the_uploaded_file(pgx):
    pgx.run([
        ("157890", "158694"),      # batch in stock for the right article -> assigned
        ("157891", "999999"),      # batch not in Magazino at all -> user is asked, says yes
        ("157892", "158700"),      # batch exists but for a different raw article -> refused
        ("157893", "158694"),      # Partita Col that is not in PG-X any more
    ])

    assert pgx.raw_batches() == {"157890": "158694", "157891": "999999", "157892": "PG-X"}
    assert pgx.upload_rows() == [("157892", "158700"), ("157893", "158694")]       # the failures stay for the user
    assert len(pgx.shown["asked"]) == 1 and "157891" in pgx.shown["asked"][0]
    message = pgx.shown["info"][0]
    assert "Assigned 2 row(s)" in message and "Removed 2 processed row(s)" in message
    assert "157892" in message and "not for the required raw article" in message
    assert "157893" in message and "no matching Partita Col" in message
    assert load_create_excel_records(pgx.shared, sheet_name="Orders") == []        # nothing is moved to Orders
    assert pgx.host._pgx_upload_running is False
    assert len(pgx.host.pgx_refreshed) == 1


def test_declining_the_missing_batch_warning_leaves_that_row_untouched(pgx):
    pgx.answers["missing_batches"] = False

    pgx.run([("157890", "158694"), ("157891", "999999")])

    assert pgx.raw_batches() == {"157890": "158694", "157891": "PG-X", "157892": "PG-X"}
    assert pgx.upload_rows() == [("157891", "999999")]
    assert "skipped because Partita GG 999999 is not in Magazino" in pgx.shown["info"][0]
    assert pgx.host._pgx_upload_running is False


def test_a_failure_while_asking_does_not_lock_later_uploads(pgx):
    pgx.answers["missing_batches"] = RuntimeError("the window was closed")

    pgx.run([("157891", "999999")])

    assert pgx.shown["error"] == ["the window was closed"]
    assert pgx.host.callback_errors == []                         # handled, not left to the event loop
    assert pgx.host._pgx_upload_running is False                 # released: the next upload is not ignored
    pgx.answers["missing_batches"] = True
    pgx.run([("157891", "999999")])
    assert pgx.raw_batches()["157891"] == "999999"


def test_starting_a_second_upload_while_one_runs_explains_why_nothing_happens(pgx):
    pgx.host._pgx_upload_running = True

    pgx.run([("157890", "158694")])

    assert pgx.raw_batches()["157890"] == "PG-X"
    assert "already running" in pgx.shown["info"][0]


def test_file_without_the_required_headers_is_rejected_before_anything_changes(pgx, monkeypatch, tmp_path):
    bad = tmp_path / "bad.xlsx"
    wb = openpyxl.Workbook()
    wb.active.append(["Colore", "Lotto"])
    wb.active.append(["1", "2"])
    wb.save(bad)
    monkeypatch.setattr(actions.filedialog, "askopenfilename", lambda **k: str(bad))

    pgx.host._upload_pgx_gg_file(object(), {}, lambda d: None, lambda: None, None)
    pgx.host.drain()

    assert "Partita Col and Partita GG" in pgx.shown["error"][0]
    assert pgx.host._pgx_upload_running is False if hasattr(pgx.host, "_pgx_upload_running") else True
    assert pgx.raw_batches() == {"157890": "PG-X", "157891": "PG-X", "157892": "PG-X"}


def test_blank_rows_in_the_uploaded_file_do_not_shift_which_rows_are_removed(pgx):
    pgx.run([("", ""), ("157890", "158694"), (None, None), ("157892", "158700"), ("157891", "158694")])

    # 157890 and 157891 are assigned; 157892 (wrong raw article) and the blank rows are what remains
    assert pgx.raw_batches() == {"157890": "158694", "157891": "158694", "157892": "PG-X"}
    remaining = [row for row in pgx.upload_rows() if any(row)]
    assert remaining == [("157892", "158700")]
