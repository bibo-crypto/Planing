"""Regression tests for the problems found while reviewing 1.1.8 (SQLite-first + lazy pages)."""
import ast
import os
import time
from pathlib import Path

import pandas as pd
import pytest

from utility import situazione_db as db

PROJECT = Path(__file__).resolve().parent.parent


@pytest.fixture
def sqlite_db(tmp_path, monkeypatch):
    monkeypatch.setattr(db, "DB_PATH", str(tmp_path / "t.sqlite3"))
    db.init_db()
    return tmp_path


# ---------------------------------------------------------------- PyInstaller: lazy pages must be real imports

LAZY_PAGES = (
    "gui.tabs.biglietti_tab", "gui.tabs.kamal_tab", "gui.tabs.ordine_med_tab",
    "gui.tabs.situazione_settimana_tab", "gui.tabs.master_data_tab",
)


def test_lazily_built_pages_are_reachable_through_import_statements():
    """PyInstaller bundles what it finds in `import` statements; a module named only inside
    __import__("...") is left out of the exe and its page opens blank there."""
    tree = ast.parse((PROJECT / "gui" / "gui.py").read_text(encoding="utf-8"))
    imported = {node.module for node in ast.walk(tree) if isinstance(node, ast.ImportFrom) and node.module}
    for module in LAZY_PAGES:
        assert module in imported, f"{module} is not imported by an import statement in gui/gui.py"
    string_imports = [node for node in ast.walk(tree)
                      if isinstance(node, ast.Call) and getattr(node.func, "id", "") == "__import__"]
    assert string_imports == []


# ---------------------------------------------------------------- scheduled PG-X report

def test_pgx_schedule_needs_enabled_and_a_recipient():
    from gui.gui import pgx_schedule_enabled

    assert pgx_schedule_enabled({"pgx_report_schedule": {"enabled": True, "recipient": "a@b.com"}})
    assert not pgx_schedule_enabled({"pgx_report_schedule": {"enabled": True, "recipient": "  "}})
    assert not pgx_schedule_enabled({"pgx_report_schedule": {"enabled": False, "recipient": "a@b.com"}})
    assert not pgx_schedule_enabled({})
    assert not pgx_schedule_enabled(None)


# ---------------------------------------------------------------- SQLite snapshots keep their types

def _situazione_frames():
    orders = pd.DataFrame({
        "cliente": ["ELVY", "ELVY", "MED"], "articolo": ["C130027S", "C130027S", "C010003S"],
        "codice": ["5305", "5305", "77"], "colore": ["EL-1", "EL-1", "MED-2"], "ordine": ["1", "2", "3"],
        "riga": ["1", "1", "1"], "data": pd.to_datetime(["2026-09-01", "2026-09-02", None]),
        "consegna": pd.to_datetime(["2026-10-01", "2026-10-05", "2026-10-09"]),
        "partita": ["157890", "157891", "157892"], "rocche": [32, 40, 10.5], "comment": ["", "", ""],
        "cq": ["S940", "S941", "S942"], "prezzo": ["", "", ""]})
    dfm = pd.DataFrame({"partita": ["157890", "157891"], "mc": [12.0, 7.0], "bagno": ["S940", "S941"], "articolo": ["A", "B"]})
    prod = pd.DataFrame({"partita": ["157890"], "end_prod": ["01/10/2026"]})
    cop = pd.DataFrame({"bagno": ["S941"], "planedate": ["PROD.03/10/2026"], "machine": ["M7"],
                        "batch_start": pd.to_datetime(["2026-10-01"])})
    usc = pd.DataFrame({"partita": ["157890"], "data_uscita": pd.to_datetime(["2026-10-03"])})
    qua = pd.DataFrame({"partita": ["157890"], "data_qualita": pd.to_datetime(["2026-10-02"])})
    return dict(wincoint=orders, dfm=dfm, data_prod=prod, copertura=cop, uscita=usc, qualita=qua)


def test_situazione_computes_the_same_result_from_frames_restored_out_of_sqlite(sqlite_db):
    from calculate import situazione as S

    frames = _situazione_frames()
    for name, frame in frames.items():
        db.save_frame_cache(name, frame)
    restored = {name: db.load_frame_cache(name) for name in frames}

    assert str(restored["wincoint"]["data"].dtype).startswith("datetime64")
    assert list(restored["wincoint"].columns) == list(frames["wincoint"].columns)
    args = ("wincoint", "dfm", "data_prod", "copertura", "uscita", "qualita")
    fresh = S.compute_situation(*[frames[k] for k in args], codes_map={}, old_comments={})
    again = S.compute_situation(*[restored[k] for k in args], codes_map={}, old_comments={})
    assert fresh.reset_index(drop=True).astype(str).equals(again.reset_index(drop=True).astype(str))


def test_saving_an_identical_frame_again_reports_no_change_and_keeps_the_rows(sqlite_db):
    frame = _situazione_frames()["wincoint"]
    assert db.save_frame_cache("wincoint", frame) is True
    conn = db.get_conn()
    first_rowids = [r[0] for r in conn.execute("SELECT rowid FROM source_snapshot_row ORDER BY rowid")]
    conn.close()

    assert db.save_frame_cache("wincoint", frame.copy()) is False

    conn = db.get_conn()
    assert [r[0] for r in conn.execute("SELECT rowid FROM source_snapshot_row ORDER BY rowid")] == first_rowids
    conn.close()
    changed = frame.copy()
    changed.loc[0, "rocche"] = 99
    assert db.save_frame_cache("wincoint", changed) is True
    assert db.load_frame_cache("wincoint").loc[0, "rocche"] == 99


def test_snapshots_written_before_dtypes_were_recorded_still_restore_their_dates(sqlite_db):
    frame = _situazione_frames()["uscita"]
    db.save_frame_cache("uscita", frame)
    conn = db.get_conn()
    conn.execute("UPDATE source_snapshot SET column_meta=NULL WHERE source_name='uscita'")   # an old snapshot
    conn.commit(); conn.close()

    assert str(db.load_frame_cache("uscita")["data_uscita"].dtype).startswith("datetime64")


def test_an_empty_frame_keeps_its_columns(sqlite_db):
    db.save_frame_cache("empty", pd.DataFrame({"partita": pd.Series(dtype=str), "data": pd.Series(dtype="datetime64[ns]")}))

    restored = db.load_frame_cache("empty")

    assert list(restored.columns) == ["partita", "data"] and restored.empty


# ---------------------------------------------------------------- Situazione startup restore

class _Row:
    def __init__(self):
        self.status = None

    def set_status(self, ok, text):
        self.status = (ok, text)


def _host():
    from gui.tabs.situazione_sources import SOURCE_ORDER, SituationSourcesMixin

    class Host(SituationSourcesMixin):
        def __init__(self):
            self.loaded_frames = {}
            self.source_rows = {key: _Row() for key in SOURCE_ORDER}
            self.codes_row, self.listini_row = _Row(), _Row()
            self._copertura_revision = 0
            self.refreshes = []

        def after(self, _ms, callback=None, *args):
            if callback:
                callback(*args)

        def _on_refresh(self, **kwargs):
            self.refreshes.append(kwargs)

    return Host()


@pytest.fixture
def no_dialogs(monkeypatch):
    from tkinter import messagebox

    shown = []
    for name in ("showwarning", "showinfo", "showerror"):
        monkeypatch.setattr(messagebox, name, lambda *a, _n=name, **k: shown.append((_n, a)))
    return shown


def test_startup_with_some_sources_missing_stays_quiet(sqlite_db, no_dialogs):
    """Upgrading from a version that only snapshotted a few sources must not greet the operator
    with a 'Missing files' box at every launch."""
    pytest.importorskip("tkinter")
    db.save_frame_cache("copertura", _situazione_frames()["copertura"])
    host = _host()

    host._auto_restore_saved_files()

    assert list(host.loaded_frames) == ["copertura"]
    assert host.refreshes == []
    assert no_dialogs == []


def test_startup_with_every_source_saved_rebuilds_the_table_once(sqlite_db, no_dialogs):
    pytest.importorskip("tkinter")
    for name, frame in _situazione_frames().items():
        db.save_frame_cache(name, frame)
    host = _host()

    host._auto_restore_saved_files()

    assert set(host.loaded_frames) == set(_situazione_frames())
    assert host.refreshes == [{"preserve_comment_history": True}]
    assert no_dialogs == []


def test_a_source_without_snapshot_is_read_once_from_its_saved_file_and_then_lives_in_sqlite(
        sqlite_db, no_dialogs, monkeypatch):
    pytest.importorskip("tkinter")
    from parsers import situazione_loaders as loaders

    frames = _situazione_frames()
    for name, frame in frames.items():
        if name != "uscita":
            db.save_frame_cache(name, frame)
    saved_file = sqlite_db / "uscita.xlsx"
    saved_file.write_text("x")
    db.save_upload("uscita", "uscita.xlsx", 1, "ok", "ok", file_path=str(saved_file))
    monkeypatch.setitem(loaders.LOADERS, "uscita", ("Uscita", lambda path: (frames["uscita"], [])))
    from gui.tabs import situazione_sources
    import types

    class InlineThread:       # run the one-time upgrade read on the test thread
        def __init__(self, target=None, daemon=None, **_):
            self._target = target

        def start(self):
            self._target()

    monkeypatch.setattr(situazione_sources, "threading", types.SimpleNamespace(Thread=InlineThread))
    host = _host()

    host._auto_restore_saved_files()

    assert "uscita" in host.loaded_frames
    assert db.load_frame_cache("uscita") is not None            # next start: SQLite only
    assert len(host.refreshes) == 1 and no_dialogs == []


# ---------------------------------------------------------------- snapshots older than their file are not served

@pytest.fixture
def registered_file(sqlite_db, monkeypatch):
    """A Magazino export registered in the shared path cache, with a SQLite snapshot built from it."""
    from utility import magazino_cache

    path = sqlite_db / "magazino.xlsx"
    path.write_text("monday")
    monkeypatch.setattr(magazino_cache, "load_magazino_cache", lambda: {"source_path": str(path)})
    return path


def test_snapshot_is_served_while_it_matches_the_registered_file(registered_file):
    from utility import source_manager

    source_manager.save("magazino_summary", pd.DataFrame({"partita": ["1"], "mag_rocche": [5]}), registered_file)

    assert source_manager.is_current("magazino_summary")
    assert source_manager.load("magazino_summary") is not None


def test_a_file_registered_by_another_page_makes_the_snapshot_stale(registered_file, monkeypatch):
    from utility import magazino_cache, source_manager

    source_manager.save("magazino_summary", pd.DataFrame({"partita": ["1"], "mag_rocche": [5]}), registered_file)
    other = registered_file.parent / "other_export.xlsx"
    other.write_text("chosen on the Kamal page")
    monkeypatch.setattr(magazino_cache, "load_magazino_cache", lambda: {"source_path": str(other)})

    assert not source_manager.is_current("magazino_summary")
    assert source_manager.load("magazino_summary") is None       # callers fall back to reading the new file


def test_todays_export_overwriting_the_same_file_name_makes_the_snapshot_stale(registered_file):
    from utility import source_manager

    source_manager.save("magazino_summary", pd.DataFrame({"partita": ["1"], "mag_rocche": [5]}), registered_file)
    stat = registered_file.stat()
    registered_file.write_text("tuesday's export")
    os.utime(registered_file, ns=(stat.st_atime_ns, stat.st_mtime_ns + 5_000_000_000))

    assert source_manager.load("magazino_summary") is None


def test_the_snapshot_is_trusted_when_the_workbook_is_gone(registered_file):
    from utility import source_manager

    source_manager.save("magazino_summary", pd.DataFrame({"partita": ["1"], "mag_rocche": [5]}), registered_file)
    registered_file.unlink()

    assert source_manager.is_current("magazino_summary")
    assert source_manager.load("magazino_summary") is not None   # SQLite-first: the file is optional


def test_sources_with_no_registered_file_are_never_stale(sqlite_db):
    from utility import source_manager

    source_manager.save("densita_lookup", pd.DataFrame({"partita": ["1"], "densita": [0.4]}), None)

    assert source_manager.is_current("densita_lookup")


# ---------------------------------------------------------------- price lookup follows the Listini content

def _listini(price):
    return pd.DataFrame([{"CLARTICOLO": "C010003S", "DESCRIZARTICOLOLI": "a", "CLCOLORE": "00101", "CLDESCR": "x",
                          "LIVELLOLPZ": 1, "PREZZOLPZ": price, "CATEGORY": "MED-Cottone"}])


def test_a_new_listini_with_the_same_shape_replaces_the_cached_price_lookup(sqlite_db, monkeypatch):
    from exporters import biglietti_exporter as exporter
    from utility import prezzi_cache, source_manager

    monkeypatch.setattr(exporter, "_PREZZO_LOOKUP_CACHE", None)
    monkeypatch.setattr(prezzi_cache, "load_prezzi_cache", lambda: {})
    source_manager.save("listini", _listini(4.0), None)
    first, _ = exporter.load_prezzo_lookup()
    assert exporter._prezzo_match_for("C010003S", "00101", first)[0] == 4.0

    source_manager.save("listini", _listini(5.5), None)          # same rows, same columns, new price
    second, _ = exporter.load_prezzo_lookup()

    assert exporter._prezzo_match_for("C010003S", "00101", second)[0] == 5.5


def test_the_price_lookup_is_not_rebuilt_while_the_listini_is_unchanged(sqlite_db, monkeypatch):
    from exporters import biglietti_exporter as exporter
    from utility import prezzi_cache, source_manager

    monkeypatch.setattr(exporter, "_PREZZO_LOOKUP_CACHE", None)
    monkeypatch.setattr(prezzi_cache, "load_prezzi_cache", lambda: {})
    source_manager.save("listini", _listini(4.0), None)
    first, _ = exporter.load_prezzo_lookup()
    reads = []
    original = source_manager.load
    monkeypatch.setattr(source_manager, "load", lambda name: (reads.append(name), original(name))[1])

    second, _ = exporter.load_prezzo_lookup()

    assert second is first and reads == []                       # not even the snapshot rows were read again


# ---------------------------------------------------------------- Prezzi tab, data restored from SQLite

def _pump(root, condition, timeout=10.0):
    end = time.time() + timeout
    while time.time() < end:
        root.update()
        if condition():
            return True
        time.sleep(0.02)
    return False


def test_re_applying_the_category_rules_clears_stale_review_flags():
    from calculate import prezzi

    frame = pd.DataFrame([{"CLARTICOLO": "C010999S", "CLCOLORE": "00101", "CATEGORY": "MED-Cottone",
                           "PREZZOLPZ": 4.0, "LIVELLOLPZ": 1, "_CATEGORY_REVIEW": True}])   # flagged at upload time

    result = prezzi.enrich_categories(frame)

    assert "_CATEGORY_REVIEW" not in result.columns or not result["_CATEGORY_REVIEW"].any()
    assert not [i for i in prezzi.validate_price_data(result) if i["key"].startswith("prezzi-category-review")]


def test_assigning_a_category_works_after_a_restart_when_listini_came_from_sqlite(sqlite_db, monkeypatch):
    tk = pytest.importorskip("tkinter")
    from calculate import prezzi
    from gui.tabs.prezzi_tab import PrezziTab
    from utility import notifications, source_manager

    try:
        root = tk.Tk()
    except tk.TclError:
        pytest.skip("no display available")
    root.withdraw()
    try:
        monkeypatch.setattr(prezzi, "_CATEGORY_OVERRIDES_PATH", sqlite_db / "overrides.json")
        monkeypatch.setattr(notifications, "_PATH", sqlite_db / "notifications.json")
        monkeypatch.setattr(prezzi, "_PREZZI_CACHE", {})
        rows = []
        for article, color, price in (("C010003S", "00101", 4.0), ("C010003S", "00102", 4.5), ("C010032S", "00101", 9.0),
                                      ("C010999S", "00101", 4.0), ("C010999S", "00777", 6.0)):
            rows.append({"CLARTICOLO": article, "DESCRIZARTICOLOLI": "art", "CLCOLORE": color, "CLDESCR": "c",
                         "LIVELLOLPZ": 1, "PREZZOLPZ": price})
        source_manager.save("listini", prezzi.enrich_categories(pd.DataFrame(rows)), None)
        issues_seen = []
        tab = PrezziTab(root, on_notifications=lambda issues: issues_seen.append(issues))
        tab._restore_from_cache()
        assert tab._loaded_source_path == "sqlite://listini"
        assert "prezzi-category-review:C010999S" in {i["key"] for batch in issues_seen for i in batch}
        notifications.add("prezzi-category-review:C010999S", "Review", "needs a category", "Prezzi")

        prezzi.save_category_override("C010999S", "MED-Cottone")
        done = []
        tab.reapply_categories(on_done=lambda: done.append(1))

        assert _pump(root, lambda: done)                          # it used to try to open the file "sqlite://listini"
        categories = tab.prezzi_df.loc[tab.prezzi_df["CLARTICOLO"] == "C010999S", "CATEGORY"]
        assert categories.eq("MED-Cottone").all()
        assert "prezzi-category-review:C010999S" not in {i["key"] for i in notifications.list_open()}
    finally:
        root.destroy()
