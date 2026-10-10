"""Integration tests for manual category assignments (Master Data > Category Review).

They cover what the unit tests of calculate/prezzi.py cannot: the caches and
notifications that sit between an assignment and the screens that depend on it.
"""
import time

import pandas as pd
import pytest

from calculate import prezzi
from exporters import biglietti_exporter as exporter


def _listini_rows():
    # C010003S is MED-Cottone in the reference map; C010999S is not in it, and only one
    # of its two coloured prices matches MED-Cottone's -> "needs review", not auto-placed.
    rows = []

    def add(article, color, level, price):
        rows.append({"CLARTICOLO": article, "DESCRIZARTICOLOLI": "art " + article, "CLCOLORE": color,
                     "CLDESCR": "col " + color, "LIVELLOLPZ": level, "PREZZOLPZ": price})

    add("C010003S", "00101", 1, 4.0)
    add("C010003S", "00102", 1, 4.5)
    add("C010032S", "00101", 1, 9.0)
    add("C010999S", "00101", 1, 4.0)
    add("C010999S", "00777", 1, 6.0)
    return rows


@pytest.fixture
def listini(tmp_path):
    path = tmp_path / "Listini.xlsx"
    pd.DataFrame(_listini_rows()).to_excel(path, index=False)
    return path


@pytest.fixture
def category_env(tmp_path, monkeypatch):
    """Everything that persists (overrides, SQLite cache, notifications) goes to a temp folder."""
    from utility import notifications, prezzi_cache, prezzi_db

    monkeypatch.setattr(prezzi, "_CATEGORY_OVERRIDES_PATH", tmp_path / "settings" / "prezzi_category_overrides.json")
    monkeypatch.setattr(prezzi_db, "DB_PATH", tmp_path / "orders.sqlite3")
    monkeypatch.setattr(prezzi, "_PREZZI_CACHE", {})
    monkeypatch.setattr(notifications, "_PATH", tmp_path / "settings" / "notifications.json")
    monkeypatch.setattr(exporter, "_PREZZO_LOOKUP_CACHE", None)
    monkeypatch.setattr(prezzi_cache, "save_prezzi_cache", lambda *a, **k: None)
    prezzi._load_reference_category_map.cache_clear()
    yield tmp_path
    prezzi._load_reference_category_map.cache_clear()


def _point_price_cache_at(monkeypatch, path):
    from utility import prezzi_cache

    monkeypatch.setattr(prezzi_cache, "load_prezzi_cache",
                        lambda: {"source_path": str(path), "source_file": path.name})


# ---------------------------------------------------------------- exporter price lookup

def test_price_lookup_follows_category_assignments_without_restart(category_env, listini, monkeypatch):
    _point_price_cache_at(monkeypatch, listini)

    before, _ = exporter.load_prezzo_lookup()
    assert ("__ARTICLE_CATEGORY__", "C010999S") not in before
    assert exporter._prezzo_match_for("C010999S", "00102", before)[:3] == ("", False, "")   # nothing to inherit yet

    prezzi.save_category_override("C010999S", "MED-Cottone")
    after, _ = exporter.load_prezzo_lookup()
    assert after[("__ARTICLE_CATEGORY__", "C010999S")] == "MED-Cottone"
    # colour 00102 was never priced for C010999S: it now takes MED-Cottone's price
    assert exporter._prezzo_match_for("C010999S", "00102", after)[:3] == (4.5, True, "MED-Cottone")

    prezzi.remove_category_override("C010999S")
    undone, _ = exporter.load_prezzo_lookup()
    assert ("__ARTICLE_CATEGORY__", "C010999S") not in undone


def test_price_lookup_is_still_cached_when_nothing_changed(category_env, listini, monkeypatch):
    _point_price_cache_at(monkeypatch, listini)

    first, _ = exporter.load_prezzo_lookup()
    second, _ = exporter.load_prezzo_lookup()

    assert first is second


def test_category_sources_fingerprint_changes_with_assignments(category_env):
    before = prezzi.category_sources_fingerprint()
    prezzi.save_category_override("C010999S", "MED-Cottone")
    assigned = prezzi.category_sources_fingerprint()
    prezzi.remove_category_override("C010999S")

    assert assigned != before
    assert prezzi.category_sources_fingerprint() != assigned


# ---------------------------------------------------------------- dropdown

def test_category_dropdown_lists_each_category_once(category_env):
    frame = prezzi.enrich_categories(pd.DataFrame(_listini_rows()))    # carries the display form "MED-Cottone"

    names = prezzi.known_categories(frame)

    assert len([n for n in names if n.casefold() == "med-cottone"]) == 1
    assert len({n.casefold() for n in names}) == len(names)
    prezzi.save_category_override("C010999S", "med-cottone")             # a different spelling of an existing one
    assert len([n for n in prezzi.known_categories(frame) if n.casefold() == "med-cottone"]) == 1


# ---------------------------------------------------------------- Prezzi tab

def _pump(root, condition, timeout=15.0):
    end = time.time() + timeout
    while time.time() < end:
        root.update()
        if condition():
            return True
        time.sleep(0.02)
    return False


@pytest.fixture
def tk_root():
    tk = pytest.importorskip("tkinter")
    try:
        root = tk.Tk()
    except tk.TclError:
        pytest.skip("no display available")
    root.withdraw()
    yield root
    root.destroy()


@pytest.fixture(autouse=True)
def _run_tab_threads_inline(monkeypatch):
    """Tk only accepts after() from a worker thread while mainloop() runs; these tests poll with
    update() instead, so the tab's background loads run inline (the callbacks still go through after())."""
    import types

    class InlineThread:
        def __init__(self, target=None, args=(), kwargs=None, daemon=None):
            self._target, self._args, self._kwargs = target, args, kwargs or {}

        def start(self):
            self._target(*self._args, **self._kwargs)

    try:
        from gui.tabs import prezzi_tab
    except Exception:  # noqa: BLE001 -- tkinter missing: the Tk tests skip themselves anyway
        return
    monkeypatch.setattr(prezzi_tab, "threading", types.SimpleNamespace(Thread=InlineThread))


def _open_keys(prefix="prezzi-"):
    from utility import notifications

    return {item["key"] for item in notifications.list_open() if item["key"].startswith(prefix)}


def _load(tab, root, path):
    tab._load_path(str(path), save_cache=False, force=True)
    assert _pump(root, lambda: not tab._uploading)


def test_loading_listini_closes_prezzi_notifications_that_no_longer_apply(category_env, listini, tk_root):
    from gui.tabs.prezzi_tab import PrezziTab
    from utility import notifications

    notifications.add("prezzi-old-problem:X", "Old", "was fixed in the new file", "Prezzi")
    notifications.add("situazione-other", "Other", "belongs to another page", "Situazione")
    tab = PrezziTab(tk_root, on_notifications=lambda issues: None)

    _load(tab, tk_root, listini)

    open_keys = _open_keys("")
    assert "prezzi-old-problem:X" not in open_keys          # not reported any more -> closed
    assert "situazione-other" in open_keys                  # other features are never touched


def test_assigning_a_category_closes_its_review_notification(category_env, listini, tk_root):
    from gui.tabs.prezzi_tab import PrezziTab
    from utility import notifications

    tab = PrezziTab(tk_root, on_notifications=lambda issues: [
        notifications.add(i["key"], i["title"], i["message"], "Prezzi") for i in issues])
    _load(tab, tk_root, listini)
    assert "prezzi-category-review:C010999S" in _open_keys()

    prezzi.save_category_override("C010999S", "MED-Cottone")
    done = []
    tab.reapply_categories(on_done=lambda: done.append(1))
    assert _pump(tk_root, lambda: done)

    assert "prezzi-category-review:C010999S" not in _open_keys()
    assert tab.prezzi_df.loc[tab.prezzi_df["CLARTICOLO"] == "C010999S", "CATEGORY"].eq("MED-Cottone").all()


def test_failed_validation_never_closes_notifications(category_env, listini, tk_root, monkeypatch):
    from gui.tabs.prezzi_tab import PrezziTab
    from utility import notifications

    notifications.add("prezzi-keep-me", "Keep", "must survive a failed check", "Prezzi")
    monkeypatch.setattr(prezzi, "validate_price_data", lambda df: (_ for _ in ()).throw(RuntimeError("boom")))
    tab = PrezziTab(tk_root, on_notifications=lambda issues: None)

    _load(tab, tk_root, listini)

    assert "prezzi-keep-me" in _open_keys()


def test_reapply_waits_for_a_running_load_instead_of_racing_it(category_env, tk_root, monkeypatch):
    from gui.tabs.prezzi_tab import PrezziTab

    tab = PrezziTab(tk_root)
    tab._loaded_source_path = "somewhere.xlsx"
    tab._uploading = True
    started = []
    monkeypatch.setattr(tab, "_load_path", lambda *a, **k: started.append((a, k)))

    tab.reapply_categories()
    _pump(tk_root, lambda: False, timeout=0.5)
    assert started == []                                    # a load is running: don't start a second one

    tab._uploading = False
    assert _pump(tk_root, lambda: started)
    assert started[0][1]["force"] is True


def test_reapply_reports_back_even_when_the_listini_file_is_gone(category_env, listini, tk_root):
    from gui.tabs.prezzi_tab import PrezziTab

    tab = PrezziTab(tk_root)
    _load(tab, tk_root, listini)
    listini.unlink()                                        # moved/deleted since it was loaded
    done = []

    tab.reapply_categories(on_done=lambda: done.append(1))

    assert _pump(tk_root, lambda: done)                     # the Category Review screen is never left waiting
    assert not tab._uploading
