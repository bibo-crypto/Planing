import pandas as pd

from calculate import prezzi
from parsers import dfm_lookup
from utility import situazione_db


def test_customer_prefix_map_includes_all_supported_customers():
    assert prezzi.customer_for_article("C010123S") == "MED"
    assert prezzi.customer_for_article("C011123S") == "MED"
    assert prezzi.customer_for_article("C130123S") == "ELVY"
    assert prezzi.customer_for_article("C170123S") == "EL KAMAL"
    assert prezzi.customer_for_article("C999123S") == ""


def test_unknown_article_uses_prefix_customer_and_articoli_marca(monkeypatch):
    monkeypatch.setattr(prezzi, "_load_articoli_marca_lookup", lambda: {"C170123S": "Cashmere"})
    frame = pd.DataFrame([{
        "CLARTICOLO": "C170123S",
        "CLCOLORE": "000001",
        "LIVELLOLPZ": 39,
        "PREZZOLPZ": 8.5,
    }])
    result = prezzi.enrich_categories(frame)
    assert result.loc[0, "CATEGORY"] == "EL KAMAL - Cashmere"


def test_dfm_cache_is_sqlite_first_and_survives_restart(tmp_path, monkeypatch):
    monkeypatch.setattr(dfm_lookup, "APP_DATA_DIR", tmp_path)
    monkeypatch.setattr(situazione_db, "DB_PATH", str(tmp_path / "dfm.sqlite3"))
    situazione_db.init_db()
    dfm_lookup._ACTIVE_DFM_CACHE.clear()
    dfm_lookup.save_dfm_cache([{"articolo": "C1301"}], "DFM.xlsx", tmp_path / "DFM.xlsx")
    assert dfm_lookup.load_dfm_cache()["entries"] == [{"articolo": "C1301"}]
    assert not (tmp_path / "settings" / "dfm_color_cache.json").exists()
    dfm_lookup._ACTIVE_DFM_CACHE.clear()
    assert dfm_lookup.load_dfm_cache()["entries"] == [{"articolo": "C1301"}]


def test_frame_cache_uses_row_snapshot_without_duplicate_blob(monkeypatch, tmp_path):
    db = tmp_path / "situation.sqlite3"
    monkeypatch.setattr(situazione_db, "DB_PATH", str(db))
    situazione_db.init_db()
    frame = pd.DataFrame({"articolo": ["A1", "A2"], "value": [1, 2]})
    situazione_db.save_frame_cache("test_source", frame)
    restored = situazione_db.load_frame_cache("test_source")
    assert restored.to_dict("records") == frame.to_dict("records")
    conn = situazione_db.get_conn()
    try:
        assert conn.execute("SELECT COUNT(*) FROM frame_cache WHERE source_name='test_source'").fetchone()[0] == 0
        assert conn.execute("SELECT COUNT(*) FROM source_snapshot_row WHERE source_name='test_source'").fetchone()[0] == 2
    finally:
        conn.close()


def test_source_snapshot_reupload_same_data_is_unchanged_and_different_data_replaces(monkeypatch, tmp_path):
    db_path = tmp_path / "situation.sqlite3"
    monkeypatch.setattr(situazione_db, "DB_PATH", str(db_path))
    situazione_db.init_db()
    first = pd.DataFrame({"article": ["A1", "A2"], "value": [1, 2]})
    same = first.copy()
    changed = pd.DataFrame({"article": ["A1", "A2"], "value": [1, 9]})
    assert situazione_db.save_frame_cache("source", first) is True
    assert situazione_db.save_frame_cache("source", same) is False
    assert situazione_db.save_frame_cache("source", changed) is True
    assert int(situazione_db.load_frame_cache("source").iloc[1]["value"]) == 9


def test_source_snapshot_load_does_not_require_original_file(monkeypatch, tmp_path):
    db_path = tmp_path / "situation.sqlite3"
    monkeypatch.setattr(situazione_db, "DB_PATH", str(db_path))
    situazione_db.init_db()
    source = tmp_path / "uploaded.xlsx"
    source.write_text("placeholder")
    frame = pd.DataFrame({"article": ["A1"], "value": [7]})
    situazione_db.save_frame_cache("source", frame, source_path=source)
    source.unlink()
    restored = situazione_db.load_frame_cache("source")
    assert restored.to_dict("records") == frame.to_dict("records")
