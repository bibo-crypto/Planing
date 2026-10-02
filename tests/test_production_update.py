import sqlite3

import pandas as pd

from calculate import prezzi
from utility import situazione_db


def test_customer_marca_category_overrides_article_fallback():
    frame = pd.DataFrame({
        "CLARTICOLO": ["C100"],
        "DESCRIZARTICOLOLI": ["fallback"],
        "CUSTOMER": ["Delta"],
        "MARCA": ["Cotone 100%"],
    })
    result = prezzi.enrich_categories(frame)
    assert result.loc[0, "CATEGORY"] == "Delta - Cotone 100%"


def test_uploaded_category_is_preserved_and_display_listini_columns_load(tmp_path):
    path = tmp_path / "category-listini.xlsx"
    pd.DataFrame([{
        "Articolo": "C100",
        "Descrizione Articolo": "Cotone",
        "Codice Colore": "000001",
        "Descrizione Colore": "Blu",
        "Category": "MED-COTTONE",
        "Livello": 39,
        "Prezzo": 2.56,
    }]).to_excel(path, index=False)

    result, errors = prezzi._load_prezzi_uncached(path)

    assert errors == []
    assert result.loc[0, "CATEGORY"] == "MED-Cottone"
    assert result.loc[0, "CLCOLORE"] == "000001"
    assert result.loc[0, "PREZZOLPZ"] == 2.56


def test_reference_categories_keep_hyphen_spacing_and_use_category_case():
    assert prezzi._format_category("MED-COTTONE") == "MED-Cottone"
    assert prezzi._format_category("ELVY-COTTON 85% -15% CASHMERE-30/1") == (
        "ELVY-Cotton 85% -15% Cashmere-30/1"
    )


def test_active_listini_article_uses_reference_category_instead_of_article_marca(tmp_path):
    path = tmp_path / "uncategorized-listini.xlsx"
    pd.DataFrame([
        {
            "DITTA": "060", "LISTINOLPZ": "9001", "ARTICOLOLPZ": "C130025S",
            "LIVELLOLPZ": 60, "DATAINIZIOVAL": "01/01/2026", "DATAFINEVAL": None,
            "CLARTICOLO": "C130025S", "DESCRIZARTICOLOLI": "0070 120.00 2",
            "CLCOLORE": "307795", "CLDESCR": "EL-77951-REATTTIVO", "PREZZOLPZ": 4.59,
        },
        {
            "DITTA": "060", "LISTINOLPZ": "9001", "ARTICOLOLPZ": "C130026S",
            "LIVELLOLPZ": 60, "DATAINIZIOVAL": "01/01/2026", "DATAFINEVAL": None,
            "CLARTICOLO": "C130026S", "DESCRIZARTICOLOLI": "0070 120.00 2",
            "CLCOLORE": "307795", "CLDESCR": "EL-77951-REATTTIVO", "PREZZOLPZ": 4.59,
        },
    ]).to_excel(path, index=False)

    result, errors = prezzi._load_prezzi_uncached(path)

    assert errors == []
    assert result["CATEGORY"].tolist() == ["ELVY-Cottone-Double", "ELVY-Cottone-Double"]


def test_unknown_article_category_uses_same_client_color_level_price_peers():
    frame = pd.DataFrame([
        {"CLARTICOLO": "C130025S", "CLCOLORE": "000001", "LIVELLOLPZ": 39, "PREZZOLPZ": 2.56},
        {"CLARTICOLO": "C130953S", "CLCOLORE": "000001", "LIVELLOLPZ": 39, "PREZZOLPZ": 2.56},
    ])

    result = prezzi.enrich_categories(frame)

    assert result["CATEGORY"].tolist() == ["ELVY-Cottone-Double", "ELVY-Cottone-Double"]
    assert not result.get("_CATEGORY_REVIEW", pd.Series(False, index=result.index)).any()


def test_unknown_article_category_does_not_use_other_customer_peers():
    frame = pd.DataFrame([
        {"CLARTICOLO": "C010003S", "CLCOLORE": "000001", "LIVELLOLPZ": 39, "PREZZOLPZ": 2.56},
        {"CLARTICOLO": "C130953S", "CLCOLORE": "000001", "LIVELLOLPZ": 39, "PREZZOLPZ": 2.56},
    ])

    result = prezzi.enrich_categories(frame)

    assert result.loc[0, "CATEGORY"] == "MED-Cottone"
    assert result.loc[1, "CATEGORY"] == ""
    assert result.loc[1, "_CATEGORY_REVIEW"]


def test_unknown_article_with_conflicting_peer_categories_is_flagged_for_review():
    frame = pd.DataFrame([
        {"CLARTICOLO": "C130025S", "CLCOLORE": "000001", "LIVELLOLPZ": 39, "PREZZOLPZ": 2.56},
        {"CLARTICOLO": "C130154S", "CLCOLORE": "000001", "LIVELLOLPZ": 39, "PREZZOLPZ": 2.56},
        {"CLARTICOLO": "C130953S", "CLCOLORE": "000001", "LIVELLOLPZ": 39, "PREZZOLPZ": 2.56},
    ])

    result = prezzi.enrich_categories(frame)

    assert result.loc[2, "CATEGORY"] == ""
    assert result.loc[2, "_CATEGORY_REVIEW"]
    issue = next(
        item for item in prezzi.validate_price_data(result)
        if item["key"] == "prezzi-category-review:C130953S"
    )
    assert "no unambiguous Category match" in issue["message"]


def test_color_code_lookup_pads_to_five_digits():
    frame = pd.DataFrame({
        "CLARTICOLO": ["C100"], "CLCOLORE": ["123"],
        "LIVELLOLPZ": [1], "PREZZOLPZ": [3.39], "CATEGORY": ["Delta - Cotton"],
    })
    lookup = prezzi.build_price_lookup(frame)
    assert ("C100", "00123") in lookup


def test_prezzi_color_search_prefers_exact_zero_padded_code_over_partial_matches():
    from gui.tabs.prezzi_tab import _filter_color_code

    frame = pd.DataFrame({
        "CLARTICOLO": ["partial", "exact", "other"],
        "CLCOLORE": ["001230", "000123", "123400"],
    })

    result = _filter_color_code(frame, "123")

    assert result["CLARTICOLO"].tolist() == ["exact"]
    assert _filter_color_code(frame, "1234")["CLARTICOLO"].tolist() == ["other"]


def test_category_fallback_uses_most_frequent_price_and_level():
    frame = pd.DataFrame([
        {"CLARTICOLO": "C101", "CLCOLORE": "000001", "LIVELLOLPZ": 39, "PREZZOLPZ": 2.56, "CATEGORY": "MED-COTTONE"},
        {"CLARTICOLO": "C102", "CLCOLORE": "000001", "LIVELLOLPZ": 39, "PREZZOLPZ": 2.56, "CATEGORY": "MED-COTTONE"},
        {"CLARTICOLO": "C103", "CLCOLORE": "000001", "LIVELLOLPZ": 40, "PREZZOLPZ": 2.56, "CATEGORY": "MED-COTTONE"},
        {"CLARTICOLO": "C104", "CLCOLORE": "000001", "LIVELLOLPZ": 40, "PREZZOLPZ": 2.56, "CATEGORY": "MED-COTTONE"},
        {"CLARTICOLO": "C105", "CLCOLORE": "000001", "LIVELLOLPZ": 45, "PREZZOLPZ": 3.04, "CATEGORY": "MED-COTTONE"},
        {"CLARTICOLO": "C106", "CLCOLORE": "000001", "LIVELLOLPZ": 45, "PREZZOLPZ": 3.04, "CATEGORY": "MED-COTTONE"},
        {"CLARTICOLO": "C107", "CLCOLORE": "000001", "LIVELLOLPZ": 45, "PREZZOLPZ": 3.04, "CATEGORY": "MED-COTTONE"},
        {"CLARTICOLO": "CNEW", "CLCOLORE": "000001", "LIVELLOLPZ": None, "PREZZOLPZ": None, "CATEGORY": "MED-COTTONE"},
    ])
    lookup = prezzi.build_price_lookup(frame)

    assert lookup[("__CATEGORY__", "000001", "med-cottone")] == (39, 2.56)
    assert lookup[("CNEW", "000001")][1] is None
    from exporters.biglietti_exporter import prezzo_match_for
    assert prezzo_match_for("CNEW", "000001", lookup) == (2.56, True, "MED-COTTONE", 39)


def test_price_changes_preserve_history_and_category_findings():
    frame = pd.DataFrame([
        {"CLARTICOLO": "A", "CLCOLORE": "1", "CLDESCR": "Blu", "CATEGORY": "Cotone",
         "PREZZOLPZ": 10.0, "LIVELLOLPZ": 1, "_START_DATE": "2025-01-01"},
        {"CLARTICOLO": "A", "CLCOLORE": "1", "CLDESCR": "Blu", "CATEGORY": "Cotone",
         "PREZZOLPZ": 12.0, "LIVELLOLPZ": 1, "_START_DATE": "2025-02-01"},
        {"CLARTICOLO": "B", "CLCOLORE": "1", "CLDESCR": "Blu", "CATEGORY": "Cotone",
         "PREZZOLPZ": 0.0, "LIVELLOLPZ": None, "_START_DATE": "2025-02-01"},
    ])

    result = prezzi.detect_price_anomalies(frame)

    assert len(result) == 3
    assert set(result["changed_on"]) == {"2025-02-01", "Category rule", "Category suggestion"}
    historical = result[result["issue"] == "Historical price change"].iloc[0]
    assert historical["pct_change"] == 20.0
    suggestion = result[result["changed_on"] == "Category suggestion"].iloc[0]
    assert suggestion["issue"] == "Missing category price; suggested 10.00 / level 1"


def test_price_validation_reports_conflicts_and_suspicious_prices():
    frame = pd.DataFrame([
        {"CLARTICOLO": "A", "CLCOLORE": "1", "PREZZOLPZ": 10.0, "LIVELLOLPZ": 1},
        {"CLARTICOLO": "A", "CLCOLORE": "1", "PREZZOLPZ": 12.0, "LIVELLOLPZ": 2},
        {"CLARTICOLO": "A", "CLCOLORE": "2", "PREZZOLPZ": 0.01, "LIVELLOLPZ": 1},
    ])

    issues = prezzi.validate_price_data(frame)

    assert {issue["key"].split(":")[0] for issue in issues} == {
        "prezzi-invalid-price", "prezzi-conflict-price", "prezzi-conflict-level",
    }


def test_price_validation_reports_category_price_conflicts_with_modal_pair():
    frame = pd.DataFrame([
        {"CLARTICOLO": "A", "CLCOLORE": "000001", "PREZZOLPZ": 2.56, "LIVELLOLPZ": 39, "CATEGORY": "MED-COTTONE"},
        {"CLARTICOLO": "B", "CLCOLORE": "000001", "PREZZOLPZ": 2.56, "LIVELLOLPZ": 39, "CATEGORY": "MED-COTTONE"},
        {"CLARTICOLO": "C", "CLCOLORE": "000001", "PREZZOLPZ": 3.04, "LIVELLOLPZ": 45, "CATEGORY": "MED-COTTONE"},
    ])

    issues = prezzi.validate_price_data(frame)

    category_issue = next(issue for issue in issues if issue["key"].startswith("prezzi-category-conflict:"))
    assert "Most frequent: 2.56, livello 39" in category_issue["message"]


def test_price_validation_warns_when_article_family_and_category_customer_disagree():
    frame = pd.DataFrame([
        {"CLARTICOLO": "C130025S", "CLCOLORE": "307795", "PREZZOLPZ": 4.59,
         "LIVELLOLPZ": 60, "CATEGORY": "MED-Cottone"},
        {"CLARTICOLO": "C010003S", "CLCOLORE": "000001", "PREZZOLPZ": 2.56,
         "LIVELLOLPZ": 39, "CATEGORY": "ELVY-Cottone-Double"},
        {"CLARTICOLO": "C130026S", "CLCOLORE": "307796", "PREZZOLPZ": 4.59,
         "LIVELLOLPZ": 60, "CATEGORY": "ELVY-Cottone-Double"},
    ])

    issues = prezzi.validate_price_data(frame)
    mismatches = [issue for issue in issues if issue["key"].startswith("prezzi-category-customer:")]

    assert len(mismatches) == 2
    assert mismatches[0]["severity"] == "high"
    assert "belongs to ELVY" in mismatches[0]["message"]
    assert "belongs to MED" in mismatches[1]["message"]


def test_notifications_add_many_persists_unique_prices_alerts_once(tmp_path):
    from unittest.mock import patch
    from utility import notifications

    path = tmp_path / "notifications.json"
    entries = [
        {"key": "prezzi-category-conflict:000001:med-cottone", "title": "Category conflict",
         "message": "Conflict for color 000001", "page": "Prezzi", "severity": "medium"},
        {"key": "prezzi-category-conflict:000002:med-cottone", "title": "Category conflict",
         "message": "Conflict for color 000002", "page": "Prezzi", "severity": "medium"},
        {"key": "prezzi-category-conflict:000001:med-cottone", "title": "Category conflict",
         "message": "Duplicate", "page": "Prezzi", "severity": "medium"},
    ]
    with patch.object(notifications, "_PATH", path):
        assert notifications.add_many(entries) == 2
        assert notifications.add_many(entries) == 0
        saved = notifications.list_all()

    assert len(saved) == 2
    assert {item["category"] for item in saved} == {"Prices"}


def test_prezzi_sqlite_cache_round_trip_and_fingerprint_invalidation(tmp_path):
    from utility.prezzi_db import load_prezzi_frame, save_prezzi_frame

    fingerprint = (1, "listini.xlsx", 123, 456, "articoli.xlsx", 789, 10)
    frame = pd.DataFrame({"CLARTICOLO": ["C100"], "PREZZOLPZ": [3.25], "CATEGORY": ["Cotone"]})
    db_path = tmp_path / "cache.sqlite3"

    save_prezzi_frame(fingerprint, frame, db_path)

    loaded = load_prezzi_frame(fingerprint, db_path)
    assert loaded is not None
    pd.testing.assert_frame_equal(loaded, frame)
    assert load_prezzi_frame((*fingerprint[:3], "changed.xlsx", *fingerprint[4:]), db_path) is None


def test_normalized_snapshot_round_trip(tmp_path, monkeypatch):
    db_path = tmp_path / "tracker.sqlite3"
    monkeypatch.setattr(situazione_db, "DB_PATH", str(db_path))
    situazione_db.init_db()
    source = pd.DataFrame({"partita": ["P-1"], "bagno": ["B-7"], "peso": [12.5]})
    situazione_db.save_frame_cache("copertura", source)
    loaded = situazione_db.load_frame_cache("copertura")
    assert loaded.to_dict(orient="records") == source.to_dict(orient="records")
    conn = sqlite3.connect(db_path)
    assert conn.execute("SELECT COUNT(*) FROM source_snapshot_row").fetchone()[0] == 1
    conn.close()
