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


def test_master_data_category_review_assignments_persist_and_reapply(tmp_path, monkeypatch):
    reference_map = tmp_path / "category-map.json"
    reference_map.write_text(
        '{"MED-Cottone": ["C010003S"]}',
        encoding="utf-8",
    )
    monkeypatch.setattr(prezzi, "_REFERENCE_MAP_PATH", reference_map)
    monkeypatch.setattr(
        prezzi, "_CATEGORY_OVERRIDES_PATH",
        tmp_path / "settings" / "prezzi_category_overrides.json",
    )
    prezzi._load_reference_category_map.cache_clear()
    try:
        frame = pd.DataFrame([
            {"CLARTICOLO": "C010003S", "DESCRIZARTICOLOLI": "Known",
             "CLCOLORE": "00101", "LIVELLOLPZ": 1, "PREZZOLPZ": 4.0},
            {"CLARTICOLO": "C010999S", "DESCRIZARTICOLOLI": "Review",
             "CLCOLORE": "00101", "LIVELLOLPZ": 1, "PREZZOLPZ": 4.0},
            {"CLARTICOLO": "C010999S", "DESCRIZARTICOLOLI": "Review",
             "CLCOLORE": "00777", "LIVELLOLPZ": 1, "PREZZOLPZ": 6.0},
        ])

        enriched = prezzi.enrich_categories(frame)
        review = prezzi.category_review_table(enriched)
        assert review.to_dict("records") == [{
            "articolo": "C010999S",
            "descrizione": "Review",
            "customer": "med",
            "colori": 2,
            "suggestion": "MED-Cottone",
            "evidence": 1,
        }]

        prezzi.save_category_override("c010999s", "MED-COTTONE")
        assigned = prezzi.enrich_categories(frame)
        assert assigned["CATEGORY"].tolist() == ["MED-Cottone"] * 3
        assert not assigned.get(
            "_CATEGORY_REVIEW", pd.Series(False, index=assigned.index)
        ).any()
        assert prezzi.load_category_overrides() == {"C010999S": "MED-Cottone"}
        assert prezzi.category_for_article("C010999S") == "MED-Cottone"

        prezzi.remove_category_override("C010999S")
        restored = prezzi.enrich_categories(frame)
        assert (restored.loc[restored["CLARTICOLO"] == "C010999S", "CATEGORY"] == "").all()
        assert restored.loc[restored["CLARTICOLO"] == "C010999S", "_CATEGORY_REVIEW"].all()
    finally:
        prezzi._load_reference_category_map.cache_clear()


def test_color_code_lookup_pads_to_five_digits():
    frame = pd.DataFrame({
        "CLARTICOLO": ["C100"], "CLCOLORE": ["123"],
        "LIVELLOLPZ": [1], "PREZZOLPZ": [3.39], "CATEGORY": ["Delta - Cotton"],
    })
    lookup = prezzi.build_price_lookup(frame)
    assert ("C100", "00123") in lookup


def test_price_match_handles_six_digit_listini_color_after_excel_drops_zeros():
    from exporters.biglietti_exporter import prezzo_match_for

    lookup = prezzi.build_price_lookup(pd.DataFrame([{
        "CLARTICOLO": "C010034S", "CLCOLORE": "000141",
        "LIVELLOLPZ": 66, "PREZZOLPZ": 5.39, "CATEGORY": "MED-Lino-23/1",
    }]))

    assert prezzo_match_for("C010034S", "141", lookup) == (5.39, False, "", 66)


def test_new_article_uses_category_price_with_six_digit_color():
    from exporters.biglietti_exporter import prezzo_match_for

    frame = pd.DataFrame([
        {"CLARTICOLO": "C010003S", "CLCOLORE": "000141", "LIVELLOLPZ": 66,
         "PREZZOLPZ": 5.39, "CATEGORY": "MED-Cottone"},
        {"CLARTICOLO": "C010005S", "CLCOLORE": "000141", "LIVELLOLPZ": 66,
         "PREZZOLPZ": 5.39, "CATEGORY": "MED-Cottone"},
    ])
    lookup = prezzi.build_price_lookup(frame)

    assert prezzo_match_for("C010006S", "141", lookup) == (
        5.39, True, "MED-Cottone", 66
    )


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

    # A was simply repriced 10 -> 12: that is history, not a category conflict
    # with itself, and the suggestion for B uses the price now in force (12).
    assert len(result) == 2
    assert set(result["changed_on"]) == {"2025-02-01", "Category suggestion"}
    historical = result[result["issue"] == "Historical price change"].iloc[0]
    assert historical["pct_change"] == 20.0
    suggestion = result[result["changed_on"] == "Category suggestion"].iloc[0]
    assert suggestion["issue"] == "Missing category price; suggested 12.00 / level 1"


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


# ---------------------------------------------------------------------------
# Category logic hardening (review of prezzi category design)
# ---------------------------------------------------------------------------

def _rows(article, items):
    return [
        {"CLARTICOLO": article, "CLCOLORE": c, "LIVELLOLPZ": lv, "PREZZOLPZ": p}
        for c, lv, p in items
    ]


def test_article_without_any_listini_row_gets_category_price_from_reference_map():
    from exporters.biglietti_exporter import prezzo_match_for

    # C010003S / C010005S / C010006S are all MED-COTTONE in the reference map;
    # C010006S has no Listini row at all (never dyed before).
    frame = prezzi.enrich_categories(pd.DataFrame(
        _rows("C010003S", [("00101", 1, 4.0)]) + _rows("C010005S", [("00101", 1, 4.0)])
    ))
    lookup = prezzi.build_price_lookup(frame)

    assert prezzo_match_for("C010006S", "101", lookup)[:3] == (4.0, True, "MED-Cottone")


def test_category_lookup_works_for_g_prefixed_twin_of_mapped_article():
    from exporters.biglietti_exporter import prezzo_match_for

    frame = prezzi.enrich_categories(pd.DataFrame(_rows("C010003S", [("00101", 1, 4.0)])))
    lookup = prezzi.build_price_lookup(frame)

    assert prezzo_match_for("G010006S", "101", lookup)[:3] == (4.0, True, "MED-Cottone")
    assert prezzi.category_for_article("G010006S") == "MED-Cottone"


def test_category_price_does_not_depend_on_row_order_and_ties_offer_no_price():
    rows = _rows("C010003S", [("00101", 1, 4.0)]) + _rows("C010005S", [("00101", 1, 5.0)])
    forward = prezzi.build_price_lookup(prezzi.enrich_categories(pd.DataFrame(rows)))
    backward = prezzi.build_price_lookup(prezzi.enrich_categories(pd.DataFrame(rows[::-1])))

    key = ("__CATEGORY__", "00101", "med-cottone")
    assert key not in forward and key not in backward  # 1 vs 1 -> ambiguous

    rows += _rows("C010006S", [("00101", 1, 5.0)])
    majority = prezzi.build_price_lookup(prezzi.enrich_categories(pd.DataFrame(rows)))
    assert majority[key] == (1, 5.0)


def test_unknown_article_with_a_single_coincidental_colour_match_is_flagged_for_review():
    frame = pd.DataFrame(
        _rows("C010003S", [("00101", 1, 4.0), ("00102", 1, 4.5)])
        + _rows("C010032S", [("00101", 1, 9.0)])
        + _rows("C010999S", [("00101", 1, 4.0), ("00777", 1, 6.0)])  # 1 of 2 colours matches
    )

    result = prezzi.enrich_categories(frame)

    unknown = result[result["CLARTICOLO"] == "C010999S"]
    assert (unknown["CATEGORY"] == "").all()
    assert unknown["_CATEGORY_REVIEW"].all()


def test_prezzi_sqlite_cache_keeps_leading_zero_colour_codes(tmp_path):
    from utility.prezzi_db import load_prezzi_frame, save_prezzi_frame

    fingerprint = (7, "listini.xlsx", 1, 2, None)
    frame = pd.DataFrame({
        "CLARTICOLO": ["C100", "C101"], "CLCOLORE": ["00123", "012345"],
        "LIVELLOLPZ": [1.0, None], "PREZZOLPZ": [3.25, None],
    })
    db_path = tmp_path / "cache.sqlite3"
    save_prezzi_frame(fingerprint, frame, db_path)

    loaded = load_prezzi_frame(fingerprint, db_path)

    assert loaded["CLCOLORE"].tolist() == ["00123", "012345"]
    assert loaded["PREZZOLPZ"].iloc[0] == 3.25 and pd.isna(loaded["PREZZOLPZ"].iloc[1])


def test_listini_cache_fingerprint_changes_with_reference_map(monkeypatch):
    before = prezzi.reference_map_fingerprint()
    assert before is not None
    monkeypatch.setattr(prezzi, "_REFERENCE_MAP_PATH", prezzi._REFERENCE_MAP_PATH.with_name("missing.json"))
    assert prezzi.reference_map_fingerprint() is None


def test_reference_category_map_reloads_when_file_changes(tmp_path, monkeypatch):
    path = tmp_path / "category-map.json"
    path.write_text('{"A-One": ["C1"]}', encoding="utf-8")
    monkeypatch.setattr(prezzi, "_REFERENCE_MAP_PATH", path)
    prezzi._load_reference_category_map.cache_clear()
    try:
        assert prezzi._reference_category_map() == {"C1": "A-One"}

        path.write_text('{"Category-Long": ["C1"]}', encoding="utf-8")

        assert prezzi._reference_category_map() == {"C1": "Category-Long"}
    finally:
        prezzi._load_reference_category_map.cache_clear()


def test_missing_reference_map_does_not_stop_listini_from_loading(tmp_path, monkeypatch):
    monkeypatch.setattr(prezzi, "_REFERENCE_MAP_PATH", tmp_path / "missing.json")
    prezzi._load_reference_category_map.cache_clear()
    try:
        frame = pd.DataFrame(_rows("C010003S", [("00101", 1, 4.0)]))
        result = prezzi.enrich_categories(frame)
        assert result["CATEGORY"].tolist() == [""]
        assert prezzi.build_price_lookup(result)[("C010003S", "00101")] == (1, 4.0)
    finally:
        prezzi._load_reference_category_map.cache_clear()


def test_duplicate_article_in_reference_map_keeps_first_and_does_not_raise(tmp_path, monkeypatch):
    path = tmp_path / "map.json"
    path.write_text('{"A-One": ["C1"], "B-Two": ["C1", "C2"]}', encoding="utf-8")
    monkeypatch.setattr(prezzi, "_REFERENCE_MAP_PATH", path)
    prezzi._load_reference_category_map.cache_clear()
    try:
        assert prezzi._load_reference_category_map() == {"C1": "A-One", "C2": "B-Two"}
    finally:
        prezzi._load_reference_category_map.cache_clear()


def _cat_row(article, price, level=1, color="1"):
    return {"CLARTICOLO": article, "CLCOLORE": color, "CLDESCR": "Blu", "CATEGORY": "Cotone",
            "PREZZOLPZ": price, "LIVELLOLPZ": level, "_START_DATE": "2025-01-01"}


def test_category_minority_price_is_reported_in_notifications_and_price_changes():
    frame = pd.DataFrame([_cat_row("A", 10.0), _cat_row("B", 10.0), _cat_row("C", 12.0)])

    issue = next(i for i in prezzi.validate_price_data(frame) if i["key"].startswith("prezzi-category-conflict:"))
    assert issue["severity"] == "medium" and "Most frequent: 10.00" in issue["message"]

    changes = prezzi.detect_price_anomalies(frame)
    row = changes[changes["issue"] == "Category price/level conflict"].iloc[0]
    assert row["CLARTICOLO"] == "C" and row["old_price"] == 10.0 and row["new_price"] == 12.0


def test_category_price_tie_is_reported_in_notifications_and_price_changes():
    frame = pd.DataFrame([_cat_row("A", 10.0), _cat_row("B", 12.0)])

    issues = prezzi.validate_price_data(frame)
    assert not [i for i in issues if i["key"].startswith("prezzi-category-conflict:")]
    issue = next(i for i in issues if i["key"].startswith("prezzi-category-tie:"))
    assert issue["severity"] == "high"
    assert "No Category price is offered" in issue["message"]

    changes = prezzi.detect_price_anomalies(frame)
    ties = changes[changes["issue"] == "Category price tie - no category price offered"]
    assert sorted(ties["CLARTICOLO"]) == ["A", "B"]
    assert sorted(ties["new_price"]) == [10.0, 12.0]
    # the price lookup agrees with what the warnings say: nothing is offered
    assert ("__CATEGORY__", "00001", "cotone") not in prezzi.build_price_lookup(frame)


def test_tied_category_gives_no_suggestion_for_missing_price():
    frame = pd.DataFrame([_cat_row("A", 10.0), _cat_row("B", 12.0), _cat_row("D", 0.0, level=None)])

    changes = prezzi.detect_price_anomalies(frame)

    missing = changes[changes["changed_on"] == "Category suggestion"].iloc[0]
    assert missing["issue"] == "Missing category price; category prices are tied, no price suggested"
    assert missing["old_price"] == ""


def test_repriced_article_does_not_make_a_tie_in_the_category_lookup():
    older = _cat_row("A", 10.0); older["_START_DATE"] = "2025-01-01"
    newer = _cat_row("A", 12.0); newer["_START_DATE"] = "2025-02-01"
    frame = pd.DataFrame([older, newer, _cat_row("B", 12.0)])

    lookup = prezzi.build_price_lookup(frame)

    assert lookup[("__CATEGORY__", "00001", "cotone")] == (1, 12.0)
    assert not [i for i in prezzi.validate_price_data(frame) if i["key"].startswith("prezzi-category-conflict:")]
