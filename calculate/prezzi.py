"""Listini loading and price lookup business rules."""
from __future__ import annotations

import json
from functools import lru_cache
from pathlib import Path
import re

import pandas as pd

from utility.utils import APP_DATA_DIR, clean_text, logger

REQUIRED_COLUMNS = [
    "DATAFINEVAL", "DATAINIZIOVAL", "CLARTICOLO", "DESCRIZARTICOLOLI", "CLCOLORE",
    "CLDESCR", "LIVELLOLPZ", "PREZZOLPZ",
]
DISPLAY_COLUMNS = [
    "CLARTICOLO", "DESCRIZARTICOLOLI", "CLCOLORE", "CLDESCR", "CATEGORY", "LIVELLOLPZ", "PREZZOLPZ",
]
HEADERS = {
    "CLARTICOLO": "Articolo",
    "DESCRIZARTICOLOLI": "Descrizione Articolo",
    "CLCOLORE": "Codice Colore",
    "CLDESCR": "Descrizione Colore",
    "CATEGORY": "Category",
    "LIVELLOLPZ": "Livello",
    "PREZZOLPZ": "Prezzo",
}

# Minimum number of matching colours before an article that is missing from the
# category reference is inferred into a category from its price peers.
MIN_CATEGORY_EVIDENCE = 2

# Memoize the parsed result in-process as well as in SQLite.
_PREZZI_CACHE: dict[tuple, pd.DataFrame] = {}

def _format_codice(value) -> str:
    if pd.isna(value):
        return ""
    text = clean_text(value)
    if text.isdigit() and len(text) > 1 and text.startswith("0"):
        return text
    try:
        as_float = float(value)
    except (TypeError, ValueError):
        return clean_text(value)
    if as_float.is_integer():
        return str(int(as_float))
    return str(as_float)


def _price_color_key(value) -> str:
    """Canonical comparison key for numeric color codes with lost zero padding."""
    text = clean_text(value)
    if text.isdigit():
        return text.lstrip("0") or "0"
    decimal_match = re.fullmatch(r"(\d+)[.,]0+", text)
    if decimal_match:
        digits = decimal_match.group(1)
        return digits.lstrip("0") or "0"
    return text.upper()


def _price_color_variants(value) -> tuple[str, ...]:
    """Return exact and common ERP-padded forms without partial-code matches."""
    text = clean_text(value)
    key = _price_color_key(text)
    if key.isdigit():
        return tuple(dict.fromkeys((text, key, key.zfill(5), key.zfill(6))))
    return (text, key) if text != key else (text,)


def _format_category(value) -> str:
    text = clean_text(value)
    if not text:
        return ""
    if "-" in text:
        prefix, body = text.split("-", 1)
        prefix = prefix.upper() if prefix.isupper() or prefix.casefold() == "med" else prefix
        body = re.sub(
            r"[A-Z][A-Z']*",
            lambda match: match.group(0)[:1] + match.group(0)[1:].lower(),
            body,
        )
        return f"{prefix}-{body}"
    return text.title() if text == text.upper() else text


def _norm_header(value) -> str:
    return re.sub(r"[^a-z0-9]", "", str(value or "").casefold())


def _find_column(columns, *aliases):
    normalized = {_norm_header(column): column for column in columns}
    for alias in aliases:
        column = normalized.get(_norm_header(alias))
        if column is not None:
            return column
    return None


_REFERENCE_MAP_PATH = Path(__file__).resolve().parent.parent / "data" / "prezzi_category_map.json"
_CATEGORY_OVERRIDES_PATH = APP_DATA_DIR / "settings" / "prezzi_category_overrides.json"


def reference_map_fingerprint() -> tuple[int, int] | None:
    """(mtime_ns, size) of the category reference file, or None if missing.

    Part of the Listini cache fingerprint: the cached frame has CATEGORY
    already baked in from this file, so editing the file must invalidate it.
    """
    try:
        stat = _REFERENCE_MAP_PATH.stat()
    except OSError:
        return None
    return (stat.st_mtime_ns, stat.st_size)


def _category_overrides_fingerprint() -> tuple[int, int] | None:
    try:
        stat = _CATEGORY_OVERRIDES_PATH.stat()
    except OSError:
        return None
    return (stat.st_mtime_ns, stat.st_size)


@lru_cache(maxsize=1)
def _load_reference_category_map(
    cache_key: tuple[str, tuple[int, int] | None] | None = None,
) -> dict[str, str]:
    """Article -> category from data/prezzi_category_map.json.

    A missing/corrupt file or a duplicated article must never stop Listini
    from loading (prices are far more important than categories): problems
    are logged loudly and the affected data is skipped instead. The current
    path and file fingerprint are part of the cache key so edits take effect
    without restarting the app.
    """
    path = Path(cache_key[0]) if cache_key is not None else _REFERENCE_MAP_PATH
    try:
        category_articles = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        logger.error("Prezzi: couldn't load category reference %s: %s -- continuing without it", path, exc)
        return {}
    result: dict[str, str] = {}
    for category, articles in category_articles.items():
        for article in articles:
            article_key = clean_text(article).upper()
            if article_key in result and result[article_key] != category:
                logger.warning(
                    "Prezzi: article %s is listed under both %r and %r in the category reference; keeping %r",
                    article_key, result[article_key], category, result[article_key],
                )
                continue
            result[article_key] = category
    return result


def _reference_category_map() -> dict[str, str]:
    path = _REFERENCE_MAP_PATH.resolve()
    return _load_reference_category_map((str(path), reference_map_fingerprint()))


def _load_category_overrides() -> dict[str, str]:
    path = _CATEGORY_OVERRIDES_PATH
    if not path.exists():
        return {}
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        logger.error("Prezzi: couldn't load manual category assignments %s: %s", path, exc)
        return {}
    if not isinstance(raw, dict):
        logger.error("Prezzi: manual category assignments in %s must be a JSON object", path)
        return {}
    return {
        clean_text(article).upper(): _format_category(category)
        for article, category in raw.items()
        if clean_text(article) and clean_text(category)
    }


def load_category_overrides() -> dict[str, str]:
    """Return saved manual article-to-category assignments."""
    return _load_category_overrides()


def save_category_override(article: str, category: str) -> None:
    article_key = clean_text(article).upper()
    category_value = _format_category(category)
    if not article_key or not category_value:
        raise ValueError("Both article and category are required.")
    overrides = _load_category_overrides()
    overrides[article_key] = category_value
    path = _CATEGORY_OVERRIDES_PATH
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(f".{path.name}.tmp")
    temporary.write_text(
        json.dumps(overrides, ensure_ascii=False, indent=2, sort_keys=True),
        encoding="utf-8",
    )
    temporary.replace(path)


def remove_category_override(article: str) -> None:
    article_key = clean_text(article).upper()
    if not article_key:
        return
    overrides = _load_category_overrides()
    if article_key not in overrides:
        return
    del overrides[article_key]
    path = _CATEGORY_OVERRIDES_PATH
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(f".{path.name}.tmp")
    temporary.write_text(
        json.dumps(overrides, ensure_ascii=False, indent=2, sort_keys=True),
        encoding="utf-8",
    )
    temporary.replace(path)


def known_categories(df: pd.DataFrame | None = None) -> list[str]:
    categories = set(_reference_category_map().values())
    categories.update(_load_category_overrides().values())
    if df is not None and not df.empty and "CATEGORY" in df.columns:
        categories.update(
            _format_category(value)
            for value in df["CATEGORY"].dropna()
            if clean_text(value)
        )
    return sorted((category for category in categories if category), key=str.casefold)


def override_warning(article: str, category: str) -> str:
    article_key = clean_text(article).upper()
    category_customer = clean_text(category).split("-", 1)[0].casefold()
    expected_customer = _category_customer_map(_reference_category_map()).get(article_key[:4], "")
    if expected_customer and category_customer and category_customer != expected_customer:
        return (
            f"Article family {article_key[:4]} belongs to {expected_customer}, "
            f"but this category belongs to {category_customer}."
        )
    return ""


def category_review_table(df: pd.DataFrame | None) -> pd.DataFrame:
    """Return one review row for each unresolved article in the Listini data."""
    columns = ["articolo", "descrizione", "customer", "colori", "suggestion", "evidence"]
    if df is None or df.empty or "_CATEGORY_REVIEW" not in df.columns:
        return pd.DataFrame(columns=columns)

    review = df[df["_CATEGORY_REVIEW"].fillna(False)].copy()
    if review.empty:
        return pd.DataFrame(columns=columns)
    article_customers = _category_customer_map(_reference_category_map())
    categorized = df[df["CATEGORY"].fillna("").astype(str).str.strip().ne("")].copy()
    suggestions: dict[str, dict[str, int]] = {}
    required = {"CLARTICOLO", "CLCOLORE", "LIVELLOLPZ", "PREZZOLPZ"}
    if required.issubset(df.columns) and not categorized.empty:
        categorized["_article_key"] = categorized["CLARTICOLO"].map(
            lambda value: clean_text(value).upper()
        )
        categorized["_customer_key"] = categorized["_article_key"].map(
            lambda article: article_customers.get(article[:4], "")
        )
        categorized["_color_key"] = categorized["CLCOLORE"].map(_signature_code)
        categorized["_level_key"] = categorized["LIVELLOLPZ"].map(_signature_number)
        categorized["_price_key"] = categorized["PREZZOLPZ"].map(_signature_number)
        peer_categories: dict[tuple, set[str]] = {}
        for customer, color, level, price, category in categorized[
            ["_customer_key", "_color_key", "_level_key", "_price_key", "CATEGORY"]
        ].dropna(subset=["_level_key", "_price_key"]).itertuples(index=False, name=None):
            if customer and color:
                peer_categories.setdefault((customer, color, level, price), set()).add(
                    _format_category(category)
                )

        for article, rows in review.groupby("CLARTICOLO", sort=False):
            article_key = clean_text(article).upper()
            customer = article_customers.get(article_key[:4], "")
            if not customer:
                continue
            article_suggestions: dict[str, int] = {}
            signatures = rows[["CLCOLORE", "LIVELLOLPZ", "PREZZOLPZ"]].drop_duplicates()
            for color, level, price in signatures.itertuples(index=False, name=None):
                level_key = _signature_number(level)
                price_key = _signature_number(price)
                color_key = _signature_code(color)
                if level_key is None or price_key is None or not color_key:
                    continue
                for candidate in peer_categories.get((customer, color_key, level_key, price_key), ()):
                    article_suggestions[candidate] = article_suggestions.get(candidate, 0) + 1
            suggestions[article_key] = article_suggestions

    rows_out = []
    for article, rows in review.groupby("CLARTICOLO", sort=False):
        article_key = clean_text(article).upper()
        candidates = suggestions.get(article_key, {})
        suggestion = next(iter(candidates)) if len(candidates) == 1 else ""
        rows_out.append({
            "articolo": article_key,
            "descrizione": next(
                (clean_text(value) for value in rows.get("DESCRIZARTICOLOLI", pd.Series(dtype=object)) if clean_text(value)),
                "",
            ),
            "customer": article_customers.get(article_key[:4], ""),
            "colori": rows["CLCOLORE"].nunique() if "CLCOLORE" in rows.columns else len(rows),
            "suggestion": suggestion,
            "evidence": candidates.get(suggestion, 0) if suggestion else 0,
        })
    return pd.DataFrame(rows_out, columns=columns).sort_values("articolo").reset_index(drop=True)


def category_for_article(article) -> str:
    """Reference category for an article, whether or not it has any Listini row.

    C<->G twins are treated as the same article, exactly like the direct
    price lookup does (build_price_lookup).
    """
    key = clean_text(article).upper()
    if not key:
        return ""
    mapping = _reference_category_map() | _load_category_overrides()
    category = mapping.get(key, "")
    if not category and key[:1] in {"C", "G"}:
        twin = ("G" if key[0] == "C" else "C") + key[1:]
        category = mapping.get(twin, "")
    return _format_category(category) if category else ""


def _category_customer_map(article_categories: dict[str, str]) -> dict[str, str]:
    customers: dict[str, str] = {}
    ambiguous: set[str] = set()
    for article, category in article_categories.items():
        prefix = article[:4]
        customer = category.split("-", 1)[0].strip().casefold()
        previous = customers.setdefault(prefix, customer)
        if previous != customer:
            ambiguous.add(prefix)
    for prefix in ambiguous:
        logger.warning(
            "Prezzi: article family %s belongs to more than one customer in the category reference; "
            "it is excluded from automatic category inference", prefix,
        )
        customers.pop(prefix, None)
    return customers


def _signature_code(value) -> str:
    code = _format_codice(value)
    return code.zfill(6) if code.isdigit() and len(code) < 6 else code


def _signature_number(value):
    if pd.isna(value):
        return None
    try:
        number = float(str(value).replace(",", ".").strip())
    except (TypeError, ValueError):
        return None
    return int(number) if number.is_integer() else number


def enrich_categories(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    provided_category = (
        out["CATEGORY"].map(clean_text)
        if "CATEGORY" in out.columns
        else pd.Series("", index=out.index, dtype="object")
    )
    article_categories = _reference_category_map()
    customers = _category_customer_map(article_categories)
    article_keys = out["CLARTICOLO"].map(lambda value: clean_text(value).upper())
    out["CATEGORY"] = article_keys.map(article_categories).fillna("")
    out["CATEGORY"] = provided_category.where(provided_category.ne(""), out["CATEGORY"])
    overrides = _load_category_overrides()
    manual_categories = article_keys.map(overrides).fillna("")
    out["CATEGORY"] = manual_categories.where(manual_categories.ne(""), out["CATEGORY"])
    customer_col = _find_column(out.columns, "CUSTOMER", "CLIENTE", "CLCLIENTE", "CODCLIENTE")
    marca_col = _find_column(out.columns, "MARCA", "BRAND", "MARCHIO")
    if customer_col and marca_col:
        customer = out[customer_col].map(clean_text)
        marca = out[marca_col].map(clean_text)
        category = customer + " - " + marca
        category = category.mask(customer.eq(""), marca).mask(marca.eq(""), customer)
        fill_category = out["CATEGORY"].eq("") & category.ne("")
        out.loc[fill_category, "CATEGORY"] = category[fill_category]

    unknown_mask = out["CATEGORY"].eq("") & article_keys.ne("")
    if unknown_mask.any():
        needed = {"CLCOLORE", "LIVELLOLPZ", "PREZZOLPZ"}
        if needed.issubset(out.columns):
            peers = out.loc[~unknown_mask & out["CATEGORY"].ne("")].copy()
            peers["_article"] = article_keys.loc[peers.index]
            peers["_customer"] = peers["_article"].map(
                lambda article: customers.get(article[:4], "")
            )
            peers["_color_key"] = peers["CLCOLORE"].map(_signature_code)
            peers["_level_key"] = peers["LIVELLOLPZ"].map(_signature_number)
            peers["_price_key"] = peers["PREZZOLPZ"].map(_signature_number)
            peers = peers.dropna(subset=["_level_key", "_price_key"])
            peers = peers[peers["_customer"].ne("") & peers["_color_key"].ne("")]
            signature_categories: dict[tuple, set[str]] = {}
            for customer, color, level, price, category in peers[
                ["_customer", "_color_key", "_level_key", "_price_key", "CATEGORY"]
            ].drop_duplicates().itertuples(index=False, name=None):
                signature_categories.setdefault((customer, color, level, price), set()).add(category)

            unresolved = []
            missing = out.loc[unknown_mask].copy()
            missing["_article"] = article_keys.loc[missing.index]
            missing["_customer"] = missing["_article"].map(
                lambda article: customers.get(article[:4], "")
            )
            missing["_color_key"] = missing["CLCOLORE"].map(_signature_code)
            missing["_level_key"] = missing["LIVELLOLPZ"].map(_signature_number)
            missing["_price_key"] = missing["PREZZOLPZ"].map(_signature_number)
            for article, group in missing.groupby("_article", sort=False):
                category_counts: dict[str, int] = {}
                comparable = 0
                signatures = group[
                    ["_customer", "_color_key", "_level_key", "_price_key"]
                ].drop_duplicates()
                for signature_customer, color, level, price in signatures.itertuples(index=False, name=None):
                    if not signature_customer or not color or level is None or price is None:
                        continue
                    comparable += 1
                    for category in signature_categories.get((signature_customer, color, level, price), ()):
                        category_counts[category] = category_counts.get(category, 0) + 1
                if category_counts:
                    highest = max(category_counts.values())
                    winners = [category for category, count in category_counts.items() if count == highest]
                    # One coincidental match is not enough when the article has
                    # several priced colours to compare (identical prices repeat
                    # a lot across categories): require MIN_CATEGORY_EVIDENCE
                    # matching colours, or every colour it has if it has fewer.
                    if len(winners) == 1 and highest >= min(MIN_CATEGORY_EVIDENCE, comparable):
                        out.loc[group.index, "CATEGORY"] = winners[0]
                        continue
                unresolved.append(article)
            if unresolved:
                out["_CATEGORY_REVIEW"] = article_keys.isin(unresolved)
            else:
                out["_CATEGORY_REVIEW"] = False
        else:
            out["_CATEGORY_REVIEW"] = unknown_mask
    out["CATEGORY"] = out["CATEGORY"].map(_format_category)
    return out


def _normalize_prezzi_columns(df: pd.DataFrame) -> pd.DataFrame:
    aliases = {
        "DATAFINEVAL": ("DATAFINEVAL", "Data Fine Val"),
        "DATAINIZIOVAL": ("DATAINIZIOVAL", "Data Inizio Val"),
        "CLARTICOLO": ("CLARTICOLO", "Articolo"),
        "DESCRIZARTICOLOLI": ("DESCRIZARTICOLOLI", "Descrizione Articolo"),
        "CLCOLORE": ("CLCOLORE", "Codice Colore"),
        "CLDESCR": ("CLDESCR", "Descrizione Colore"),
        "CATEGORY": ("CATEGORY", "Categoria"),
        "LIVELLOLPZ": ("LIVELLOLPZ", "Livello"),
        "PREZZOLPZ": ("PREZZOLPZ", "Prezzo"),
    }
    renames = {}
    for target, choices in aliases.items():
        source = _find_column(df.columns, *choices)
        if source is not None and source != target:
            renames[source] = target
    return df.rename(columns=renames)


def _latest_price_rows(df: pd.DataFrame) -> pd.DataFrame:
    """Keep only the current row of each Articolo+Colore.

    load_prezzi() keeps every distinct historical price of a combo, but only
    the newest one is the price in force. Category rules (majority price,
    ties, conflicts) must not count a superseded price as a "member" -- an
    article that was simply repriced would otherwise look like a conflict
    with itself.
    """
    if "_START_DATE" in df.columns:
        df = df.sort_values("_START_DATE", kind="mergesort", na_position="first")
    return df.drop_duplicates(["CLARTICOLO", "CLCOLORE"], keep="last")


def _modal_category_references(usable: pd.DataFrame, group_keys: list[str]) -> pd.DataFrame:
    """Per (colour, category): the price/level most members agree on.

    Same rule as build_price_lookup so the warnings never disagree with the
    price actually offered: the strict majority price wins; when several
    prices tie for most frequent, ``_price_tie`` is True (build_price_lookup
    then offers no category price at all) and the lowest tied price is only
    used as a stable stand-in so results never depend on row order.
    """
    price_counts = (
        usable.groupby(group_keys + ["_price"], sort=False, dropna=False)
        .size().rename("_count").reset_index()
    )
    top = price_counts.groupby(group_keys, sort=False, dropna=False)["_count"].transform("max")
    leaders = price_counts[price_counts["_count"].eq(top)]
    leaders = leaders.assign(
        _price_tie=leaders.groupby(group_keys, sort=False, dropna=False)["_price"].transform("size").gt(1)
    )
    modal_prices = leaders.sort_values("_price", kind="mergesort").drop_duplicates(group_keys)
    selected = usable.merge(modal_prices[group_keys + ["_price"]], on=group_keys + ["_price"], how="inner")
    level_counts = (
        selected.groupby(group_keys + ["_level"], sort=False, dropna=False)
        .size().rename("_count").reset_index()
    )
    modal_levels = (
        level_counts.sort_values(["_count", "_level"], ascending=[False, True], kind="mergesort")
        .drop_duplicates(group_keys)
    )
    merged = modal_prices[group_keys + ["_price", "_price_tie"]].merge(
        modal_levels[group_keys + ["_level"]],
        on=group_keys,
        how="inner",
    ).rename(columns={"_price": "_reference_price", "_level": "_reference_level"})
    return merged[[*group_keys, "_reference_price", "_reference_level", "_price_tie"]]


def load_prezzi(path: str | Path) -> tuple[pd.DataFrame | None, list[str]]:
    fingerprint = None
    try:
        source = Path(path).resolve()
        stat = source.stat()
        # 8: include the category reference and manual-override fingerprints;
        # both are baked into CATEGORY in the cached frame.
        fingerprint = (
            8, str(source), stat.st_mtime_ns, stat.st_size,
            reference_map_fingerprint(), _category_overrides_fingerprint(),
        )
    except OSError:
        pass
    if fingerprint is not None:
        cached_df = _PREZZI_CACHE.get(fingerprint)
        if cached_df is not None:
            return cached_df.copy(deep=True), []
        try:
            from utility.prezzi_db import load_prezzi_frame
            cached_df = load_prezzi_frame(fingerprint)
        except Exception:  # noqa: BLE001 -- cache failures must not block Excel loading
            cached_df = None
        if cached_df is not None:
            _PREZZI_CACHE[fingerprint] = cached_df
            return cached_df.copy(deep=True), []

    result = _load_prezzi_uncached(path)
    if fingerprint is not None and result[0] is not None and not result[1]:
        _PREZZI_CACHE[fingerprint] = result[0]
        try:
            from utility.prezzi_db import save_prezzi_frame
            save_prezzi_frame(fingerprint, result[0])
        except Exception as exc:  # noqa: BLE001 -- cache failures must not block the loaded data
            from utility.utils import logger
            logger.warning("Prezzi: could not save SQLite cache: %s", exc)
    return result


def _load_prezzi_uncached(path: str | Path) -> tuple[pd.DataFrame | None, list[str]]:
    try:
        df = _normalize_prezzi_columns(pd.read_excel(path, dtype=object))
    except Exception as exc:  # noqa: BLE001
        return None, [f"Couldn't read the file: {exc}"]

    required = [column for column in REQUIRED_COLUMNS if column not in {"DATAFINEVAL", "DATAINIZIOVAL"}]
    missing = [c for c in required if c not in df.columns]
    if missing:
        return None, [f"Missing columns in the Listini file: {', '.join(missing)}"]

    if "DATAFINEVAL" in df.columns:
        df = df[df["DATAFINEVAL"].isna()]
    df = df[df["CLCOLORE"].notna()]
    start_date = (
        pd.to_datetime(df["DATAINIZIOVAL"], format="%d/%m/%Y", errors="coerce")
        if "DATAINIZIOVAL" in df.columns
        else pd.Series(pd.NaT, index=df.index)
    )
    df = df.assign(_start=start_date).sort_values("_start", na_position="first")
    df = df.drop_duplicates(subset=["CLARTICOLO", "CLCOLORE", "PREZZOLPZ"], keep="last")

    out_data = {
        "CLARTICOLO": df["CLARTICOLO"].map(clean_text),
        "DESCRIZARTICOLOLI": df["DESCRIZARTICOLOLI"].map(clean_text),
        "CLCOLORE": df["CLCOLORE"].map(_format_codice),
        "CLDESCR": df["CLDESCR"].map(clean_text),
        "LIVELLOLPZ": pd.to_numeric(df["LIVELLOLPZ"], errors="coerce"),
        "PREZZOLPZ": pd.to_numeric(df["PREZZOLPZ"], errors="coerce"),
        # Not one of DISPLAY_COLUMNS, so it never shows in the Prezzi tab or
        # its export -- kept only so detect_price_anomalies() below can tell
        # which of a repeated Articolo+Colore's surviving distinct prices
        # came first.
        "_START_DATE": df["_start"].dt.strftime("%Y-%m-%d").fillna(""),
    }
    category_col = _find_column(df.columns, "CATEGORY")
    if category_col is not None:
        out_data["CATEGORY"] = df[category_col].map(clean_text)
    out = pd.DataFrame(out_data)
    customer_col = _find_column(df.columns, "CUSTOMER", "CLIENTE", "CLCLIENTE", "CODCLIENTE")
    marca_col = _find_column(df.columns, "MARCA", "BRAND", "MARCHIO")
    if customer_col:
        out["CUSTOMER"] = df[customer_col].map(clean_text).to_numpy()
    if marca_col:
        out["MARCA"] = df[marca_col].map(clean_text).to_numpy()
    return enrich_categories(out).reset_index(drop=True), []


def detect_price_anomalies(df: pd.DataFrame, min_pct_change: float = 10.0) -> pd.DataFrame:
    """Flag every Articolo+Colore that has more than one surviving distinct
    price (load_prezzi() already dedupes identical prices for the same
    combo, so any group with 2+ rows here is a genuine price change, not a
    duplicate export row) and the jump from the previous price to the next
    is at least ``min_pct_change`` percent in either direction.

    Returns one row per flagged transition, largest change first, so a
    handful of genuine repricings don't get buried under small rounding-size
    changes.
    """
    columns = ["CLARTICOLO", "CLCOLORE", "CLDESCR", "CATEGORY", "old_price", "new_price", "pct_change", "changed_on", "issue"]
    if df is None or df.empty:
        return pd.DataFrame(columns=columns)

    sort_columns = ["CLARTICOLO", "CLCOLORE", "_START_DATE"]
    ordered = df.sort_values(sort_columns, kind="mergesort", na_position="last").copy()
    price_groups = ordered.groupby(["CLARTICOLO", "CLCOLORE"], sort=False)
    ordered["_old_price"] = price_groups["PREZZOLPZ"].shift()
    ordered["_pct_change"] = (
        (ordered["PREZZOLPZ"] - ordered["_old_price"]) / ordered["_old_price"] * 100
    )
    historical_mask = (
        ordered["_old_price"].notna()
        & ordered["PREZZOLPZ"].notna()
        & ordered["_old_price"].ne(0)
        & ordered["_pct_change"].abs().ge(min_pct_change)
    )
    ordered["_group_description"] = price_groups["CLDESCR"].transform("last")
    if "CATEGORY" in ordered:
        ordered["_group_category"] = price_groups["CATEGORY"].transform("last")
    historical = ordered.loc[historical_mask].copy()
    historical_rows = pd.DataFrame({
        "CLARTICOLO": historical["CLARTICOLO"],
        "CLCOLORE": historical["CLCOLORE"],
        "CLDESCR": historical["_group_description"],
        "CATEGORY": historical["_group_category"] if "_group_category" in historical else "",
        "old_price": historical["_old_price"],
        "new_price": historical["PREZZOLPZ"],
        "pct_change": historical["_pct_change"].round(1),
        "changed_on": historical["_START_DATE"].replace("", "(no date)"),
        "issue": "Historical price change",
    })

    category_rows = []
    if "CATEGORY" in df.columns:
        work = _latest_price_rows(df).copy()
        work["CATEGORY"] = work["CATEGORY"].fillna("").astype(str).str.strip()
        work["_price"] = pd.to_numeric(work["PREZZOLPZ"], errors="coerce").round(4)
        work["_level"] = pd.to_numeric(work["LIVELLOLPZ"], errors="coerce")
        group_keys = ["CLCOLORE", "CATEGORY"]
        categorized = work[work["CATEGORY"].ne("")]
        usable = categorized[categorized["_price"].notna() & categorized["_price"].gt(0.01)]
        references = _modal_category_references(usable, group_keys)
        checked = categorized.merge(references, on=group_keys, how="inner", sort=False)

        missing = checked[checked["_price"].isna() | checked["_price"].eq(0)].copy()
        if not missing.empty:
            level_suffix = missing["_reference_level"].map(
                lambda value: f" / level {float(value):g}" if pd.notna(value) else ""
            )
            tied = missing["_price_tie"].astype(bool)
            suggestions = pd.DataFrame({
                "CLARTICOLO": missing["CLARTICOLO"],
                "CLCOLORE": missing["CLCOLORE"],
                "CLDESCR": missing["CLDESCR"],
                "CATEGORY": missing["CATEGORY"],
                # A tied category has no defensible price to suggest.
                "old_price": missing["_reference_price"].where(~tied, ""),
                "new_price": "",
                "pct_change": "",
                "changed_on": "Category suggestion",
                "issue": (
                    "Missing category price; suggested "
                    + missing["_reference_price"].map(lambda value: f"{float(value):.2f}")
                    + level_suffix
                ).where(~tied, "Missing category price; category prices are tied, no price suggested"),
            })
            category_rows.append(suggestions)

        level_conflict = (
            checked["_level"].notna()
            & checked["_reference_level"].notna()
            & checked["_level"].ne(checked["_reference_level"])
        )
        tie_group = checked["_price_tie"].astype(bool)
        conflict_mask = (
            checked["_price"].notna()
            & checked["_price"].ne(0)
            & ~tie_group
            & (
                (checked["_price"] - checked["_reference_price"]).abs().gt(0.01)
                | level_conflict
            )
        )
        # Prices tied for "most frequent": there is no reference to compare
        # against and build_price_lookup offers no category price for the
        # colour, so every priced member of the group is listed.
        ties = checked.loc[
            tie_group & checked["_price"].notna() & checked["_price"].gt(0.01)
        ].drop_duplicates(
            subset=["CLARTICOLO", "CLCOLORE", "CATEGORY", "_price", "_level"]
        )
        if not ties.empty:
            category_rows.append(pd.DataFrame({
                "CLARTICOLO": ties["CLARTICOLO"],
                "CLCOLORE": ties["CLCOLORE"],
                "CLDESCR": ties["CLDESCR"],
                "CATEGORY": ties["CATEGORY"],
                "old_price": "",
                "new_price": ties["_price"],
                "pct_change": "",
                "changed_on": "Category rule",
                "issue": "Category price tie - no category price offered",
            }))
        conflicts = checked.loc[conflict_mask].drop_duplicates(
            subset=["CLARTICOLO", "CLCOLORE", "CATEGORY", "_price", "_level"]
        ).copy()
        if not conflicts.empty:
            conflict_pct = (
                (conflicts["_price"] - conflicts["_reference_price"])
                / conflicts["_reference_price"] * 100
            ).round(1)
            category_rows.append(pd.DataFrame({
                "CLARTICOLO": conflicts["CLARTICOLO"],
                "CLCOLORE": conflicts["CLCOLORE"],
                "CLDESCR": conflicts["CLDESCR"],
                "CATEGORY": conflicts["CATEGORY"],
                "old_price": conflicts["_reference_price"],
                "new_price": conflicts["_price"],
                "pct_change": conflict_pct,
                "changed_on": "Category rule",
                "issue": "Category price/level conflict",
            }))
    result = pd.concat([historical_rows, *category_rows], ignore_index=True)
    result = result.reindex(columns=columns)
    if result.empty:
        return pd.DataFrame(columns=columns)
    return result.sort_values("pct_change", key=lambda s: pd.to_numeric(s, errors="coerce").fillna(0).abs(), ascending=False).reset_index(drop=True)


def validate_price_data(df: pd.DataFrame) -> list[dict[str, str]]:
    """Return actionable data-quality notices for an uploaded Listini file."""
    if df is None or df.empty:
        return [{"key": "prezzi-empty", "title": "Prezzi file is empty", "message": "The active Listini file contains no usable color prices.", "severity": "high"}]
    issues = []

    def _color_fields(value) -> dict[str, str]:
        codice = _format_codice(value)
        descriptions = []
        if "CLDESCR" in df.columns:
            color_codes = df["CLCOLORE"].map(_format_codice)
            descriptions = (
                df.loc[color_codes.eq(codice), "CLDESCR"]
                .map(clean_text)
                .loc[lambda values: values.ne("")]
                .drop_duplicates()
                .tolist()
            )
        return {"codice": codice, "colore": "; ".join(descriptions)}

    numeric_price = pd.to_numeric(df["PREZZOLPZ"], errors="coerce")
    invalid_price = numeric_price.notna() & (numeric_price > 0) & (numeric_price <= 0.01)
    if invalid_price.any():
        issues.append({"key": "prezzi-invalid-price", "title": "Suspicious 0.01 prices in Prezzi", "message": f"{int(invalid_price.sum())} row(s) have price 0.01; blank and zero prices are treated as not priced.", "severity": "high"})
    # A missing/zero price is intentional for colours not priced in Listini.
    counts = df.groupby(["CLARTICOLO", "CLCOLORE"], sort=False).agg(
        price_count=("PREZZOLPZ", "nunique"),
        level_count=("LIVELLOLPZ", "nunique"),
    )
    conflicts = counts[(counts["price_count"] > 1) | (counts["level_count"] > 1)]
    for (articolo, colore), row in conflicts.iterrows():
        if row["price_count"] > 1:
            issues.append({"key": f"prezzi-conflict-price:{articolo}:{colore}", "title": "Conflicting color prices", "message": f"Articolo {articolo}, colore {colore} has {int(row['price_count'])} different prices.", "severity": "high", **_color_fields(colore)})
        if row["level_count"] > 1:
            issues.append({"key": f"prezzi-conflict-level:{articolo}:{colore}", "title": "Conflicting color levels", "message": f"Articolo {articolo}, colore {colore} has {int(row['level_count'])} different LVL values.", "severity": "medium", **_color_fields(colore)})
    if "CATEGORY" in df.columns:
        categorized = _latest_price_rows(df).copy()
        categorized["CATEGORY"] = categorized["CATEGORY"].fillna("").astype(str).str.strip()
        categorized["_price"] = pd.to_numeric(categorized["PREZZOLPZ"], errors="coerce").round(4)
        categorized["_level"] = pd.to_numeric(categorized["LIVELLOLPZ"], errors="coerce")
        categorized = categorized[
            categorized["CATEGORY"].ne("")
            & categorized["_price"].notna()
            & categorized["_price"].gt(0.01)
        ]
        group_keys = ["CLCOLORE", "CATEGORY"]
        pair_counts = (
            categorized.groupby(group_keys + ["_price", "_level"], sort=False, dropna=False)
            .size().rename("_row_count").reset_index()
        )
        group_pair_counts = (
            pair_counts.groupby(group_keys, sort=False).size().rename("_pair_count").reset_index()
        )
        conflicting_groups = group_pair_counts[group_pair_counts["_pair_count"].gt(1)].merge(
            _modal_category_references(categorized, group_keys),
            on=group_keys,
            how="inner",
        )
        for row in conflicting_groups.itertuples(index=False):
            color, category = row[0], row[1]
            pair_count, common_price, common_level, price_tie = row[2], row[3], row[4], bool(row[5])
            level_text = f", livello {float(common_level):g}" if pd.notna(common_level) else ""
            if price_tie:
                # build_price_lookup offers no Category price for a tie, so
                # this is more serious than a plain minority-vs-majority conflict.
                message = (
                    f"Colore {color}, Category {category} has {pair_count} different price/level pairs "
                    f"and no single price is the most frequent. No Category price is offered for "
                    f"this colour until it is resolved."
                )
            else:
                message = (
                    f"Colore {color}, Category {category} has {pair_count} different price/level pairs. "
                    f"Most frequent: {float(common_price):.2f}{level_text}."
                )
            # Distinct key/title for a tie: open notifications are de-duplicated
            # by key, so a minority conflict that later becomes a tie must
            # raise a fresh (high severity) notification instead of staying
            # hidden behind the old one.
            issues.append({
                "key": (
                    f"prezzi-category-{'tie' if price_tie else 'conflict'}:"
                    f"{_format_codice(color)}:{category.casefold()}"
                ),
                "title": "Category prices tied - no price offered" if price_tie else "Conflicting Category prices",
                "message": message,
                "severity": "high" if price_tie else "medium",
                **_color_fields(color),
            })
        article_customer = _category_customer_map(_reference_category_map())
        customer_categories = set(article_customer.values())
        category_customer = categorized["CATEGORY"].map(
            lambda value: value.split("-", 1)[0].strip().casefold()
        )
        article_family = categorized["CLARTICOLO"].map(
            lambda value: clean_text(value).upper()[:4]
        )
        expected_customer = article_family.map(article_customer).fillna("")
        misplaced = categorized[
            expected_customer.ne("")
            & category_customer.isin(customer_categories)
            & expected_customer.ne(category_customer)
        ]
        for (article, color, category), _ in misplaced.groupby(
            ["CLARTICOLO", "CLCOLORE", "CATEGORY"], sort=False
        ):
            expected = article_customer[clean_text(article).upper()[:4]].upper()
            actual = category.split("-", 1)[0].strip().upper()
            issues.append({
                "key": (
                    f"prezzi-category-customer:{clean_text(article).upper()}:"
                    f"{_format_codice(color)}:{category.casefold()}"
                ),
                "title": "Article and Category customer mismatch",
                "message": (
                    f"Articolo {clean_text(article)}, colore {_format_codice(color)} is "
                    f"categorized as {actual}, but article family "
                    f"{clean_text(article).upper()[:4]} belongs to {expected}."
                ),
                "severity": "high",
                **_color_fields(color),
            })
    if "_CATEGORY_REVIEW" in df.columns and "CLARTICOLO" in df.columns:
        review_articles = (
            df.loc[df["_CATEGORY_REVIEW"].fillna(False), "CLARTICOLO"]
            .map(clean_text).drop_duplicates()
        )
        for article in review_articles:
            if article:
                issues.append({
                    "key": f"prezzi-category-review:{article.upper()}",
                    "title": "Category needs review",
                    "message": (
                        f"Articolo {article} has no unambiguous Category match for its client, "
                        "color, level, and price."
                    ),
                    "severity": "medium",
                })
    return issues


def build_price_lookup(df: pd.DataFrame) -> dict[tuple[str, str], tuple]:
    lookup: dict[tuple[str, str], tuple] = {}
    if df is None or df.empty:
        return lookup
    fields = ["CLARTICOLO", "CLCOLORE", "LIVELLOLPZ", "PREZZOLPZ"]
    has_category = "CATEGORY" in df.columns
    if has_category:
        fields.append("CATEGORY")
    has_color_description = "CLDESCR" in df.columns
    if has_color_description:
        fields.append("CLDESCR")
    category_price_counts: dict[tuple[str, str], dict[object, int]] = {}
    category_level_counts: dict[tuple[str, str, object], dict[object, int]] = {}
    article_category_counts: dict[str, dict[str, int]] = {}
    description_prices: dict[tuple[str, str], set[tuple[object, object]]] = {}
    # Rows arrive oldest-first, so the last row of a key is its current price;
    # superseded prices must not vote in the category majority.
    current: dict[tuple[str, str], tuple] = {}
    for values in df[fields].itertuples(index=False, name=None):
        articolo, colore, nivel_value, prezzo_value = values[:4]
        category = clean_text(values[4]) if has_category else ""
        description_value = values[4 + has_category] if has_color_description else None
        description = (
            clean_text(description_value).casefold()
            if description_value is not None and pd.notna(description_value)
            else ""
        )
        key = (clean_text(articolo).upper(), _format_codice(colore))
        if key[1].isdigit() and len(key[1]) < 5:
            key = (key[0], key[1].zfill(5))
        livello = None if pd.isna(nivel_value) else nivel_value
        prezzo = None if pd.isna(prezzo_value) else prezzo_value
        lookup[key] = (livello, prezzo)
        current[key] = (livello, prezzo, category, description)
        if key[0][:1] in {"C", "G"}:
            swapped = ("G" if key[0][0] == "C" else "C") + key[0][1:]
            lookup.setdefault((swapped, key[1]), (livello, prezzo))
    for key, (livello, prezzo, category, description) in current.items():
        if description:
            try:
                valid_description_price = pd.notna(prezzo) and float(prezzo) > 0.01
            except (TypeError, ValueError):
                valid_description_price = False
            if valid_description_price:
                description_key = (key[0], description)
                description_prices.setdefault(description_key, set()).add((livello, prezzo))
                if key[0][:1] in {"C", "G"}:
                    twin = ("G" if key[0][0] == "C" else "C") + key[0][1:]
                    description_prices.setdefault((twin, description), set()).add((livello, prezzo))
        if category:
            article_counts = article_category_counts.setdefault(key[0], {})
            article_counts[category] = article_counts.get(category, 0) + 1
            try:
                valid_price = pd.notna(prezzo) and float(prezzo) > 0.01
            except (TypeError, ValueError):
                valid_price = False
            if valid_price:
                category_key = (_price_color_key(key[1]), category.casefold())
                # Rounded only for counting, so 2.56 and 2.5600000001 (Excel
                # float noise) are the same price and not a fake conflict.
                price_key = round(float(prezzo), 4)
                price_counts = category_price_counts.setdefault(category_key, {})
                price_counts[price_key] = price_counts.get(price_key, 0) + 1
                level_counts = category_level_counts.setdefault((*category_key, price_key), {})
                level_key = None if pd.isna(livello) else livello
                level_counts[level_key] = level_counts.get(level_key, 0) + 1

    for (article, description), prices in description_prices.items():
        if len(prices) == 1:
            lookup[("__ARTICLE_COLOR_DESCRIPTION__", article, description)] = next(iter(prices))

    conflicting: list[tuple[str, str, list]] = []
    dropped_ties = 0
    for category_key, price_counts in category_price_counts.items():
        color, category = category_key
        if len(price_counts) > 1:
            conflicting.append((category, color, sorted(price_counts)))
        # The price most members agree on wins. A tie has no defensible
        # answer, and picking one by row order would silently change the
        # price between two Listini files with identical data -- so no
        # category price is offered for that colour at all.
        top = max(price_counts.values())
        leaders = [price for price, count in price_counts.items() if count == top]
        if len(leaders) > 1:
            dropped_ties += 1
            continue
        most_common_price = leaders[0]
        level_counts = category_level_counts[(*category_key, most_common_price)]
        top_level = max(level_counts.values())
        # Deterministic level on ties: lowest numeric level, missing last.
        most_common_level = sorted(
            (level for level, count in level_counts.items() if count == top_level),
            key=lambda level: (level is None, level if level is not None else 0),
        )[0]
        price_level = (most_common_level, most_common_price)
        for color_variant in _price_color_variants(color):
            lookup[("__CATEGORY__", color_variant, category)] = price_level
    if conflicting:
        examples = "; ".join(
            f"{category} colour {color}: {', '.join(f'{price:g}' for price in prices)}"
            for category, color, prices in conflicting[:5]
        )
        logger.warning(
            "Prezzi: %d category/colour group(s) have members with different prices "
            "(%d tie(s) -> no category price offered). First: %s",
            len(conflicting), dropped_ties, examples,
        )

    for article, category_counts in article_category_counts.items():
        # Highest count first; ties broken by name so the result never
        # depends on row order.
        category = sorted(category_counts, key=lambda name: (-category_counts[name], name))[0]
        lookup[("__ARTICLE_CATEGORY__", article)] = category
        if article[:1] in {"C", "G"}:
            twin = ("G" if article[0] == "C" else "C") + article[1:]
            lookup.setdefault(("__ARTICLE_CATEGORY__", twin), category)
    # Articles with no Listini row at all (never dyed before) still belong to
    # their reference category -- this is the main reason the categories exist.
    for article, category in _reference_category_map().items():
        formatted = _format_category(category)
        lookup.setdefault(("__ARTICLE_CATEGORY__", article), formatted)
        twin = ("G" if article[:1] == "C" else "C") + article[1:] if article[:1] in {"C", "G"} else ""
        if twin:
            lookup.setdefault(("__ARTICLE_CATEGORY__", twin), formatted)
    return lookup
