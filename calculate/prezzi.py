"""Listini loading and price lookup business rules."""
from __future__ import annotations

import json
from functools import lru_cache
from pathlib import Path
import re

import pandas as pd

from utility.utils import clean_text, logger

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


@lru_cache(maxsize=1)
def _load_reference_category_map() -> dict[str, str]:
    reference_path = Path(__file__).resolve().parent.parent / "data" / "prezzi_category_map.json"
    try:
        category_articles = json.loads(reference_path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        logger.error("Prezzi: couldn't load category reference %s: %s", reference_path, exc)
        raise RuntimeError(f"Price category reference file is unavailable: {reference_path}") from exc
    result: dict[str, str] = {}
    for category, articles in category_articles.items():
        for article in articles:
            article_key = clean_text(article).upper()
            if article_key in result and result[article_key] != category:
                raise ValueError(f"Article {article_key} has multiple reference categories.")
            result[article_key] = category
    return result


def _category_customer_map(article_categories: dict[str, str]) -> dict[str, str]:
    customers: dict[str, str] = {}
    for article, category in article_categories.items():
        prefix = article[:4]
        customer = category.split("-", 1)[0].strip().casefold()
        previous = customers.setdefault(prefix, customer)
        if previous != customer:
            raise ValueError(f"Article family {prefix} has multiple category customers.")
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
    article_categories = _load_reference_category_map()
    customers = _category_customer_map(article_categories)
    article_keys = out["CLARTICOLO"].map(lambda value: clean_text(value).upper())
    out["CATEGORY"] = article_keys.map(article_categories).fillna("")
    out["CATEGORY"] = provided_category.where(provided_category.ne(""), out["CATEGORY"])
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
                signatures = group[
                    ["_customer", "_color_key", "_level_key", "_price_key"]
                ].drop_duplicates()
                for signature_customer, color, level, price in signatures.itertuples(index=False, name=None):
                    if not signature_customer or not color or level is None or price is None:
                        continue
                    for category in signature_categories.get((signature_customer, color, level, price), ()):
                        category_counts[category] = category_counts.get(category, 0) + 1
                if category_counts:
                    highest = max(category_counts.values())
                    winners = [category for category, count in category_counts.items() if count == highest]
                    if len(winners) == 1:
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


def _modal_category_references(usable: pd.DataFrame, group_keys: list[str]) -> pd.DataFrame:
    price_counts = (
        usable.groupby(group_keys + ["_price"], sort=False, dropna=False)
        .size().rename("_count").reset_index()
    )
    modal_prices = (
        price_counts.sort_values("_count", ascending=False, kind="mergesort")
        .drop_duplicates(group_keys)
    )
    selected = usable.merge(modal_prices[group_keys + ["_price"]], on=group_keys + ["_price"], how="inner")
    level_counts = (
        selected.groupby(group_keys + ["_level"], sort=False, dropna=False)
        .size().rename("_count").reset_index()
    )
    modal_levels = (
        level_counts.sort_values("_count", ascending=False, kind="mergesort")
        .drop_duplicates(group_keys)
    )
    return modal_prices[group_keys + ["_price"]].merge(
        modal_levels[group_keys + ["_level"]],
        on=group_keys,
        how="inner",
    ).rename(columns={"_price": "_reference_price", "_level": "_reference_level"})


def load_prezzi(path: str | Path) -> tuple[pd.DataFrame | None, list[str]]:
    fingerprint = None
    try:
        source = Path(path).resolve()
        stat = source.stat()
        fingerprint = (
            6, str(source), stat.st_mtime_ns, stat.st_size,
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
        work = df.copy()
        work["CATEGORY"] = work["CATEGORY"].fillna("").astype(str).str.strip()
        work["_price"] = pd.to_numeric(work["PREZZOLPZ"], errors="coerce")
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
            suggestions = pd.DataFrame({
                "CLARTICOLO": missing["CLARTICOLO"],
                "CLCOLORE": missing["CLCOLORE"],
                "CLDESCR": missing["CLDESCR"],
                "CATEGORY": missing["CATEGORY"],
                "old_price": missing["_reference_price"],
                "new_price": "",
                "pct_change": "",
                "changed_on": "Category suggestion",
                "issue": (
                    "Missing category price; suggested "
                    + missing["_reference_price"].map(lambda value: f"{float(value):.2f}")
                    + level_suffix
                ),
            })
            category_rows.append(suggestions)

        level_conflict = (
            checked["_level"].notna()
            & checked["_reference_level"].notna()
            & checked["_level"].ne(checked["_reference_level"])
        )
        conflict_mask = (
            checked["_price"].notna()
            & checked["_price"].ne(0)
            & (
                (checked["_price"] - checked["_reference_price"]).abs().gt(0.01)
                | level_conflict
            )
        )
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
            issues.append({"key": f"prezzi-conflict-price:{articolo}:{colore}", "title": "Conflicting color prices", "message": f"Articolo {articolo}, colore {colore} has {int(row['price_count'])} different prices.", "severity": "high"})
        if row["level_count"] > 1:
            issues.append({"key": f"prezzi-conflict-level:{articolo}:{colore}", "title": "Conflicting color levels", "message": f"Articolo {articolo}, colore {colore} has {int(row['level_count'])} different LVL values.", "severity": "medium"})
    if "CATEGORY" in df.columns:
        categorized = df.copy()
        categorized["CATEGORY"] = categorized["CATEGORY"].fillna("").astype(str).str.strip()
        categorized["_price"] = pd.to_numeric(categorized["PREZZOLPZ"], errors="coerce")
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
            pair_count, common_price, common_level = row[2], row[3], row[4]
            level_text = f", livello {float(common_level):g}" if pd.notna(common_level) else ""
            issues.append({
                "key": f"prezzi-category-conflict:{_format_codice(color)}:{category.casefold()}",
                "title": "Conflicting Category prices",
                "message": (
                    f"Colore {color}, Category {category} has {pair_count} different price/level pairs. "
                    f"Most frequent: {float(common_price):.2f}{level_text}."
                ),
                "severity": "medium",
            })
        article_customer = _category_customer_map(_load_reference_category_map())
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
    category_price_counts: dict[tuple[str, str], dict[object, int]] = {}
    category_level_counts: dict[tuple[str, str, object], dict[object, int]] = {}
    article_category_counts: dict[str, dict[str, int]] = {}
    for values in df[fields].itertuples(index=False, name=None):
        articolo, colore, nivel_value, prezzo_value = values[:4]
        category = clean_text(values[4]) if has_category else ""
        key = (clean_text(articolo).upper(), _format_codice(colore))
        if key[1].isdigit() and len(key[1]) < 5:
            key = (key[0], key[1].zfill(5))
        livello = None if pd.isna(nivel_value) else nivel_value
        prezzo = None if pd.isna(prezzo_value) else prezzo_value
        lookup[key] = (livello, prezzo)
        if category:
            article_counts = article_category_counts.setdefault(key[0], {})
            article_counts[category] = article_counts.get(category, 0) + 1
            try:
                valid_price = pd.notna(prezzo) and float(prezzo) > 0.01
            except (TypeError, ValueError):
                valid_price = False
            if valid_price:
                category_key = (key[1], category.casefold())
                price_counts = category_price_counts.setdefault(category_key, {})
                price_counts[prezzo] = price_counts.get(prezzo, 0) + 1
                level_counts = category_level_counts.setdefault((*category_key, prezzo), {})
                level_key = None if pd.isna(livello) else livello
                level_counts[level_key] = level_counts.get(level_key, 0) + 1
        if key[0][:1] in {"C", "G"}:
            swapped = ("G" if key[0][0] == "C" else "C") + key[0][1:]
            lookup.setdefault((swapped, key[1]), (livello, prezzo))
    for category_key, price_counts in category_price_counts.items():
        color, category = category_key
        most_common_price = max(price_counts, key=price_counts.get)
        level_counts = category_level_counts[(*category_key, most_common_price)]
        most_common_level = max(level_counts, key=level_counts.get)
        lookup[("__CATEGORY__", color, category)] = (most_common_level, most_common_price)
    for article, category_counts in article_category_counts.items():
        lookup[("__ARTICLE_CATEGORY__", article)] = max(category_counts, key=category_counts.get)
    return lookup
