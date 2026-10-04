"""Refresh, enrichment, and persistence workflows for Situazione."""

from __future__ import annotations

import re
import threading
from pathlib import Path
from tkinter import messagebox

import pandas as pd
import utility.situazione_db as db
import calculate.situazione as business_logic
from calculate.abbina_suggestions import build_suggestions
from utility import notifications
from utility.densita_cache import load_densita_cache
from utility.utils import logger
from .situazione_sources import SOURCE_BUTTON_NAMES, SOURCE_ORDER


class SituationRefreshMixin:
    def _on_refresh(self, preserve_comment_history=False):
        if "wincoint" not in self.loaded_frames:
            messagebox.showwarning("Missing data", "Upload the WINCOINT orders file before refreshing.")
            return

        missing = [SOURCE_BUTTON_NAMES[k] for k in SOURCE_ORDER if k not in self.loaded_frames]
        if missing:
            messagebox.showwarning(
                "Missing files",
                "These files have not been uploaded in this session:\n- " + "\n- ".join(missing) +
                "\n\nUpload them once, or use Upload Data after their paths have been saved. "
                "Refresh was cancelled so existing derived data is not cleared."
            )
            return

        # previous New Comment per Partita -- this becomes this round's Old Comment,
        # and the cascade in situazione_logic actively uses it, not just for history
        existing = db.get_all_states()
        old_comments = {p: s["new_comment"] for p, s in existing.items()}

        result_df = business_logic.compute_situation(
            orders_df=self.loaded_frames.get("wincoint"),
            dfm_df=self.loaded_frames.get("dfm"),
            data_prod_df=self.loaded_frames.get("data_prod"),
            copertura_df=self.loaded_frames.get("copertura"),
            uscita_df=self.loaded_frames.get("uscita"),
            qualita_df=self.loaded_frames.get("qualita"),
            codes_map=db.load_codes(),
            old_comments=old_comments,
        )

        # --- safety check: has ANYTHING changed vs what's already stored? ---
        any_change = False
        for _, r in result_df.iterrows():
            prev = existing.get(r["partita"])
            if prev is None or prev["new_comment"] != r["new_comment"]:
                any_change = True
                break

        current_partite = {
            str(value).strip()
            for value in result_df.get("partita", pd.Series(dtype=str)).tolist()
            if str(value).strip()
        }
        saved_partite = {str(value).strip() for value in existing}
        partite_changed = current_partite != saved_partite

        same_data = not any_change and not partite_changed and bool(existing)
        if preserve_comment_history or same_data:
            # The user may intentionally upload the same files again to
            # rebuild the Treeview.  Do not show a warning, and do not treat
            # the refresh as a new comment-history step: keep each row's
            # already stored Old Comm. exactly as it is.
            result_df["old_comment"] = result_df["partita"].map(
                lambda partita: existing.get(partita, {}).get("old_comment", "")
            )
        else:
            proceed = messagebox.askyesno(
                "Confirm refresh",
                f"{len(result_df)} batches will be updated. The current New Comment will move to Old Comment "
                "for batches whose status changed. Continue?"
            )
            if not proceed:
                return

        removed = db.remove_states_not_in(current_partite)
        added, updated, unchanged = db.upsert_states(result_df.to_dict(orient="records"))
        self.summary_lbl.config(text=f"Added: {added}  |  Updated: {updated}  |  Unchanged: {unchanged}")
        logger.info(
            "Situazione: refreshed — added=%d updated=%d unchanged=%d removed=%d",
            added, updated, unchanged, removed,
        )
        self._load_table_from_db()
    def _notify_raw_yarn_available(self) -> None:
        """Fire a notification for every row whose Filato Disponibile
        (raw_yarn_match) column has a real match -- yarn that a PG-X-
        flagged color was waiting for is now available in Magazino Filato.
        Keyed by partita so re-runs update/replace the same notice rather
        than piling up duplicates while the match is still there. Resolve
        notices for partitas that no longer have a valid match."""
        prefix = "situazione-raw-yarn-available:"
        if self.current_df.empty or "raw_yarn_match" not in self.current_df.columns:
            matched = self.current_df.iloc[0:0]
        else:
            matched = self.current_df[
                self.current_df["raw_yarn_match"].astype(str).str.strip() != ""
            ]
        matched_rows = [
            row for _, row in matched.iterrows()
            if str(row.get("partita", "")).strip()
        ]
        current_keys = {
            f"{prefix}{str(row.get('partita', '')).strip()}" for row in matched_rows
        }
        notifications.resolve_missing(prefix, current_keys)
        if not self._on_notification:
            return
        for row in matched_rows:
            partita = str(row.get("partita", "")).strip()
            self._on_notification(
                f"{prefix}{partita}",
                "Filato disponibile per un colore in attesa",
                f"Partita {partita} (Bagno {row.get('bagno', '')}, Articolo {row.get('articolo', '')}): "
                f"filato disponibile in Magazino — {row.get('raw_yarn_match', '')}.",
                "Situazione Generale", "medium",
            )
    def _notify_abbina_pending(self, suggestions: pd.DataFrame | None = None) -> None:
        """Fire one aggregate notification when the "Da abbinare" matching
        window (see _open_abbina) currently has suggestions waiting --
        computed here too so it surfaces without the user having to open
        that window first. Pass a precomputed suggestions frame when the
        caller already has one (e.g. from a background thread); otherwise
        it's computed on the calling thread, so prefer passing one in from
        anywhere already running off the UI thread."""
        if not self._on_notification or self.current_df.empty:
            return
        if suggestions is None:
            try:
                suggestions = build_suggestions(self.current_df, max_extra_percent=0.20)
            except Exception as exc:  # noqa: BLE001
                logger.error("Situazione: Da abbinare check failed: %s", exc)
                return
        if suggestions is None or suggestions.empty:
            notifications.resolve("situazione-abbina-pending")
            return
        groups = suggestions.apply(lambda row: f"{row.get('codice', '')}|{row.get('colore', '')}", axis=1).nunique()
        self._on_notification(
            "situazione-abbina-pending",
            "Colori da abbinare",
            f"{groups} gruppo/i colore ({len(suggestions)} riga/e) in attesa di abbinamento — vedi \"Da abbinare\".",
            "Situazione Generale", "medium",
        )
    def _recompute_raw_yarn_match(self) -> None:
        """Fill "Filato Disponibile" for PG-X rows from Magazino Filato's
        current stock -- best-effort, never blocks: if Magazino hasn't been
        loaded yet this just leaves the column blank."""
        if self.current_df.empty:
            self._notify_raw_yarn_available()
            return
        if "comment" not in self.current_df.columns:
            self.current_df["raw_yarn_match"] = ""
            self._notify_raw_yarn_available()
            return
        magazino_tab = self.magazino_tab
        magazino_summary = getattr(magazino_tab, "magazino_summary", None) if magazino_tab else None
        lotti_summary = getattr(magazino_tab, "lotti_summary", None) if magazino_tab else None
        try:
            self.current_df["raw_yarn_match"] = business_logic.compute_raw_yarn_matches(
                self.current_df, magazino_summary, lotti_summary
            )
        except Exception as exc:  # noqa: BLE001
            logger.error("Situazione: raw yarn auto-match failed: %s", exc)
            return
        self._notify_raw_yarn_available()
    def refresh_raw_yarn_match(self) -> None:
        """Public hook: re-run the Filato Disponibile match against whatever
        Magazino Filato/LOTTI data is loaded right now, and re-render."""
        if self.current_df.empty:
            return
        self._recompute_raw_yarn_match()
        self.current_df = business_logic.compute_delivery_dates(self.current_df)
        self._data_revision += 1
        self._render_tree(self.current_df)
    def refresh_prezzo_densita(self) -> None:
        """Public hook: re-run the Prezzo/Densita lookup (e.g. after the
        Densita' Query workbook is uploaded from the Biglietti tab) and
        re-render, without a full WINCOINT refresh."""
        if self.current_df.empty:
            return
        self._recompute_prezzo_densita()
        self._data_revision += 1
        self._render_tree(self.current_df)
    def refresh_raw_yarn_match_async(self) -> None:
        """Refresh raw-yarn suggestions after the cached UI is visible."""
        if self.current_df.empty or getattr(self, "_raw_match_syncing", False):
            return

        snapshot = self.current_df.copy()
        magazino_tab = self.magazino_tab
        magazino_summary = getattr(magazino_tab, "magazino_summary", None)
        lotti_summary = getattr(magazino_tab, "lotti_summary", None)
        magazino_snapshot = magazino_summary.copy() if isinstance(magazino_summary, pd.DataFrame) else magazino_summary
        lotti_snapshot = lotti_summary.copy() if isinstance(lotti_summary, pd.DataFrame) else lotti_summary
        self._raw_match_syncing = True

        def worker():
            try:
                matches = business_logic.compute_raw_yarn_matches(
                    snapshot, magazino_snapshot, lotti_snapshot
                )
                error = None
            except Exception as exc:  # noqa: BLE001
                matches = [""] * len(snapshot)
                error = exc
            try:
                # Computed here (background thread), not in apply_result, so
                # this can't add latency to the UI thread -- Da abbinare's
                # matching pass is not necessarily cheap.
                abbina_suggestions = build_suggestions(snapshot, max_extra_percent=0.20)
            except Exception:  # noqa: BLE001
                abbina_suggestions = None

            def apply_result():
                self._raw_match_syncing = False
                if error:
                    logger.error("Situazione: async raw yarn match failed: %s", error)
                    return
                if not self.winfo_exists() or len(self.current_df) != len(snapshot):
                    return
                self.current_df["raw_yarn_match"] = matches
                self.current_df = business_logic.compute_delivery_dates(self.current_df)
                self._data_revision += 1
                self._render_tree(self.current_df)
                self._notify_raw_yarn_available()
                self._notify_abbina_pending(abbina_suggestions)

            self.after(0, apply_result)

        threading.Thread(target=worker, daemon=True).start()
    def _recompute_prezzo_densita_for_frame(self, frame: pd.DataFrame) -> tuple[pd.DataFrame, list[tuple]]:
        """Fill Prezzo Listini (by Articolo+Codice, from the same Listini
        used by the Biglietti tab's Prezzi source, plus the $2 machine
        surcharge on 24/32/56) and Densita`(360-390) (by Partita, from the
        Densita' Query workbook uploaded on either the Biglietti or
        Situazione tab) -- best-effort, never blocks: blank if the relevant
        source has not been uploaded yet.

        IMPORTANT: this never touches the "prezzo" column (Prezzo Ord.) --
        that is the price actually entered in Wincoint (read in
        situazione_loaders.load_wincoint_orders and carried straight
        through by compute_situation/the DB). "prezzo_lisini" is only the
        expected value computed here, kept in its own column so the two can
        be compared and shown side by side, and so a mismatch survives being
        looked at rather than silently overwriting the real Wincoint price.

        Returns (result_df, pending_notifications): this runs on a
        background thread when called from _recompute_prezzo_densita_async,
        and self._on_notification ultimately reaches a Tk widget's
        .after() -- calling that from a non-main thread is not something
        Tkinter guarantees, so notifications are only collected here and
        must be fired by the caller on the main thread (see
        _fire_prezzo_notifications)."""
        import exporters.biglietti_exporter as biglietti_exporter
        result = frame.copy()
        pending: list[tuple] = []
        if result.empty:
            result["prezzo_lisini"] = ""
            result["densita"] = ""
            return result, pending
        try:
            price_lookup, _source = biglietti_exporter.load_prezzo_lookup()
        except Exception as exc:  # noqa: BLE001
            logger.warning("Situazione: Prezzo lookup failed: %s", exc)
            price_lookup = {}
        densita_map: dict[int, dict] = {}
        try:
            cache = load_densita_cache()
            path = cache.get("source_path")
            if path and Path(path).is_file():
                densita_map, _errors = biglietti_exporter.load_densita_query(Path(path))
        except Exception as exc:  # noqa: BLE001
            logger.warning("Situazione: Densita' Query lookup failed: %s", exc)

        def _prezzo_row(row):
            base, used_category, category, level = biglietti_exporter.prezzo_match_for(
                row.get("articolo", ""), row.get("codice", ""), price_lookup
            )
            expected = biglietti_exporter.apply_machine_surcharge(base, row.get("mc", ""))
            current = row.get("prezzo", "")
            try:
                curr_clean = str(current or "").replace("$", "").replace("USD", "").replace("usd", "").strip().replace(",", ".")
                current_num = float(curr_clean) if curr_clean else None
            except (TypeError, ValueError):
                current_num = None
            try:
                exp_clean = str(expected or "").replace("$", "").replace("USD", "").replace("usd", "").strip().replace(",", ".")
                expected_num = float(exp_clean) if exp_clean else None
            except (TypeError, ValueError):
                expected_num = None
            article = str(row.get("articolo", "")).strip()
            code = str(row.get("codice", "")).strip()
            colore = str(row.get("colore", "")).strip()
            partita = str(row.get("partita", "")).strip()
            partita_col = partita if partita else (colore if colore else code)
            if partita and colore and partita != colore:
                partita_col = f"{partita} ({colore})"
            if article and code:
                identity = f"{article}:{code}:{partita_col}"
                missing_key = f"situazione-price-missing:{identity}"
                category_key = f"situazione-price-category-fallback:{identity}"
                if used_category and expected_num is not None:
                    level_text = f", livello {float(level):g}" if level is not None else ""
                    pending.append((
                        "add", category_key, "Prezzo Listini da Category",
                        f"Articolo {article}, color code {code}: no direct price; used Category {category} price {expected_num:.2f}{level_text}.",
                        "Situazione Generale", "medium", partita_col,
                    ))
                else:
                    pending.append(("resolve", category_key))
                if expected_num is None:
                    if current_num in (None, 0.0):
                        pending.append(("resolve", missing_key))
                    else:
                        pending.append((
                            "add", missing_key, "Prezzo Listini mancante",
                            f"Articolo {article}, color code {code}: Prezzo Ord. {current_num:.2f} exists but Listini has no matching price.", "Situazione Generale", "medium", partita_col,
                        ))
                elif current_num is not None and abs(current_num - 0.01) < 1e-9:
                    pending.append((
                        "add", f"situazione-price-suspicious:{identity}", "Prezzo Ord. sospetto (0.01)",
                        f"Articolo {article}, color code {code}: Prezzo Ord. è 0.01.", "Situazione Generale", "high", partita_col,
                    ))
                elif current_num is not None and expected_num is not None and abs(current_num - expected_num) > 0.01:
                    pending.append((
                        "add", f"situazione-price-mismatch:{identity}:{current_num}:{expected_num}", "Color price differs from Prezzi",
                        f"Articolo {article}, color code {code}: Prezzo Ord. {current_num:.2f}, expected {expected_num:.2f} from Listini (machine rule included).", "Situazione Generale", "high", partita_col,
                    ))
            return expected

        def _densita_row(row):
            if not densita_map:
                return ""
            # Situazione's "partita" is the colored batch (Partita Col).
            # Densita' Query is keyed by the raw-yarn batch (Partita GG),
            # recorded in the Wincoint comment as PG-<number>-...
            match = re.search(r"\bPG\s*[-:]\s*(\d+)\b", str(row.get("comment", "")), re.IGNORECASE)
            if not match:
                return ""
            key = int(match.group(1))
            return densita_map.get(key, {}).get("densita", "")

        result["prezzo_lisini"] = result.apply(_prezzo_row, axis=1)
        result["densita"] = result.apply(_densita_row, axis=1)
        return result, pending
    def _fire_prezzo_notifications(self, pending: list[tuple]) -> None:
        """Runs on the main thread only (see _recompute_prezzo_densita_for_frame's
        docstring) -- the actual self._on_notification/notifications.resolve calls."""
        for entry in pending:
            if entry[0] == "add" and self._on_notification:
                self._on_notification(*entry[1:])
            elif entry[0] == "resolve":
                notifications.resolve(entry[1])
    def _recompute_prezzo_densita(self) -> None:
        self.current_df, pending = self._recompute_prezzo_densita_for_frame(self.current_df)
        self._fire_prezzo_notifications(pending)
    def _recompute_prezzo_densita_async(self) -> None:
        if self._price_densita_syncing or self.current_df.empty:
            return
        self._price_densita_syncing = True
        snapshot = self.current_df.copy()

        def worker():
            try:
                result, pending = self._recompute_prezzo_densita_for_frame(snapshot)
                error = None
            except Exception as exc:  # noqa: BLE001
                result, pending, error = None, [], exc

            def apply_result():
                self._price_densita_syncing = False
                if error:
                    logger.warning("Situazione: async Prezzo/Densita lookup failed: %s", error)
                    return
                if not self.winfo_exists() or len(self.current_df) != len(result):
                    return
                self.current_df["prezzo_lisini"] = result["prezzo_lisini"].to_numpy()
                self.current_df["densita"] = result["densita"].to_numpy()
                self._data_revision += 1
                self._render_tree(self.current_df)
                self._fire_prezzo_notifications(pending)  # main thread: safe to touch Tk here

            self.after(0, apply_result)

        threading.Thread(target=worker, daemon=True).start()
    def _load_table_from_db(self):
        states = db.get_all_states()
        self.current_df = pd.DataFrame(states.values())
        self._data_revision += 1
        if not self.current_df.empty and "bagno" in self.current_df.columns:
            self.current_df = self.current_df.sort_values(
                by="bagno", ascending=True, key=lambda s: s.astype(str)
            )
            self.sort_state["bagno"] = False  # next click on Bagno heading reverses to Z-A
        self._recompute_raw_yarn_match()
        self.current_df = business_logic.compute_delivery_dates(self.current_df)
        # "prezzo" (Prezzo Ord.) is the real Wincoint price and already
        # comes from the DB row itself -- it must NOT be blanked here.
        # "prezzo_lisini" (the Listini-derived expected price) is the only
        # one that needs a placeholder until the async lookup below fills
        # it in and, critically, runs the mismatch check against "prezzo".
        if "prezzo" not in self.current_df.columns:
            self.current_df["prezzo"] = ""
        self.current_df["prezzo_lisini"] = ""
        self.current_df["densita"] = ""
        self._render_tree(self.current_df)
        # Recompute Prezzo Listini (and fire the price-mismatch notification)
        # after every load from the DB -- not just once at tab startup. This
        # is what used to be missing: a plain Refresh left Prezzo Listini
        # blank and the comparison never ran until the Listini file was
        # re-uploaded or the app restarted.
        self._recompute_prezzo_densita_async()
        for callback in tuple(self._table_loaded_callbacks):
            try:
                callback()
            except Exception as exc:  # noqa: BLE001
                logger.warning("Situazione: could not update dependent view: %s", exc)
    def add_table_loaded_callback(self, callback):
        """Register a callback invoked after the Situazione Treeview reloads."""
        if callback not in self._table_loaded_callbacks:
            self._table_loaded_callbacks.append(callback)
