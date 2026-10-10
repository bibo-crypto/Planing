"""Source restoration and upload workflows for the Situazione tab."""

from __future__ import annotations

import os
import threading
from datetime import datetime
from pathlib import Path
from tkinter import messagebox

import parsers.situazione_loaders as data_loaders
import utility.situazione_db as db
from parsers.dfm_lookup import build_dfm_lookup, load_dfm_cache, save_dfm_cache, save_dfm_workbook
from parsers.prod_lookup import load_prod_cache, save_prod_cache
from utility.path_manager import save_source, source_path
from utility.utils import logger

SOURCE_ORDER = ["copertura", "data_prod", "dfm", "wincoint", "uscita", "qualita"]
SOURCE_BUTTON_NAMES = {
    "copertura": "Copertura",
    "data_prod": "Produzione",
    "dfm": "DFM",
    "wincoint": "Wincoint",
    "uscita": "Uscita",
    "qualita": "Qualita",
    "codes": "Articoli",
    "listini": "Listini",
}


class SituationSourcesMixin:
    def _on_upload_data(self):
        """Reload every previously uploaded source from its saved file path."""
        uploads = db.get_all_uploads()
        reloaded = []
        missing = []

        for key in SOURCE_ORDER:
            info = uploads.get(key, {})
            path = info.get("file_path", "")
            if not path:
                missing.append(SOURCE_BUTTON_NAMES[key] + " (not uploaded yet)")
            elif not os.path.isfile(path):
                missing.append(f"{SOURCE_BUTTON_NAMES[key]} ({path})")
            else:
                self._handle_upload(key, path)
                reloaded.append(SOURCE_BUTTON_NAMES[key])

        codes_info = uploads.get("codes", {})
        codes_path = codes_info.get("file_path", "")
        if not codes_path:
            missing.append("Articoli (not uploaded yet)")
        elif not os.path.isfile(codes_path):
            missing.append(f"Articoli ({codes_path})")
        else:
            self._handle_codes_upload("codes", codes_path)
            reloaded.append("Articoli")

        listini_info = uploads.get("listini", {})
        listini_path = listini_info.get("file_path", "")
        if not listini_path:
            missing.append("Listini (not uploaded yet)")
        elif not os.path.isfile(listini_path):
            missing.append(f"Listini ({listini_path})")
        else:
            self._handle_listini_upload("listini", listini_path)
            reloaded.append("Listini")

        summary = []
        if reloaded:
            summary.append("Reloaded: " + ", ".join(reloaded))
        if missing:
            summary.append("Not available:\n- " + "\n- ".join(missing))
        messagebox.showinfo("Upload Data", "\n\n".join(summary) or "No saved file paths found.")
    def _refresh_after_restore(self):
        """Rebuild the Situazione table at startup, but only when every source is available.

        With a source missing, _on_refresh() opens a "Missing files" box -- noise at
        every launch for an operator who simply hasn't uploaded that file yet.
        """
        missing = [key for key in SOURCE_ORDER if key not in self.loaded_frames]
        if missing:
            logger.info("Situazione: startup refresh skipped, no saved data yet for: %s", ", ".join(missing))
            return
        try:
            self._on_refresh(preserve_comment_history=True)
        except Exception as exc:  # noqa: BLE001
            logger.warning("Situazione: SQLite startup refresh failed: %s", exc)

    def _auto_restore_saved_files(self):
        """Restore normalized source data from SQLite first.

        The original Excel files are optional after a successful upload. They
        are only read again when the operator explicitly uploads/replaces a
        source. A source with no snapshot yet (saved by an older version) is
        parsed once from its saved path and its snapshot is written, so the next
        start is SQLite only.
        """
        if getattr(self, "_auto_restore_started", False):
            return
        self._auto_restore_started = True
        uploads = db.get_all_uploads()
        restored = []
        for key in SOURCE_ORDER:
            frame = db.load_frame_cache(key)
            if frame is not None and not frame.empty:
                self.loaded_frames[key] = frame
                restored.append(key)
                if key == "copertura":
                    self._copertura_revision += 1
                if key in self.source_rows:
                    self.source_rows[key].set_status(True, f"✅ {len(frame)} rows (SQLite)")
        codes = db.load_frame_cache("codes")
        if codes is not None and not codes.empty:
            db.save_codes(codes)
            self.codes_row.set_status(True, f"✅ Saved ({len(codes)} codes) (SQLite)")
            restored.append("codes")
        listini = db.load_frame_cache("listini")
        if listini is not None and not listini.empty:
            self.listini_row.set_status(True, f"✅ Saved ({len(listini)} rows) (SQLite)")
            restored.append("listini")
        if restored:
            logger.info("Situazione: restored %s from SQLite", ", ".join(restored))

        # Sources with no snapshot but a saved workbook that still exists: one-time upgrade.
        legacy = {}
        for key in SOURCE_ORDER + ["codes"]:
            if key in restored:
                continue
            path = str(uploads.get(key, {}).get("file_path", ""))
            if path and os.path.isfile(path):
                legacy[key] = path
        if not legacy:
            self._startup_snapshot_current = bool(restored)
            self._refresh_after_restore()
            return

        self._startup_restore_in_progress = True

        def worker():
            loaded, errors = {}, {}
            for key, path in legacy.items():
                try:
                    if key == "codes":
                        df, load_errors = data_loaders.load_codes(path)
                    else:
                        df, load_errors = data_loaders.LOADERS[key][1](path)
                    if not load_errors and df is not None and not df.empty:
                        loaded[key] = df
                    else:
                        errors[key] = "; ".join(load_errors) if load_errors else "file is empty"
                except Exception as exc:  # noqa: BLE001
                    errors[key] = str(exc)

            def apply_result():
                self._startup_restore_in_progress = False
                for key, df in loaded.items():
                    try:
                        db.save_frame_cache(key, df, source_path=legacy[key])
                    except Exception as exc:  # noqa: BLE001
                        logger.warning("Could not cache %s in SQLite: %s", key, exc)
                    if key == "codes":
                        db.save_codes(df)
                        self.codes_row.set_status(True, f"✅ Saved ({len(df)} codes)")
                        continue
                    self.loaded_frames[key] = df
                    if key == "copertura":
                        self._copertura_revision += 1
                    self.source_rows[key].set_status(True, f"✅ {len(df)} rows")
                    if key == "dfm":
                        self._save_shared_dfm(legacy[key])
                if errors:
                    logger.warning("Situazione: legacy startup restore had errors: %s", errors)
                self._refresh_after_restore()

            self.after(0, apply_result)

        threading.Thread(target=worker, daemon=True).start()

    def _restore_saved_copertura(self, uploads):
        """Restore Copertura even when the main Situazione snapshot is current."""
        info = uploads.get("copertura", {}) if isinstance(uploads, dict) else {}
        path = str(info.get("file_path", ""))
        if not path or not os.path.isfile(path) or "copertura" in self.loaded_frames:
            return

        def worker():
            try:
                df, errors = data_loaders.load_schedulato(path)
            except Exception as exc:  # noqa: BLE001
                df, errors = None, [str(exc)]

            def apply_result():
                if errors or df is None or df.empty:
                    logger.warning("Situazione: saved Copertura restore failed: %s", errors)
                    return
                self.loaded_frames["copertura"] = df
                self._copertura_revision += 1
                self.source_rows["copertura"].set_status(True, f"✅ {len(df)} rows")

            self.after(0, apply_result)

        threading.Thread(target=worker, daemon=True).start()
    @staticmethod
    def _saved_snapshot_is_current(uploads):
        """SQLite snapshots are authoritative; source files may be offline/moved."""
        if not db.get_all_states():
            return False
        return bool(db.load_frame_cache("copertura") is not None)
    def sync_shared_dfm(self):
        """DFM is intentionally never restored from disk at startup.

        DFM is a large historical reference and must only enter the process
        after the operator explicitly selects it in the upload control.
        """
        return
    def sync_shared_async(self):
        """Restore shared Excel files without blocking the Tk event loop."""
        # Startup restore owns the source loading pass. Running this second
        # pass at the same time would read DFM/Produzione twice.
        if self._startup_restore_in_progress:
            return
        if self._startup_snapshot_current:
            self._startup_snapshot_current = False
            return
        if self._shared_syncing:
            return
        dfm_path = str(load_dfm_cache().get("source_path", ""))
        prod_path = str(load_prod_cache().get("source_path", ""))
        needs_dfm = bool(dfm_path and os.path.isfile(dfm_path) and self._shared_dfm_path != dfm_path)
        needs_prod = bool(prod_path and os.path.isfile(prod_path) and self._shared_prod_path != prod_path)
        if not (needs_dfm or needs_prod):
            return

        self._shared_syncing = True

        def worker():
            prod_result = None
            # DFM data is restored from its SQLite snapshot; the workbook is upload-only.
            if prod_path and os.path.isfile(prod_path) and self._shared_prod_path != prod_path:
                try:
                    prod_result = data_loaders.load_data_prod(prod_path)
                except Exception as exc:  # noqa: BLE001
                    logger.warning("Could not restore shared Produzione file: %s", exc)

            def apply_result():
                self._shared_syncing = False
                if prod_result and not prod_result[1] and prod_result[0] is not None and not prod_result[0].empty:
                    df = prod_result[0]
                    self.loaded_frames["data_prod"] = df
                    self._shared_prod_path = prod_path
                    self.source_rows["data_prod"].set_status(True, f"✅ {len(df)} rows")
                    db.save_upload("data_prod", Path(prod_path).name, len(df), "ok", f"✅ {len(df)} rows - {Path(prod_path).name}", file_path=prod_path)
                self.after_idle(self.sync_shared_async)

            self.after(0, apply_result)

        threading.Thread(target=worker, daemon=True).start()
    def sync_remaining_shared_sources(self):
        """Load all non-specialized sources saved by another page.

        DFM and Produzione have dedicated parsers/caches and are handled by
        ``sync_shared_async``. The remaining Situation inputs use the same
        centralized path registry, so a file uploaded in Overview or another
        page is parsed here automatically.
        """
        if self._other_shared_syncing:
            return
        source_keys = ("copertura", "wincoint", "uscita", "qualita", "articoli", "listini")
        paths = {}
        for key in source_keys:
            path = source_path(key, existing_only=True)
            if path and self._shared_source_paths.get(key) != str(path):
                paths[key] = path
        if not paths:
            return
        self._other_shared_syncing = True

        def worker():
            results = {}
            for key, path in paths.items():
                try:
                    if key == "articoli":
                        result = data_loaders.load_codes(str(path))
                    elif key == "listini":
                        import calculate.prezzi as prezzi_logic
                        result = prezzi_logic.load_prezzi(str(path))
                    else:
                        result = data_loaders.LOADERS[key][1](str(path))
                    results[key] = (path, result)
                except Exception as exc:  # noqa: BLE001
                    logger.warning("Could not restore shared %s file: %s", key, exc)

            def apply_result():
                self._other_shared_syncing = False
                for key, (path, result) in results.items():
                    df, errors = result
                    if errors or df is None or df.empty:
                        continue
                    self._shared_source_paths[key] = str(path)
                    if key == "articoli":
                        db.save_codes(df)
                        self.codes_row.set_status(True, f"✅ Saved ({len(df)} codes)")
                    elif key == "listini":
                        self.listini_row.set_status(True, f"✅ Saved ({len(df)} rows)")
                        self.refresh_prezzo_densita()
                    else:
                        self.loaded_frames[key] = df
                        if key == "copertura":
                            self._copertura_revision += 1
                        self.source_rows[key].set_status(True, f"✅ {len(df)} rows")
                        db.save_upload(key, path.name, len(df), "ok", f"✅ {len(df)} rows - {path.name}", file_path=str(path))
                self.after_idle(self.sync_remaining_shared_sources)

            self.after(0, apply_result)

        threading.Thread(target=worker, daemon=True).start()
    def _save_shared_dfm(self, path):
        """Persist one uploaded DFM workbook for all supported customer prefixes."""
        try:
            counts = save_dfm_workbook(Path(path), Path(path).name)
            self._shared_dfm_path = str(Path(path))
            if self._on_shared_cache_changed:
                self._on_shared_cache_changed()
            logger.info("DFM imported into SQLite: %s", counts)
        except Exception as exc:  # noqa: BLE001
            logger.warning("Could not update shared DFM reference: %s", exc)

    def sync_shared_prod(self):
        """Load the Produzione file selected in either page from the shared cache."""
        cache = load_prod_cache()
        source_path = Path(str(cache.get("source_path", "")))
        if not source_path.is_file():
            return
        if getattr(self, "_shared_prod_path", "") == str(source_path):
            return

        try:
            df, errors = data_loaders.load_data_prod(str(source_path))
        except Exception as exc:  # noqa: BLE001
            logger.warning("Could not restore shared Produzione file: %s", exc)
            return
        if errors or df is None or df.empty:
            return

        self.loaded_frames["data_prod"] = df
        self._shared_prod_path = str(source_path)
        msg = f"✅ {len(df)} rows - {source_path.name}"
        self.source_rows["data_prod"].set_status(True, f"✅ {len(df)} rows")
        db.save_upload("data_prod", source_path.name, len(df), "ok", msg, file_path=str(source_path))
    def _save_shared_prod(self, path):
        """Update the shared Produzione cache when Situazione is the upload source."""
        try:
            save_prod_cache(Path(path))
            self._shared_prod_path = str(Path(path))
            if self._on_shared_cache_changed:
                self._on_shared_cache_changed()
        except Exception as exc:  # noqa: BLE001
            logger.warning("Could not update shared Produzione reference: %s", exc)
    def _handle_upload(self, key, path, cache_path=None):
        """cache_path, if given, is what gets persisted/shown instead of
        path -- used by the Overview 'import everything' button, which
        loads from a temp extract of the real (master) file."""
        display_path = cache_path or path
        label, loader_fn = data_loaders.LOADERS[key]
        try:
            df, errors = loader_fn(path)
        except Exception as exc:  # noqa: BLE001
            errors = [f"An error occurred while reading the file: {exc}"]
            df = None

        if errors or df is None or df.empty:
            msg = "; ".join(errors) if errors else "The file is empty after filtering"
            self.source_rows[key].set_status(False, f"❌ {msg}")
            db.save_upload(key, os.path.basename(display_path), 0, "error", msg, file_path=str(display_path))
            self.loaded_frames.pop(key, None)
            return

        self.loaded_frames[key] = df
        try:
            changed = db.save_frame_cache(key, df, source_path=display_path)
            if not changed:
                logger.info("Situazione: %s upload is identical to the SQLite snapshot", key)
        except Exception as exc:  # noqa: BLE001
            logger.warning("Could not cache %s in SQLite: %s", key, exc)
        if key == "copertura":
            self._copertura_revision += 1
        msg = f"✅ {len(df)} rows - {os.path.basename(display_path)}"
        self.source_rows[key].set_status(True, f"✅ {len(df)} rows")
        db.save_upload(key, os.path.basename(display_path), len(df), "ok", msg, file_path=str(display_path))
        if key == "dfm":
            self._save_shared_dfm(display_path)
        if key == "data_prod":
            self._save_shared_prod(display_path)
        # Every successful individual upload becomes available to all pages.
        save_source(key, display_path)
        logger.info("Situazione: %s uploaded — %d rows (%s)", key, len(df), os.path.basename(display_path))
    def _handle_codes_upload(self, _key, path, cache_path=None):
        display_path = cache_path or path
        try:
            df, errors = data_loaders.load_codes(path)
        except Exception as exc:  # noqa: BLE001
            errors = [f"An error occurred while reading the file: {exc}"]
            df = None

        if errors or df is None or df.empty:
            msg = "; ".join(errors) if errors else "The file is empty"
            self.codes_row.set_status(False, f"❌ {msg}")
            return

        db.save_codes(df)
        try:
            changed = db.save_frame_cache("codes", df, source_path=display_path)
        except Exception as exc:
            changed = True
            logger.warning("Could not cache Articoli in SQLite: %s", exc)
        save_source("articoli", display_path)
        msg = f"✅ Saved ({len(df)} codes) - {os.path.basename(display_path)}"
        self.codes_row.set_status(True, msg)
        db.save_upload("codes", os.path.basename(display_path), len(df), "ok", msg, file_path=str(display_path))
        if not changed:
            messagebox.showwarning("No changes", "Articoli is identical to the data already stored in SQLite; the existing database data was kept.")
        logger.info("Situazione: yarn codes reference updated — %d codes (%s)", len(df), os.path.basename(display_path))

        # Also feed Biglietti's Marca-based Titolo cache, so uploading
        # Articoli from either tab benefits both -- best-effort, never
        # blocks the TITOLO-based save above if this file's Marca column
        # isn't found.
        try:
            import utility.articoli_cache as articoli_cache
            import exporters.biglietti_exporter as biglietti_exporter
            marca_map, _errors = biglietti_exporter.load_articoli_marca_map(Path(path))
            if marca_map:
                articoli_cache.save_articoli_cache(display_path)
                try:
                    import pandas as pd
                    from utility.source_manager import save as save_source_frame
                    marca_df = pd.DataFrame({"articolo": list(marca_map), "marca": list(marca_map.values())})
                    save_source_frame("articoli_marca", marca_df, display_path)
                except Exception:
                    pass
        except Exception:
            pass
    def _handle_listini_upload(self, _key, path, cache_path=None):
        display_path = cache_path or path
        try:
            import utility.prezzi_cache as prezzi_cache
            import calculate.prezzi as prezzi_logic
            df, errors = prezzi_logic.load_prezzi(path)
        except Exception as exc:  # noqa: BLE001
            errors = [f"An error occurred while reading the file: {exc}"]
            df = None

        if errors or df is None or df.empty:
            msg = "; ".join(errors) if errors else "The file is empty"
            self.listini_row.set_status(False, f"❌ {msg}")
            return

        prezzi_cache.save_prezzi_cache(display_path)
        try:
            changed = db.save_frame_cache("listini", df, source_path=display_path)
        except Exception as exc:
            changed = True
            logger.warning("Could not cache Listini in SQLite: %s", exc)
        save_source("listini", display_path)
        msg = f"✅ Saved ({len(df)} rows) - {os.path.basename(display_path)}"
        self.listini_row.set_status(True, msg)
        db.save_upload("listini", os.path.basename(display_path), len(df), "ok", msg, file_path=str(display_path))
        if not changed:
            messagebox.showwarning("No changes", "Listini is identical to the data already stored in SQLite; the existing database data was kept.")
        logger.info("Situazione: Listini (Prezzi) updated — %d rows (%s)", len(df), os.path.basename(display_path))
        self.refresh_prezzo_densita()
    def _refresh_source_labels_from_db(self):
        uploads = db.get_all_uploads()
        for key, row_widget in self.source_rows.items():
            info = uploads.get(key)
            if info:
                ok = info["status"] == "ok"
                if ok and not info.get("file_path"):
                    row_widget.set_status(
                        False,
                        f"⚠ {info['message']} (select once to save path)",
                    )
                else:
                    count = info.get("row_count", "")
                    status_text = f"✅ {count} rows" if ok else f"❌ {info['message']}"
                    row_widget.set_status(ok, status_text)
        codes_info = uploads.get("codes")
        if codes_info:
            if codes_info["status"] == "ok" and not codes_info.get("file_path"):
                self.codes_row.set_status(False, "⚠ Saved status only (select once to save path)")
            else:
                self.codes_row.set_status(codes_info["status"] == "ok",
                                           f"✅ Saved ({codes_info.get('row_count', '')} codes)")
        listini_info = uploads.get("listini")
        if listini_info:
            if listini_info["status"] == "ok" and not listini_info.get("file_path"):
                self.listini_row.set_status(False, "⚠ Saved status only (select once to save path)")
            else:
                self.listini_row.set_status(listini_info["status"] == "ok",
                                             f"✅ Saved ({listini_info.get('row_count', '')} rows)")
