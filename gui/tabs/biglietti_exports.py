"""Lazy exports used by the Biglietti page and its shared-orders windows."""

from utility.utils import lazy_call

build_el_kamal_stem = lazy_call("exporters.biglietti_exporter", "build_el_kamal_stem")
build_output_stem = lazy_call("exporters.biglietti_exporter", "build_output_stem")
detect_order_format = lazy_call("exporters.biglietti_exporter", "detect_order_format")
enrich_records = lazy_call("exporters.biglietti_exporter", "enrich_records")
export_word = lazy_call("exporters.biglietti_exporter", "export_word")
export_workbook = lazy_call("exporters.biglietti_exporter", "export_workbook")
load_articoli_marca_lookup = lazy_call(
    "exporters.biglietti_exporter", "load_articoli_marca_lookup"
)
load_articoli_titolo_map = lazy_call(
    "exporters.biglietti_exporter", "load_articoli_titolo_map"
)
load_densita_query = lazy_call("exporters.biglietti_exporter", "load_densita_query")
load_el_kamal_order = lazy_call("exporters.biglietti_exporter", "load_el_kamal_order")
load_order = lazy_call("exporters.biglietti_exporter", "load_order")
load_prezzo_lookup = lazy_call("exporters.biglietti_exporter", "load_prezzo_lookup")
load_vmm22_ratio_from_magazino = lazy_call(
    "exporters.biglietti_exporter", "load_vmm22_ratio_from_magazino"
)
_filato_rows = lazy_call("exporters.biglietti_exporter", "_filato_rows")
append_create_excel = lazy_call("exporters.biglietti_exporter", "append_create_excel")
deduplicate_create_excel = lazy_call(
    "exporters.biglietti_exporter", "deduplicate_create_excel"
)
load_create_excel_records = lazy_call(
    "exporters.biglietti_exporter", "load_create_excel_records"
)
sync_workbook_history = lazy_call(
    "exporters.biglietti_exporter", "sync_workbook_history"
)
save_pg_x_partita = lazy_call("exporters.biglietti_exporter", "save_pg_x_partita")
pg_x_batch_stock_status = lazy_call(
    "exporters.biglietti_exporter", "pg_x_batch_stock_status"
)
remove_uploaded_pgx_rows = lazy_call(
    "exporters.biglietti_exporter", "remove_uploaded_pgx_rows"
)
update_pg_x_row = lazy_call("exporters.biglietti_exporter", "update_pg_x_row")
update_order_row = lazy_call("exporters.biglietti_exporter", "update_order_row")
move_pg_x_to_orders = lazy_call("exporters.biglietti_exporter", "move_pg_x_to_orders")
delete_pg_x_row = lazy_call("exporters.biglietti_exporter", "delete_pg_x_row")
delete_shipped_shared_rows = lazy_call(
    "exporters.biglietti_exporter", "delete_shipped_shared_rows"
)

export_filato_full = lazy_call("pipelines.ordini_elvy", "export_filato_full")
RawYarnMatch = lazy_call("pipelines.ordini_elvy", "RawYarnMatch")
