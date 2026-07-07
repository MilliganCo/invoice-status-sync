# Changelog

## 2026-07-07 — repository audit

- Removed obsolete duplicate `invoice_sync_laststable.py` from the default branch.
- Documented `invoice_sync.py` as the canonical current script.
- Confirmed inspected current config uses example paths rather than real local/network paths.
- Marked repository private until sample workbooks, business names and history are sanitized.

## Current version

- Excel invoice/status synchronization with adaptive worksheet/header detection.
- Watch mode with deduplication, stable-file wait and save retries.
- Size normalization and formula preservation for status workbooks.
