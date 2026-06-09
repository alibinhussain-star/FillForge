# Consumer Electronics Module — 18 New SubTypes + Per-SubType Titles

## Original Problem Statement
Expand the CE Module to support 18 new Product SubTypes plus apply the bespoke
title / internalTitle formulas defined in the `CE - Mapping Logic` sheet.

## Architecture
- **Backend:** Flask (`app.py`) on :5050. Loads SubTypes dynamically from
  `Consumer_Electronics_Template.xlsx` at startup.
- **Frontend:** SPA in `templates/index.html`. Dropdown loaded via `/ce_subtypes`.

## What's Implemented (Jun 2026)
- ✅ All 20 SubTypes present across Category List / CE - PV Template / CE - Mapping Logic.
- ✅ `_build_ce_header_row_map()` now keys by **Column D = SubType** per spec.
- ✅ `CE_TEMPLATE_PATH` points at the new file.
- ✅ Frontend help-text updated (20 verticals).
- ✅ **Per-SubType title + internalTitle formulas implemented** for all 17 SubTypes
  that have a formula in `CE - Mapping Logic`:
  Feature Phones, Smart Phone, Mobile Adapters & Cables, Hair Trimmer, Speakers,
  Mobile Case & Covers, Earphones, Headsets, Memory Cards, Mobile Cables, Mobile Holders,
  Screen Guards / Protectors, Neck Bands, Microphone, Power Bank, Smart Watches, TWS Ear Buds.
  Projectors / Soundbars / Webcams have no formula in the spec → generic fallback.
- ✅ Added 20+ new column hints to `CE_DUMP_COL_HINTS` so seller-dump columns
  (Output Voltage, Adapter Connector Type, Speaker Type, Compatible Brand, Material,
  Case Cover Type, Rotation, Screen Guard Type, Coverage, Mic Type, Output Connector Type,
  Number of Output Ports, Port Type, Memory Card Type, Storage Capacity, Speed Class, etc.)
  are auto-detected and fed into the title builders.

## Verification (`test_all_titles.py`)
18 rows (one per SubType) submitted → HTTP 200, grand_filled=18. Sample outputs:
- Smart Phone: `Samsung Galaxy A15 50MP Camera Smart Phone, 6GB+128GB, Blue, (Fresh)`
- Memory Cards: `SanDisk 128GB microSDXC Memory Cards, Class 10`
- Screen Guards: `AmazonBasics Fresh Tempered Glass Full Screen for Apple iPhone 15 Pro`
- Smart Watches: `Noise Pulse 2 Max Fresh 1.85" Display Smart Watches, Jet Black`
- Power Bank: `Mi Fresh 20000mAh Power Bank 2 USB Type-C, Black`

## Files
- `/app/ce_app/app.py`
- `/app/ce_app/templates/index.html`
- `/app/ce_app/Consumer_Electronics_Template.xlsx`
- `/app/ce_app/test_ingest.py`  (basic Smart-Watches+Earphones E2E)
- `/app/ce_app/test_all_titles.py`  (all 18 SubTypes, title verification)
- `/app/ce_module_updated.zip`  (bundle)

## Backlog
- P2: Per-SubType **productDescription** templates (description still uses smartphone-centric copy).
- P2: Surface mapping logic as a read-only UI viewer.
- P3: Unit tests per SubType.
