# Notion Articles Sync — Unit Tests

These tests pin down the upsert behavior of the sync module. They are
self-contained: every test stubs `fetch_all_items` and `api_call` so no Notion
network access is required.

## Running

```bash
cd unit_tests
python -m pytest -v
```

## Test files

### test01_rate_limiter.py
Token-bucket `RateLimiter`: basic acquire, timeout behavior, thread safety,
token regeneration.

### test02_api_calls.py
`api_call` retry/backoff and concurrent invocation. Validates the call
counter is thread-safe.

### test03_incremental_sync.py
The upsert loop end-to-end (with mocked Notion). Confirms that a
`last_edited_time`-bumped candidate whose mapped fields match the target is
**not** re-synced, that real field changes are enqueued, that missing targets
become creates, and that excluded items short-circuit before any target lookup.

### test04_value_equivalence.py
The `_values_equivalent` and `items_are_different` helpers:
- ISO date `Z` vs `+00:00` round-trip
- empty value variants (`None`, `""`, `[]`)
- `select` ↔ `multi_select` coercion (`"x"` ≡ `["x"]`)
- `multi_select` order independence
- explicit guarantee that `last_edited_time` is ignored

### test05_source_tracking_idempotent.py
`update_source_tracking` skips its PATCH when the source already tracks the
given archive id. Without this guard the sync feeds itself: every update
bumps `source.last_edited_time`, the next incremental run re-fetches the
same row, the comparator sees no change, but the *previous* version of the
code wrote anyway — perpetuating the loop.

## Why the previous test suite was wrong

The earlier suite (and a "solution demonstration" doc) framed the recurring
sync slowness as a *progress reporting* / UX problem. It was not. The actual
defect was that incremental sync trusted `last_edited_time` as a change
signal. Notion bumps that timestamp for many reasons unrelated to mapped
field content (schema edits, formula recomputes, our own write-back), so
once the database was touched DB-wide the sync flagged ~100% of rows as
"changed" on every run. The progress logs just made the symptom visible.

The tests in this directory now exercise the real fix: a per-field semantic
comparison plus an idempotent source write-back.
