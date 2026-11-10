# Parking outcome state

How a parking run and each day inside it ended is a closed set of states declared in `parking/booking.py`, not a free-text word each module interprets for itself. The method is in project-scaffolding's `docs/implicit-state-machines.md`; this page records the two machines that follow it here.

## Why it is declared

A poll or the Thursday burst books one slot per day. The poller and the burst orchestrator share the booking code, and the burst runs one worker process per day that reports back as a JSON line. Before the states were declared, five modules each decided what a status word meant: the worker, the orchestrator, the Telegram report, the poller and the sweep line. A worker's word entered the orchestrator unchecked, so a misspelt or new word passed through silently, and the run status was derived inline in the burst while the poller always said `checked`. The sweep line then read any `checked` run with nothing booked as "a slot was free but booking failed", which is the wrong sentence for a booking that was sent and never confirmed.

Both sets are `StrEnum`s whose values are the words already on the wire (a worker's JSONL line, the Telegram text), so nothing stored or sent changes.

## Day states (`DayStatus`, machine `parking-day`)

One day's attempt. `DayOutcome.status` is coerced through `DayStatus` on construction, so a word outside the set raises `ValueError`.

| State | Word | Meaning |
|---|---|---|
| `BOOKED` | `booked` | confirmed ACTIVE on the site |
| `WOULD_BOOK` | `would-book` | dry-run: a slot was free, nothing was sent |
| `FAILED` | `failed` | slots were free but every booking was refused or errored |
| `UNCONFIRMED` | `unconfirmed` | a booking was sent and the site never showed it ACTIVE |
| `NO_SLOT` | `no-slot` | nothing free for the day |
| `COVERED` | `covered` | the day already had an ACTIVE booking (burst worker only) |
| `ERROR` | `error` | the attempt broke: auth, rate limit, server, a worker that died, or a worker word outside this set |
| `UNKNOWN` | `unknown` | no result arrived in time; never counted as booked or as nothing free |

`is_covered` (`BOOKED`, `COVERED`) and `is_broken` (`ERROR`, `UNKNOWN`) are the two groupings the readers need; a reader asks the state instead of spelling the tuple.

## Run states (`RunStatus`, machine `parking-run`)

How a whole poll or burst ended. `RunResult.status` is coerced through `RunStatus` the same way.

| State | Word | Meaning |
|---|---|---|
| `INACTIVE` | `inactive` | polling is off in this time window |
| `BACKOFF` | `backoff` | not due yet |
| `NO_TOKEN` | `no-token` | no login token in the environment |
| `TOKEN_EXPIRED` | `token-expired` | the login token is past its expiry |
| `NOTHING_PENDING` | `nothing-pending` | every released target day already has a booking |
| `NO_SLOTS` | `no-slots` | days are open but none had a free slot |
| `CHECKED` | `checked` | at least one day was tried; the per-day lists say how each ended |
| `ERROR` | `error` | the run broke before or while trying |

## One derive function

`derive_run_status(outcomes)` is the only place that maps day outcomes to a run status:

| Day outcomes | Run status |
|---|---|
| none | `NOTHING_PENDING` |
| any `ERROR` or `UNKNOWN` | `ERROR` |
| all `NO_SLOT` | `NO_SLOTS` |
| anything else | `CHECKED` |

`result_from_outcomes(pending, outcomes)` builds the whole `RunResult` from it (status plus the booked, would-book, failed and unconfirmed day lists and the error text), and both the burst and the poller call it.

## Who reads it

| Reader | What it does with the states |
|---|---|
| `parking/burst_worker.py` | produces day outcomes, writes the final JSONL line, exits non-zero on `ERROR` |
| `parking/burst.py` | `decode_outcome` turns a worker line into an outcome (an unknown word becomes `ERROR` with the word in `error`); builds the run result; decides the quiet period from `is_covered` |
| `parking/poller.py` | builds the run result; `sweep_summary` gives `UNCONFIRMED` and `FAILED` days their own sentences |
| `parking/report.py` | one Telegram message per state: booked, dry-run, unconfirmed, failed, broken |

A new reader calls `derive_run_status` or asks the state; it does not re-spell a word or re-derive the run from the outcomes. A new mode is a new member plus a handler in every reader above.

## Proof

`tests/parking/test_outcome_state.py` holds the table test for `derive_run_status`, the wire-word round trip, the `sweep_summary` sentence per state, and the tests marked `state-machine-proof: parking-day` and `state-machine-proof: parking-run`, which show a word outside either set is refused.
