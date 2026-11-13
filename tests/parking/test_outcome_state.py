"""The declared day and run outcome states: closed sets, one derive function, one decode path."""

import json
from datetime import datetime
from zoneinfo import ZoneInfo

import pytest

from tools.parking import burst, poller
from tools.parking.booking import (
    DayOutcome,
    DayStatus,
    RunResult,
    RunStatus,
    derive_run_status,
    result_from_outcomes,
)

NOW = datetime(2026, 9, 20, 12, 0, tzinfo=ZoneInfo("Europe/Madrid"))


def outcomes(*statuses):
    return [DayOutcome(f"2026-09-{21 + i}", status, error="why" if status in ("error", "unknown") else "")
            for i, status in enumerate(statuses)]


# The wire words are the pre-declaration strings: a JSONL line an old worker wrote must still decode.
@pytest.mark.parametrize("word", ["booked", "would-book", "failed", "unconfirmed", "no-slot", "covered",
                                  "error", "unknown"])
def test_every_day_word_an_old_worker_wrote_decodes_to_the_same_outcome(word):
    line = json.loads(json.dumps({"done": True, "status": word, "label": "L", "error": "E"}))
    assert burst.decode_outcome("2026-09-21", line) == DayOutcome("2026-09-21", word, "L", "E")


def test_day_and_run_statuses_serialise_as_their_plain_words():
    assert json.dumps({"status": DayStatus.WOULD_BOOK, "run": RunStatus.NO_SLOTS}) == \
        '{"status": "would-book", "run": "no-slots"}'
    assert f"{DayStatus.BOOKED} {RunStatus.CHECKED}" == "booked checked"


# state-machine-proof: parking-day
def test_a_day_status_outside_the_declared_set_is_refused():
    with pytest.raises(ValueError):
        DayOutcome("2026-09-21", "bokked")


# state-machine-proof: parking-day
def test_an_unknown_worker_word_becomes_an_error_with_the_word_kept():
    for event in ({"done": True, "status": "bokked", "label": "L"}, {"done": True}):
        outcome = burst.decode_outcome("2026-09-21", event)
        assert outcome.status is DayStatus.ERROR
        assert repr(event.get("status")) in outcome.error and outcome.label == event.get("label", "")


# state-machine-proof: parking-run
def test_a_run_status_outside_the_declared_set_is_refused():
    with pytest.raises(ValueError):
        RunResult("checkd")


@pytest.mark.parametrize("statuses, expected", [
    ([], RunStatus.NOTHING_PENDING),
    (["booked"], RunStatus.CHECKED),
    (["covered"], RunStatus.CHECKED),
    (["would-book"], RunStatus.CHECKED),
    (["failed"], RunStatus.CHECKED),
    (["unconfirmed"], RunStatus.CHECKED),
    (["no-slot"], RunStatus.NO_SLOTS),
    (["no-slot", "no-slot"], RunStatus.NO_SLOTS),
    (["no-slot", "booked"], RunStatus.CHECKED),
    (["no-slot", "failed"], RunStatus.CHECKED),
    (["booked", "unconfirmed", "failed"], RunStatus.CHECKED),
    (["error"], RunStatus.ERROR),
    (["unknown"], RunStatus.ERROR),
    (["booked", "error"], RunStatus.ERROR),
    (["no-slot", "unknown"], RunStatus.ERROR),
    (["failed", "unknown", "booked"], RunStatus.ERROR),
])
def test_derive_run_status_table(statuses, expected):
    assert derive_run_status(outcomes(*statuses)) is expected


def test_result_from_outcomes_sorts_each_day_into_its_own_list():
    result = result_from_outcomes(
        ["a", "b", "c", "d", "e"],
        [DayOutcome("a", DayStatus.BOOKED), DayOutcome("b", DayStatus.WOULD_BOOK),
         DayOutcome("c", DayStatus.FAILED, error="x"), DayOutcome("d", DayStatus.UNCONFIRMED),
         DayOutcome("e", DayStatus.UNKNOWN, error="late")])
    assert (result.booked, result.would_book, result.failed, result.unconfirmed) == (["a"], ["b"], ["c"], ["d"])
    assert result.status is RunStatus.ERROR and result.error == "e: late"


def head(*, booked=(), would_book=(), failed=(), unconfirmed=(), status=RunStatus.CHECKED, pending=("2026-09-21",)):
    return RunResult(status, pending=list(pending), booked=list(booked), would_book=list(would_book),
                     failed=list(failed), unconfirmed=list(unconfirmed))


@pytest.mark.parametrize("result, sentence", [
    (head(failed=["2026-09-21"]),
     "Parking sweep Sun 12:00 ⚠️ 1 open day(s), a slot was free but booking failed: 2026-09-21 - check the site."),
    (head(unconfirmed=["2026-09-21"]),
     "Parking sweep Sun 12:00 ⚠️ 1 open day(s), booking sent but not confirmed: 2026-09-21 - check the site."),
    (head(), "Parking sweep Sun 12:00 ⚠️ 1 open day(s), nothing booked - check the site."),
    (head(booked=["2026-09-21"]), "Parking sweep Sun 12:00 ✅ 1 open day(s), booked: 2026-09-21."),
    (head(status=RunStatus.NO_SLOTS), "Parking sweep Sun 12:00 ✅ 1 open day(s), no free slot for them."),
    (RunResult(RunStatus.NO_TOKEN), "Parking sweep Sun 12:00 ❌ login needed (no-token), site not checked."),
])
def test_sweep_summary_has_its_own_sentence_per_state(result, sentence):
    assert poller.sweep_summary(result, NOW) == sentence
