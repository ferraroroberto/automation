import json
import subprocess
import time

import pytest

from parking import notify
from parking.notify import FleetNotifier

BASE = {"NOTIFY_PYTHON": "py.exe", "NOTIFY_SCRIPT": "notify_send.py"}


def test_explicit_chat_wins_over_category():
    cmd = FleetNotifier({**BASE, "NOTIFY_CATEGORY": "attention", "NOTIFY_CHAT": "-100123"}).command("hi")
    assert cmd == ["py.exe", "notify_send.py", "--chat", "-100123", "--text", "hi"]
    assert "--category" not in cmd


def test_category_used_when_no_chat():
    cmd = FleetNotifier({**BASE, "NOTIFY_CATEGORY": "log"}).command("hi")
    assert cmd == ["py.exe", "notify_send.py", "--category", "log", "--text", "hi"]


def test_category_defaults_to_attention():
    assert FleetNotifier(BASE).command("hi")[2:4] == ["--category", "attention"]


def test_file_command_carries_file_and_text():
    cmd = FleetNotifier({**BASE, "NOTIFY_CHAT": "-100123"}).file_command("hi", "C:/tmp/site.png")
    assert cmd == ["py.exe", "notify_send.py", "--chat", "-100123", "--file", "C:/tmp/site.png", "--text", "hi"]


def test_files_command_carries_all_files_and_text():
    cmd = FleetNotifier({**BASE, "NOTIFY_CHAT": "-100123"}).files_command(
        "hi", ["C:/tmp/a.png", "C:/tmp/b.png"])
    assert cmd == ["py.exe", "notify_send.py", "--chat", "-100123",
                   "--files", "C:/tmp/a.png", "C:/tmp/b.png", "--text", "hi"]


class FakeRun:
    """Stands in for subprocess.run: records argv, answers each send with fresh message ids."""

    def __init__(self):
        self.calls = []
        self.next_id = 100

    def __call__(self, argv, **_kwargs):
        self.calls.append(list(argv))
        stdout = ""
        if "--print-ids" in argv:
            count = len(argv[argv.index("--files") + 1:argv.index("--text")]) if "--files" in argv else 1
            ids = list(range(self.next_id, self.next_id + count))
            self.next_id += count
            stdout = json.dumps(ids) + "\n"
        return subprocess.CompletedProcess(argv, 0, stdout=stdout, stderr="")

    def deletes(self):
        return [c[c.index("--delete-ids") + 1:] for c in self.calls if "--delete-ids" in c]

    def texts(self):
        return [c[c.index("--text") + 1] for c in self.calls if "--text" in c]


@pytest.fixture
def fake_run(monkeypatch):
    fake = FakeRun()
    monkeypatch.setattr(notify.subprocess, "run", fake)
    return fake


def ledger(path):
    return json.loads(path.read_text(encoding="utf-8"))["disposable"]


def test_only_disposable_sends_are_recorded(fake_run, tmp_path):
    path = tmp_path / "sent.json"
    n = FleetNotifier(BASE, ledger_path=path)
    n.send("booked", disposable=False)
    n.send_files("sweep", ["a.png", "b.png"], disposable=True)
    assert ledger(path) == [101, 102]  # the booking (100) is never recorded, so never deleted


def test_first_send_of_a_run_deletes_the_previous_runs_messages_once(fake_run, tmp_path):
    path = tmp_path / "sent.json"
    path.write_text(json.dumps({"disposable": [7, 8, 9]}), encoding="utf-8")
    n = FleetNotifier(BASE, ledger_path=path)
    n.send("one", disposable=True)
    n.send("two", disposable=True)
    assert fake_run.deletes() == [["7", "8", "9"]]  # once, before the first send
    assert fake_run.calls[0][-4:] == ["--delete-ids", "7", "8", "9"]
    assert ledger(path) == [100, 101]  # only this run's messages remain for the next run


def test_ledger_is_cleared_even_when_the_delete_fails(monkeypatch, tmp_path):
    path = tmp_path / "sent.json"
    path.write_text(json.dumps({"disposable": [7]}), encoding="utf-8")

    def failing(argv, **_kwargs):
        return subprocess.CompletedProcess(argv, 1, stdout="[]", stderr="")

    monkeypatch.setattr(notify.subprocess, "run", failing)
    FleetNotifier(BASE, ledger_path=path).send("x", disposable=True)
    assert ledger(path) == []  # never retried: Telegram can't delete messages older than 48h


def test_without_a_ledger_nothing_is_tracked_or_deleted(fake_run):
    FleetNotifier(BASE).send("x", disposable=True)
    assert "--print-ids" not in fake_run.calls[0] and fake_run.deletes() == []


def test_pings_queued_while_waiting_go_out_as_one_message(fake_run, tmp_path, monkeypatch):
    monkeypatch.setattr(notify, "PING_MIN_INTERVAL_SECONDS", 0.3)
    n = FleetNotifier(BASE, ledger_path=tmp_path / "sent.json")
    n.send("armed", disposable=True)  # sets the pace: the next message waits out the interval
    started = time.monotonic()
    for i in range(3):
        n.ping(f"check {i}")
    assert time.monotonic() - started < 0.1  # ping() never blocks the caller
    n.close()
    assert fake_run.texts() == ["armed", "check 0\ncheck 1\ncheck 2"]
    assert ledger(tmp_path / "sent.json") == [100, 101]  # pings are disposable
