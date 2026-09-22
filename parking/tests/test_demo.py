from datetime import datetime
from zoneinfo import ZoneInfo

from parking import demo
from parking.config import CONFIG_PATH, load_config


class Recorder:
    def __init__(self):
        self.sent = []

    def send(self, text, disposable=False):
        self.sent.append(text)
        return True


def test_demo_walks_through_a_full_poll_and_labels_every_message():
    cfg = load_config(CONFIG_PATH)
    rec = Recorder()
    today = datetime(2026, 9, 21, 9, 0, tzinfo=ZoneInfo("Europe/Madrid"))  # a Monday
    result = demo.run_demo(cfg, demo.DemoNotifier(rec, pause=lambda _s: None), today)
    assert result.status == "checked" and len(result.booked) == demo.DEMO_DAYS
    assert all(text.startswith("[DEMO] ") for text in rec.sent)
    log = " | ".join(rec.sent)
    for step in ("Rehearsal", "poll started", "Free slot found", "Booking", "Parking booked", "Poll finished"):
        assert step in log
    assert log.count("Parking booked") == demo.DEMO_DAYS
