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
