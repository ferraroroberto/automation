import json
from datetime import time
from pathlib import Path

import pytest

from parking.config import apply_env_overrides, load_config


@pytest.fixture
def cfg():
    return load_config(Path(__file__).resolve().parent.parent / "config.json")


def test_shipped_config_is_dry_run(cfg):
    assert cfg.dry_run is True


def test_shipped_config_parses_thursday_burst(cfg):
    burst = cfg.thursday_burst
    assert burst is not None
    assert burst.enabled is True
    assert burst.target == time(16, 0)
    assert burst.watch_before_minutes == 6
    assert burst.duration_seconds == 90
    assert burst.poll_interval_seconds == 1.5


def test_missing_thursday_burst_block_is_none(tmp_path):
    raw = json.loads((Path(__file__).resolve().parent.parent / "config.json").read_text())
    del raw["thursday_burst"]
    path = tmp_path / "config.json"
    path.write_text(json.dumps(raw))
    assert load_config(path).thursday_burst is None


@pytest.mark.parametrize("value", ["false", "FALSE", "0", "no", " False "])
def test_explicit_false_turns_dry_run_off(cfg, value):
    assert apply_env_overrides(cfg, {"PARKING_DRY_RUN": value}).dry_run is False


@pytest.mark.parametrize("value", ["", "true", "1", "yes", "flase", "off", "garbage"])
def test_anything_else_keeps_dry_run_on(cfg, value):
    assert apply_env_overrides(cfg, {"PARKING_DRY_RUN": value}).dry_run is True


def test_unset_keeps_config_value(cfg):
    assert apply_env_overrides(cfg, {}).dry_run is True
