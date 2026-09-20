import json
from pathlib import Path

import pytest

from parking.config import apply_env_overrides, load_config


@pytest.fixture
def cfg():
    return load_config(Path(__file__).resolve().parent.parent / "config.json")


def test_shipped_config_is_dry_run(cfg):
    assert cfg.dry_run is True


@pytest.mark.parametrize("value", ["false", "FALSE", "0", "no", " False "])
def test_explicit_false_turns_dry_run_off(cfg, value):
    assert apply_env_overrides(cfg, {"PARKING_DRY_RUN": value}).dry_run is False


@pytest.mark.parametrize("value", ["", "true", "1", "yes", "flase", "off", "garbage"])
def test_anything_else_keeps_dry_run_on(cfg, value):
    assert apply_env_overrides(cfg, {"PARKING_DRY_RUN": value}).dry_run is True


def test_unset_keeps_config_value(cfg):
    assert apply_env_overrides(cfg, {}).dry_run is True
