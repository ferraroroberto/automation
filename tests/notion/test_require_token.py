"""`utils.require_notion_api_token` returns the `.env` token and fails loudly when it is missing."""

import sys
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "tools" / "notion"))

import utils as notion_utils  # noqa: E402


class RequireTokenTests(unittest.TestCase):
    def test_returns_the_environment_token(self) -> None:
        with mock.patch.object(notion_utils, "load_dotenv"), mock.patch.dict(
            "os.environ", {"NOTION_API_TOKEN": "secret_abc"}
        ):
            self.assertEqual(notion_utils.require_notion_api_token(), "secret_abc")

    def test_raises_when_the_token_is_unset_or_empty(self) -> None:
        for env in ({}, {"NOTION_API_TOKEN": ""}):
            with mock.patch.object(notion_utils, "load_dotenv"), mock.patch.dict(
                "os.environ", env, clear=True
            ):
                with self.assertRaises(ValueError):
                    notion_utils.require_notion_api_token()


if __name__ == "__main__":
    unittest.main()
