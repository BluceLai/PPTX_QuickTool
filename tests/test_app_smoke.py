from __future__ import annotations

import unittest
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"

if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from pptx_quicktool import __version__
from pptx_quicktool.app import APP_TITLE, create_main_window
from pptx_quicktool.__main__ import main


class AppSmokeTests(unittest.TestCase):
    def test_package_has_version(self) -> None:
        self.assertEqual(__version__, "0.1.0")

    def test_create_main_window_sets_title(self) -> None:
        root = create_main_window()
        try:
            self.assertEqual(root.title(), APP_TITLE)
        finally:
            root.destroy()

    def test_smoke_mode_exits_successfully(self) -> None:
        self.assertEqual(main(["--smoke"]), 0)


if __name__ == "__main__":
    unittest.main()
