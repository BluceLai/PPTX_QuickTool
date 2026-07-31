from __future__ import annotations

import argparse

from .app import create_main_window


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Run PPTX QuickTool.")
    parser.add_argument(
        "--smoke",
        action="store_true",
        help="Create and close the desktop window without starting the event loop.",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    root = create_main_window()

    if args.smoke:
        root.update()
        root.destroy()
        return 0

    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
