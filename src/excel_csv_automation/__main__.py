"""Launch the local application or check the installed environment."""

import argparse
import importlib
import sys
from pathlib import Path


def main() -> int:
    parser = argparse.ArgumentParser(description="Excel and CSV processing automation")
    parser.add_argument(
        "--check", action="store_true", help="Check installed dependencies and exit"
    )
    args = parser.parse_args()
    if args.check:
        for module_name in ("streamlit", "pandas", "openpyxl"):
            module = importlib.import_module(module_name)
            print(f"{module_name}: {module.__version__}")
        print("Environment ready. Transaction processing is not implemented yet.")
        return 0

    from streamlit.web import cli

    app_path = Path(__file__).with_name("app.py")
    sys.argv = ["streamlit", "run", str(app_path), "--server.address=127.0.0.1"]
    return cli.main()


if __name__ == "__main__":
    raise SystemExit(main())
