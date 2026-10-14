"""Launcher: `python gui.py` mo giao dien ma khong can pip install.
Code that su nam o src/dataextraction/gui.py."""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "src"))

from dataextraction.gui import main  # noqa: E402

if __name__ == "__main__":
    main()
