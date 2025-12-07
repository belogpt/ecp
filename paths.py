import os
import sys


def get_resource_path(relative_path: str) -> str:
    """Return absolute path to a bundled read-only resource."""
    if hasattr(sys, "_MEIPASS"):
        base_path = sys._MEIPASS  # type: ignore[attr-defined]
    else:
        base_path = os.path.abspath(os.path.dirname(__file__))
    return os.path.join(base_path, relative_path)


def get_data_path(filename: str) -> str:
    """Return absolute path for files that can be created or modified."""
    if hasattr(sys, "frozen") or getattr(sys, "frozen", False):
        base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.abspath(os.path.dirname(__file__))
    return os.path.join(base_path, filename)
