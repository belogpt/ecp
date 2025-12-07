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
        # В собранной версии стараемся писать в профайл пользователя, чтобы не упереться
        # в запрет записи в Program Files.
        if os.name == "nt":
            base_dir = (
                os.getenv("LOCALAPPDATA")
                or os.getenv("APPDATA")
                or os.path.expanduser("~")
            )
            base_path = os.path.join(base_dir, "ep_viewer")
        else:
            base_path = os.path.expanduser("~/.local/share/ep_viewer")
    else:
        base_path = os.path.abspath(os.path.dirname(__file__))

    os.makedirs(base_path, exist_ok=True)
    return os.path.join(base_path, filename)
