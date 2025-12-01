
import json
from pathlib import Path
from typing import Any, Dict

CONFIG_PATH = Path(__file__).resolve().parent.parent / "settings.json"

DEFAULT_SETTINGS: Dict[str, Any] = {
    "last_open_dir": "",
    "last_save_dir": "",
}


def load_settings() -> Dict[str, Any]:
    if not CONFIG_PATH.exists():
        return DEFAULT_SETTINGS.copy()

    try:
        with CONFIG_PATH.open("r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        return DEFAULT_SETTINGS.copy()

    merged = DEFAULT_SETTINGS.copy()
    merged.update(data)
    return merged


def save_settings(settings: Dict[str, Any]) -> None:
    try:
        with CONFIG_PATH.open("w", encoding="utf-8") as f:
            json.dump(settings, f, indent=2)
    except Exception:
        pass
