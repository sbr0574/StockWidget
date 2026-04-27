import json
import os


def config_paths(app_name: str) -> tuple[str, str]:
    config_dir = os.path.join(os.getenv("APPDATA") or os.path.expanduser("~"), app_name)
    config_file = os.path.join(config_dir, "SW_config.json")
    return config_dir, config_file


def load_config(app_name: str) -> dict:
    _, config_file = config_paths(app_name)
    try:
        with open(config_file, "r", encoding="utf-8") as file:
            return json.load(file)
    except Exception:
        return {}


def save_config(app_name: str, cfg: dict):
    config_dir, config_file = config_paths(app_name)
    os.makedirs(config_dir, exist_ok=True)
    tmp_file = config_file + ".tmp"
    with open(tmp_file, "w", encoding="utf-8") as file:
        json.dump(cfg, file, ensure_ascii=False, indent=2)
    os.replace(tmp_file, config_file)
