import json
import os

APP_NAME = "StockWidget"

def config_paths() -> tuple[str, str]:
    config_dir = os.path.join(os.getenv("APPDATA") or os.path.expanduser("~"), APP_NAME)
    return config_dir

def load_file(file_name: str, except_ret: dict = {}) -> dict:
    try:
        with open(os.path.join(config_paths(), file_name), "r", encoding="utf-8") as file:
            return json.load(file)
    except Exception:
        return except_ret

def save_file(cfg: dict, file_name: str):
    os.makedirs(config_paths(), exist_ok=True)
    config_file = os.path.join(config_paths(), file_name)
    tmp_file = config_file + ".tmp"
    with open(tmp_file, "w", encoding="utf-8") as file:
        json.dump(cfg, file, ensure_ascii=False, indent=2)
    os.replace(tmp_file, config_file)
