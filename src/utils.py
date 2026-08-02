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

def save_file(data: dict, file_name: str):
    os.makedirs(config_paths(), exist_ok=True)
    config_file = os.path.join(config_paths(), file_name)
    tmp_file = config_file + ".tmp"
    with open(tmp_file, "w", encoding="utf-8") as file:
        json.dump(data, file, ensure_ascii=False, indent=2)
    os.replace(tmp_file, config_file)

def find_suggestions(codes: dict, text: str, limit: int = 20) -> list[dict]:
    q = str(text or "").strip().lower()
    if not q:
        return []

    scored = []
    for key, info in codes.items():
        key = str(key)
        code = str(info.get("code", ""))
        name = str(info.get("name", ""))
        py = str(info.get("py", ""))
        abbr = str(info.get("abbr", ""))
        score = 0
        if key.startswith(q):
            score = 100
        elif code.startswith(q):
            score = 95
        elif q in key:
            score = 90
        elif q in code:
            score = 85
        elif name.startswith(q):
            score = 70
        elif q in name:
            score = 60
        elif py.startswith(q):
            score = 50
        elif q in py:
            score = 40
        elif abbr.startswith(q):
            score = 30
        elif q in abbr:
            score = 20

        if score > 0:
            scored.append((score, 
                           {"key": key, 
                            "market": info.get("market",""),
                            "code": code, 
                            "name": name, 
                            "type": info.get("type","")}
                            ))

    scored.sort(key=lambda pair: (-pair[0], pair[1].get("code", "")))
    return [code for _, code in scored[:limit]]