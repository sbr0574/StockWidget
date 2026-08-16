import re, requests

GITHUB_REPO = "sbr0574/StockWidget"
PROJECT_URL = "https://github.com/sbr0574/StockWidget"


def _version_tuple(version: str) -> tuple:
    nums = [int(p) for p in re.findall(r"\d+", str(version or ""))][:3]
    while len(nums) < 3:
        nums.append(0)
    return tuple(nums)


def check_for_update(current_version) -> bool:
    """检查 GitHub Releases 是否有比当前版本更新的版本。

    网络异常 / 无 Releases / 版本号解析失败时均视为无更新。
    """
    try:
        resp = requests.get(
            f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest",
            timeout=5,
            headers={"User-Agent": "StockWidget"},
        )
        if resp.status_code != 200:
            return False
        tag = (resp.json().get("tag_name") or "").lstrip("vV")
        return _version_tuple(tag) > _version_tuple(current_version)
    except Exception:
        return False
