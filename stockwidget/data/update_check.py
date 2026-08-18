import re
import requests

GITHUB_REPO = "sbr0574/StockWidget"
PROJECT_URL = "https://github.com/sbr0574/StockWidget"
GITEE_URL = "https://gitee.com/sbr0574/StockWidget"
RELEASES_URL = PROJECT_URL + "/releases"
LICENSE_URL = PROJECT_URL + "/blob/main/LICENSE"
README_URL = PROJECT_URL + "#readme"


def _version_tuple(version: str) -> tuple:
    nums = [int(p) for p in re.findall(r"\d+", str(version or ""))][:3]
    while len(nums) < 3:
        nums.append(0)
    return tuple(nums)


def get_latest_release() -> dict | None:
    """获取 GitHub 最新 Release 信息；失败返回 None。"""
    try:
        resp = requests.get(
            f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest",
            timeout=5,
            headers={"User-Agent": "StockWidget"},
        )
        if resp.status_code != 200:
            return None
        data = resp.json()
        tag = (data.get("tag_name") or "").lstrip("vV")
        if not tag:
            return None
        return {
            "version": tag,
            "url": data.get("html_url") or f"{RELEASES_URL}/tag/{tag}",
        }
    except Exception:
        return None


def get_update_info(current_version) -> tuple[bool, str | None, str | None]:
    """返回 (是否有更新, 最新版本号, 最新版本地址)。"""
    info = get_latest_release()
    if not info:
        return False, None, None
    has_update = _version_tuple(info["version"]) > _version_tuple(current_version)
    return has_update, info["version"], info["url"]


def check_for_update(current_version) -> bool:
    """检查 GitHub Releases 是否有比当前版本更新的版本（兼容旧调用）。"""
    has_update, _, _ = get_update_info(current_version)
    return has_update