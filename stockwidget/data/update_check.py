import re

import requests


REPOSITORY = "sbr0574/StockWidget"


def project_links(use_gitee: bool = False) -> dict[str, str]:
    """返回 GitHub 或 Gitee 上的项目链接。"""
    if use_gitee:
        project = f"https://gitee.com/{REPOSITORY}"
        return {
            "project": project,
            "releases": project + "/releases",
            "license": project + "/blob/main/LICENSE",
            "issues": project + "/issues",
            "readme": project,
        }
    project = f"https://github.com/{REPOSITORY}"
    return {
        "project": project,
        "releases": project + "/releases",
        "license": project + "/blob/main/LICENSE",
        "issues": project + "/issues",
        "readme": project + "#readme",
    }


def github_available(timeout=5) -> bool:
    try:
        response = requests.get(
            "https://github.com/",
            timeout=timeout,
            headers={"User-Agent": "Mozilla/5.0"},
        )
        return response.ok and "GitHub" in response.text
    except requests.RequestException:
        return False


def _version_tuple(version: str) -> tuple:
    nums = [int(part) for part in re.findall(r"\d+", str(version or ""))][:3]
    while len(nums) < 3:
        nums.append(0)
    return tuple(nums)


def _release_version(api_url: str) -> str | None:
    try:
        response = requests.get(
            api_url,
            timeout=(2.5, 3),
            headers={"User-Agent": "StockWidget"},
        )
        if response.status_code != 200:
            return None
        data = response.json()
        tag = str(data.get("tag_name") or "").lstrip("vV")
        if not tag:
            return None
        return tag
    except (requests.RequestException, ValueError):
        return None


def get_latest_release() -> str | None:
    """按 GitHub、Gitee 顺序获取最新 Release。"""
    api_urls = (
        f"https://api.github.com/repos/{REPOSITORY}/releases/latest",
        f"https://gitee.com/api/v5/repos/{REPOSITORY}/releases/latest",
    )
    for api_url in api_urls:
        version = _release_version(api_url)
        if version is not None:
            return version
    return None


def get_update_info(current_version) -> tuple[bool, str | None]:
    """返回 (是否有更新, 最新版本号)。"""
    latest_version = get_latest_release()
    if not latest_version:
        return False, None
    has_update = _version_tuple(latest_version) > _version_tuple(current_version)
    return has_update, latest_version
