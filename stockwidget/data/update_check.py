import re
import requests

GITHUB_REPO = "sbr0574/StockWidget"
GITEE_REPO = "sbr0574/StockWidget"
GITHUB = "github"
GITEE = "gitee"


def project_links(source: str = GITHUB) -> dict[str, str]:
    """返回指定托管平台的项目、发布、文档和反馈链接。"""
    if source == GITEE:
        project = f"https://gitee.com/{GITEE_REPO}"
        return {
            "project": project,
            "releases": project + "/releases",
            "license": project + "/blob/main/LICENSE",
            "issues": project + "/issues",
            "readme": project,
            "repository_label": "Gitee仓库",
        }
    project = f"https://github.com/{GITHUB_REPO}"
    return {
        "project": project,
        "releases": project + "/releases",
        "license": project + "/blob/main/LICENSE",
        "issues": project + "/issues",
        "readme": project + "#readme",
        "repository_label": "GitHub仓库",
    }


def github_available(timeout=5):
    try:
        r = requests.get(
            "https://github.com/",
            timeout=timeout,
            headers={"User-Agent": "Mozilla/5.0"},
        )
        return r.ok and "GitHub" in r.text
    except requests.RequestException:
        return False


def _version_tuple(version: str) -> tuple:
    nums = [int(p) for p in re.findall(r"\d+", str(version or ""))][:3]
    while len(nums) < 3:
        nums.append(0)
    return tuple(nums)


def get_latest_release(source: str = GITHUB) -> dict | None:
    """获取指定托管平台的最新 Release；未发布或访问失败返回 None。"""
    links = project_links(source)
    if source == GITEE:
        api_url = f"https://gitee.com/api/v5/repos/{GITEE_REPO}/releases/latest"
    else:
        api_url = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
    try:
        resp = requests.get(
            api_url,
            timeout=(2.5, 3),
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
            "url": data.get("html_url") or f"{links['releases']}/tag/{tag}",
        }
    except (requests.RequestException, ValueError):
        return None


def get_update_info(current_version, source: str = GITHUB) -> tuple[bool, str | None, str | None]:
    """返回 (是否有更新, 最新版本号, 最新版本地址)。"""
    info = get_latest_release(source)
    if not info:
        return False, None, None
    has_update = _version_tuple(info["version"]) > _version_tuple(current_version)
    return has_update, info["version"], info["url"]
