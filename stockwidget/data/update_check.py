import re

import requests

from stockwidget.data.network_errors import request_error_message


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


def _release_version(api_url: str, *, errors: list[str] | None = None) -> str | None:
    def failed(message):
        if errors is not None:
            errors.append(message)
        return None

    try:
        response = requests.get(
            api_url,
            timeout=(2.5, 3),
            headers={"User-Agent": "StockWidget"},
        )
        if response.status_code != 200:
            return failed(request_error_message(requests.HTTPError(response=response)))
        data = response.json()
        if not isinstance(data, dict):
            return failed("响应内容不是版本信息")
        tag = str(data.get("tag_name") or "").lstrip("vV")
        if not tag:
            return failed("响应中缺少版本号")
        return tag
    except requests.RequestException as exc:
        return failed(request_error_message(exc))
    except ValueError:
        return failed("响应不是有效 JSON")


def get_latest_release(*, errors: list[str] | None = None) -> str | None:
    """按 GitHub、Gitee 顺序获取最新 Release。"""
    api_urls = (
        f"https://api.github.com/repos/{REPOSITORY}/releases/latest",
        f"https://gitee.com/api/v5/repos/{REPOSITORY}/releases/latest",
    )
    for api_url in api_urls:
        source_errors = []
        version = (_release_version(api_url) if errors is None
                   else _release_version(api_url, errors=source_errors))
        if version is not None:
            return version
        if errors is not None:
            source = "GitHub" if "api.github.com" in api_url else "Gitee"
            errors.append(f"{source}：" + ("；".join(source_errors) or "未获取到版本信息"))
    return None


def get_update_info(current_version, *, errors: list[str] | None = None) -> tuple[bool, str | None]:
    """返回 (是否有更新, 最新版本号)。"""
    latest_version = (get_latest_release() if errors is None
                      else get_latest_release(errors=errors))
    if not latest_version:
        return False, None
    has_update = _version_tuple(latest_version) > _version_tuple(current_version)
    return has_update, latest_version
