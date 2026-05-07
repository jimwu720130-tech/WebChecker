"""GitHub Gist 雲端持久層。

Streamlit Cloud 容器檔案系統屬於暫時性，每次重啟會還原到 git clone 之初始狀態，
導致使用者於 UI 新增的常用網站（favorites.json）／排除規則（config.json）
被丟失。本模組將該兩檔以 GitHub Gist 作為「資料系統 of record」：

  - 容器啟動時：app.py 觸發 ``pull_to_local_files`` 將 Gist 上之最新內容覆寫至工作目錄之
    favorites.json / config.json，使後續 ``load_*`` 直接讀本機檔即可。
  - 使用者新增／修改：webchecker_core.save_* 寫完本機檔後，呼叫 ``save`` 同步至 Gist。

設計原則
========
- 不引入第三方套件（urllib + json 即可）。
- 所有失敗皆**最佳努力**（best-effort）：失敗時 stderr 輸出訊息，**不拋出例外**，
  以免凍結 Streamlit UI。
- 與 streamlit 解耦：本模組僅讀環境變數，由 app.py 自 ``st.secrets`` 將憑證注入：

    WC_GIST_TOKEN: GitHub Personal Access Token，需具 ``gist`` scope
    WC_GIST_ID:    儲存資料用之 Gist ID

  未設定其一者，``is_enabled()`` 為 False，所有操作為 no-op。
"""

from __future__ import annotations

import os
import sys
import json
import urllib.request
import urllib.error
from typing import Dict, Optional, Tuple

_GIST_API = "https://api.github.com/gists/{gid}"
_HTTP_TIMEOUT = 20

# 模組 cache：避免一個 Streamlit rerun 內反覆對同一 gist 發 GET（呼叫端應自行 invalidate）
_FETCH_CACHE: Optional[Dict[str, str]] = None


def _credentials() -> Optional[Tuple[str, str]]:
    token = (os.environ.get("WC_GIST_TOKEN") or "").strip()
    gid = (os.environ.get("WC_GIST_ID") or "").strip()
    if not token or not gid:
        return None
    return token, gid


def is_enabled() -> bool:
    """是否已設定 Gist 憑證。未設定時所有寫入操作將靜默 no-op。"""
    return _credentials() is not None


def _request(method: str, url: str, token: str, body: Optional[dict] = None) -> dict:
    headers = {
        "Authorization": f"token {token}",
        "Accept": "application/vnd.github+json",
        "User-Agent": "WebChecker-cloud-persistence",
        "X-GitHub-Api-Version": "2022-11-28",
    }
    data: Optional[bytes] = None
    if body is not None:
        data = json.dumps(body).encode("utf-8")
        headers["Content-Type"] = "application/json"
    req = urllib.request.Request(url, data=data, headers=headers, method=method)
    with urllib.request.urlopen(req, timeout=_HTTP_TIMEOUT) as resp:
        raw = resp.read().decode("utf-8")
    return json.loads(raw) if raw else {}


def fetch_all(use_cache: bool = True) -> Dict[str, str]:
    """讀取 Gist 上之所有檔案內容；回傳 ``{filename: text_content}``。

    失敗時回傳 ``{}``，不拋例外。``use_cache=True`` 時同一 Python 進程內僅實際 GET 一次。
    Gist 過大被 truncate 時會跟進 ``raw_url`` 取完整內容。
    """
    global _FETCH_CACHE
    if use_cache and _FETCH_CACHE is not None:
        return dict(_FETCH_CACHE)
    creds = _credentials()
    if not creds:
        return {}
    token, gid = creds
    try:
        info = _request("GET", _GIST_API.format(gid=gid), token)
    except Exception as e:
        sys.stderr.write(
            f"[cloud_persistence] fetch_all failed: {type(e).__name__}: {e}\n"
        )
        return {}
    files = info.get("files") or {}
    out: Dict[str, str] = {}
    for name, meta in files.items():
        if not isinstance(meta, dict):
            continue
        content = meta.get("content")
        if content is None and meta.get("truncated"):
            raw_url = meta.get("raw_url")
            if raw_url:
                try:
                    with urllib.request.urlopen(raw_url, timeout=_HTTP_TIMEOUT) as resp:
                        content = resp.read().decode("utf-8")
                except Exception as e:
                    sys.stderr.write(
                        f"[cloud_persistence] follow raw_url failed for {name}: "
                        f"{type(e).__name__}: {e}\n"
                    )
                    continue
        if isinstance(content, str):
            out[name] = content
    _FETCH_CACHE = dict(out)
    return out


def invalidate_cache() -> None:
    """寫入後可呼叫，使下次 fetch_all 重抓最新內容。"""
    global _FETCH_CACHE
    _FETCH_CACHE = None


def save(filename: str, content: str) -> bool:
    """更新 Gist 中指定檔案的內容；成功 True、失敗 False。

    GitHub Gist 之 PATCH 行為：files 中**未列出**之檔案不受影響；列出但 content 為空字串
    時 GitHub 會視為刪除，故呼叫端應確保傳入合法 JSON 字串而非空。
    """
    creds = _credentials()
    if not creds:
        return False
    token, gid = creds
    if not isinstance(content, str) or not content.strip():
        sys.stderr.write(
            f"[cloud_persistence] save '{filename}' aborted: empty content\n"
        )
        return False
    try:
        _request(
            "PATCH",
            _GIST_API.format(gid=gid),
            token,
            {"files": {filename: {"content": content}}},
        )
    except Exception as e:
        sys.stderr.write(
            f"[cloud_persistence] save '{filename}' failed: {type(e).__name__}: {e}\n"
        )
        return False
    invalidate_cache()
    return True


def pull_to_local_files(target_files: Dict[str, str]) -> int:
    """將 Gist 上之檔案內容覆寫到指定本地路徑；回傳實際成功覆寫之檔案數。

    ``target_files``: ``{gist_filename: local_filesystem_path}``。

    安全性
    ------
    僅當 Gist 內容可成功 ``json.loads`` 才覆寫本機，避免 Gist 被誤改成壞 JSON 時連同
    本機可用副本一起破壞。憑證未設或 Gist 抓取失敗時為 no-op。
    """
    if not is_enabled():
        return 0
    data = fetch_all(use_cache=False)
    if not data:
        return 0
    written = 0
    for gname, local_path in target_files.items():
        content = data.get(gname)
        if content is None:
            continue
        try:
            json.loads(content)
        except Exception as e:
            sys.stderr.write(
                f"[cloud_persistence] pull_to_local skipped '{gname}': "
                f"invalid JSON ({type(e).__name__}: {e})\n"
            )
            continue
        try:
            with open(local_path, "w", encoding="utf-8") as f:
                f.write(content)
            written += 1
        except Exception as e:
            sys.stderr.write(
                f"[cloud_persistence] pull_to_local write fail '{local_path}': "
                f"{type(e).__name__}: {e}\n"
            )
    return written
