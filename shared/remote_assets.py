"""Remote asset helpers used by the PPT renderer."""

from __future__ import annotations

import mimetypes
import os
from pathlib import Path
from tempfile import NamedTemporaryFile
from typing import Iterable
from urllib.parse import urlparse
from urllib.request import Request, urlopen


_CONTENT_TYPE_SUFFIX = {
    "image/jpeg": ".jpg",
    "image/jpg": ".jpg",
    "image/png": ".png",
    "image/gif": ".gif",
    "image/webp": ".webp",
    "image/svg+xml": ".svg",
}


def _safe_suffix(url: str, content_type: str | None, fallback_suffix: str) -> str:
    parsed = urlparse(url)
    suffix = Path(parsed.path).suffix.lower()
    if suffix and 1 < len(suffix) <= 10:
        return suffix

    if content_type:
        media_type = content_type.split(";", 1)[0].strip().lower()
        if media_type in _CONTENT_TYPE_SUFFIX:
            return _CONTENT_TYPE_SUFFIX[media_type]
        guessed = mimetypes.guess_extension(media_type)
        if guessed:
            return guessed

    return fallback_suffix if fallback_suffix.startswith(".") else f".{fallback_suffix}"


def download_remote_asset(
    url: str,
    *,
    timeout: float = 6.0,
    prefix: str = "asset_",
    fallback_suffix: str = ".bin",
) -> str:
    """Download a remote HTTP(S) asset into a temp file and return its path."""

    parsed = urlparse(url)
    if parsed.scheme not in {"http", "https"}:
        raise ValueError(f"unsupported remote asset scheme: {parsed.scheme or '<empty>'}")

    request = Request(
        url,
        headers={
            "User-Agent": "pro-ppt-gen/1.6.4",
            "Accept": "image/*,*/*;q=0.8",
        },
    )
    with urlopen(request, timeout=timeout) as response:
        content_type = response.headers.get("Content-Type")
        suffix = _safe_suffix(url, content_type, fallback_suffix)
        payload = response.read()

    tmp = NamedTemporaryFile(delete=False, prefix=prefix, suffix=suffix)
    try:
        with tmp:
            tmp.write(payload)
        return tmp.name
    except Exception:
        try:
            os.unlink(tmp.name)
        except OSError:
            pass
        raise


def cleanup_temp_files(paths: Iterable[str | os.PathLike[str]]) -> None:
    """Best-effort cleanup for renderer-created temp assets."""

    for path in paths:
        try:
            os.unlink(path)
        except FileNotFoundError:
            continue
        except OSError:
            continue


__all__ = ["download_remote_asset", "cleanup_temp_files"]
