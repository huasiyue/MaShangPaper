from __future__ import annotations

import re
import shutil
from dataclasses import dataclass
from pathlib import Path
from uuid import uuid4

from fastapi import UploadFile

from app.core.settings import ASSETS_DIR


ASSET_URL_PATTERN = re.compile(r"https?://[^)\s]+/api/assets/(?P<asset_id>[a-f0-9]+)|/api/assets/(?P<asset_id_rel>[a-f0-9]+)")


@dataclass(slots=True)
class StoredAsset:
    asset_id: str
    filename: str
    path: Path


def _safe_filename(name: str) -> str:
    base = Path(name).name or "asset"
    return re.sub(r"[^A-Za-z0-9._-]+", "_", base)


async def save_asset(upload: UploadFile) -> StoredAsset:
    asset_id = uuid4().hex
    filename = _safe_filename(upload.filename or "asset")
    target = ASSETS_DIR / f"{asset_id}_{filename}"
    content = await upload.read()
    target.write_bytes(content)
    await upload.close()
    return StoredAsset(asset_id=asset_id, filename=filename, path=target)


def find_asset_path(asset_id: str) -> Path:
    candidates = sorted(ASSETS_DIR.glob(f"{asset_id}_*"))
    if not candidates:
        raise FileNotFoundError(f"Asset not found: {asset_id}")
    return candidates[0]


def rename_asset(asset_id: str, filename: str) -> StoredAsset:
    current = find_asset_path(asset_id)
    next_filename = _safe_filename(filename)
    target = current.with_name(f"{asset_id}_{next_filename}")
    current.rename(target)
    return StoredAsset(asset_id=asset_id, filename=next_filename, path=target)


def delete_asset(asset_id: str) -> None:
    target = find_asset_path(asset_id)
    target.unlink(missing_ok=True)


def store_asset_from_path(path: Path) -> StoredAsset:
    asset_id = uuid4().hex
    filename = _safe_filename(path.name)
    target = ASSETS_DIR / f"{asset_id}_{filename}"
    shutil.copy2(path, target)
    return StoredAsset(asset_id=asset_id, filename=filename, path=target)


def rewrite_asset_urls_to_local_paths(content: str) -> str:
    def replacer(match: re.Match[str]) -> str:
        asset_id = match.group("asset_id") or match.group("asset_id_rel")
        if not asset_id:
            return match.group(0)

        try:
            return str(find_asset_path(asset_id))
        except FileNotFoundError:
            return match.group(0)

    return ASSET_URL_PATTERN.sub(replacer, content)
