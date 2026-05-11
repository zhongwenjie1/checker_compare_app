# -*- coding: utf-8 -*-
import json
from pathlib import Path


def read_local_version(base_dir: str | Path):
    version_path = Path(base_dir) / "version.json"
    if not version_path.exists():
        return {}
    try:
        return json.loads(version_path.read_text(encoding="utf-8"))
    except Exception:
        return {}
