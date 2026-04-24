import json
import os
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple


def _utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def _safe_mkdir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def _atomic_write_json(path: Path, payload: Dict[str, Any]) -> None:
    """
    Write JSON atomically (best-effort) to avoid partial files.
    Uses write-then-replace to support Windows.
    """
    _safe_mkdir(path.parent)
    tmp_path = path.with_suffix(path.suffix + f".{uuid.uuid4().hex}.tmp")
    with tmp_path.open("w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
        f.flush()
        os.fsync(f.fileno())
    os.replace(tmp_path, path)


def _read_json(path: Path) -> Dict[str, Any]:
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)


def sessions_dir(output_path: Path) -> Path:
    return output_path / "edit_sessions"


def session_path(output_path: Path, session_id: str) -> Path:
    return sessions_dir(output_path) / f"{session_id}.json"


def create_session(
    *,
    output_path: Path,
    client_name: str,
    process_name: str,
    master_id: Optional[int],
    mappings: List[Dict[str, Any]],
    created_by: Optional[str] = None,
) -> Dict[str, Any]:
    session_id = uuid.uuid4().hex
    now = _utc_now_iso()
    payload: Dict[str, Any] = {
        "session_id": session_id,
        "status": "draft",
        "client_name": client_name,
        "process_name": process_name,
        "master_id": master_id,
        "created_at": now,
        "updated_at": now,
        "created_by": created_by,
        "approved_at": None,
        "approved_by": None,
        "mappings": mappings,
        "approval_result": None,
        "history": [
            {
                "ts": now,
                "action": "create",
                "by": created_by,
            }
        ],
    }
    _atomic_write_json(session_path(output_path, session_id), payload)
    return payload


def get_session(*, output_path: Path, session_id: str) -> Dict[str, Any]:
    path = session_path(output_path, session_id)
    if not path.exists():
        raise FileNotFoundError(session_id)
    return _read_json(path)


def list_sessions(*, output_path: Path) -> List[Dict[str, Any]]:
    root = sessions_dir(output_path)
    if not root.exists():
        return []
    sessions: List[Dict[str, Any]] = []
    for fp in sorted(root.glob("*.json"), key=lambda p: p.stat().st_mtime, reverse=True):
        try:
            data = _read_json(fp)
            sessions.append(
                {
                    "session_id": data.get("session_id") or fp.stem,
                    "status": data.get("status", "draft"),
                    "client_name": data.get("client_name"),
                    "process_name": data.get("process_name"),
                    "master_id": data.get("master_id"),
                    "created_at": data.get("created_at"),
                    "updated_at": data.get("updated_at"),
                    "approved_at": data.get("approved_at"),
                }
            )
        except Exception:
            # Skip corrupted sessions rather than failing the whole listing.
            continue
    return sessions


def update_session(
    *,
    output_path: Path,
    session_id: str,
    mappings: Optional[List[Dict[str, Any]]] = None,
    client_name: Optional[str] = None,
    process_name: Optional[str] = None,
    master_id: Optional[int] = None,
    updated_by: Optional[str] = None,
    note: Optional[str] = None,
) -> Dict[str, Any]:
    data = get_session(output_path=output_path, session_id=session_id)
    if data.get("status") != "draft":
        raise PermissionError("Only draft sessions can be updated.")

    if mappings is not None:
        data["mappings"] = mappings
    if client_name is not None:
        data["client_name"] = client_name
    if process_name is not None:
        data["process_name"] = process_name
    if master_id is not None:
        data["master_id"] = master_id

    now = _utc_now_iso()
    data["updated_at"] = now
    data.setdefault("history", []).append(
        {
            "ts": now,
            "action": "update",
            "by": updated_by,
            "note": note,
        }
    )
    _atomic_write_json(session_path(output_path, session_id), data)
    return data


def approve_session(
    *,
    output_path: Path,
    session_id: str,
    approved_by: Optional[str] = None,
    approval_result: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    data = get_session(output_path=output_path, session_id=session_id)
    if data.get("status") != "draft":
        raise PermissionError("Only draft sessions can be approved.")

    now = _utc_now_iso()
    data["status"] = "approved"
    data["approved_at"] = now
    data["approved_by"] = approved_by
    data["updated_at"] = now
    data["approval_result"] = approval_result
    data.setdefault("history", []).append(
        {
            "ts": now,
            "action": "approve",
            "by": approved_by,
        }
    )
    _atomic_write_json(session_path(output_path, session_id), data)
    return data


def delete_session(*, output_path: Path, session_id: str) -> None:
    path = session_path(output_path, session_id)
    if path.exists():
        path.unlink()
