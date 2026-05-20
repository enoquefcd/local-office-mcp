#!/usr/bin/env python3
"""
Read Microsoft Teams local IndexedDB cache (read-only, no API needed).
Requires: pip install git+https://github.com/cclgroupltd/ccl_chrome_indexeddb.git

Works while Teams is running — copies LevelDB to a temp folder before reading.
"""

import sys
import json
import os
import re
import shutil
import tempfile
import argparse
from pathlib import Path

try:
    from ccl_chromium_reader import ccl_chromium_indexeddb as idb
except ImportError:
    print(json.dumps({
        "error": "ccl_chromium_reader not installed. Run: "
                 "pip install git+https://github.com/cclgroupltd/ccl_chrome_indexeddb.git"
    }))
    sys.exit(1)


# ---------------------------------------------------------------------------
# Path discovery
# ---------------------------------------------------------------------------

_TEAMS_LDB_NAME = "https_teams.microsoft.com_0.indexeddb.leveldb"


def _win_appdata() -> tuple[Path, Path]:
    """Return (APPDATA, LOCALAPPDATA) on Windows or via WSL /mnt/c."""
    ad  = os.environ.get("APPDATA", "")
    lad = os.environ.get("LOCALAPPDATA", "")
    if ad and lad:
        return Path(ad), Path(lad)

    # WSL: probe /mnt/c/Users for a user that has AppData
    users = Path("/mnt/c/Users")
    if users.exists():
        skip = {"Public", "Default", "Default User", "All Users", "desktop.ini"}
        for d in sorted(users.iterdir()):
            if d.name in skip or not d.is_dir():
                continue
            try:
                ad_  = d / "AppData" / "Roaming"
                lad_ = d / "AppData" / "Local"
                if ad_.exists() and lad_.exists():
                    return ad_, lad_
            except PermissionError:
                continue

    return Path(""), Path("")


def find_teams_idb() -> Path | None:
    """Return the Teams IndexedDB .leveldb folder, preferring new Teams."""
    appdata, localappdata = _win_appdata()

    candidates = [
        # New Teams 2.x — Store/MSIX install (WV2Profile_tfw variant, 2024+)
        localappdata / "Packages" / "MSTeams_8wekyb3d8bbwe"
            / "LocalCache" / "Microsoft" / "MSTeams"
            / "EBWebView" / "WV2Profile_tfw" / "IndexedDB" / _TEAMS_LDB_NAME,
        # New Teams 2.x — Default profile variant
        localappdata / "Packages" / "MSTeams_8wekyb3d8bbwe"
            / "LocalCache" / "Microsoft" / "MSTeams"
            / "EBWebView" / "Default" / "IndexedDB" / _TEAMS_LDB_NAME,
        # New Teams 2.x — per-user installer (non-Store)
        localappdata / "Microsoft" / "Teams"
            / "EBWebView" / "Default" / "IndexedDB" / _TEAMS_LDB_NAME,
        # Classic Teams 1.x
        appdata / "Microsoft" / "Teams" / "IndexedDB" / _TEAMS_LDB_NAME,
    ]

    for p in candidates:
        if p and p.exists():
            return p

    return None


# ---------------------------------------------------------------------------
# Open IDB (copy to temp to avoid lock)
# ---------------------------------------------------------------------------

def open_idb(ldb_path: Path) -> tuple:
    """
    Copy LevelDB (and blob dir if present) to a temp folder and open.
    The original is locked by Teams; the copy is not.
    Returns (WrappedIndexDB, tmp_root). Caller must shutil.rmtree(tmp_root).
    """
    tmp_root = Path(tempfile.mkdtemp(prefix="teams_idb_"))
    dst_ldb = tmp_root / ldb_path.name
    shutil.copytree(ldb_path, dst_ldb, ignore=shutil.ignore_patterns("LOCK"))

    # blob dir lives next to the .leveldb folder with the same stem + .blob
    blob_src = ldb_path.parent / ldb_path.name.replace(".leveldb", ".blob")
    dst_blob = None
    if blob_src.exists():
        dst_blob = tmp_root / blob_src.name
        shutil.copytree(blob_src, dst_blob)

    return idb.WrappedIndexDB(dst_ldb, dst_blob), tmp_root


# ---------------------------------------------------------------------------
# Helpers to navigate the multi-database IDB structure
# ---------------------------------------------------------------------------

def _get_wrapped_db(widb, name_fragment: str):
    """
    Return the first WrappedDatabase whose name contains name_fragment.
    Uses 'Teams:<fragment>:' as the match prefix to avoid false positives
    like 'streams-replychain-manager' matching 'replychain-manager'.
    """
    precise = f"Teams:{name_fragment}:"
    for db_id in widb.database_ids:
        if precise in db_id.name:
            return idb.WrappedDatabase(widb._raw_db, db_id)
    # fallback to loose match
    for db_id in widb.database_ids:
        if name_fragment in db_id.name:
            return idb.WrappedDatabase(widb._raw_db, db_id)
    return None


def _iter_store(wrapped_db, store_name: str):
    """Yield record values from a named object store, silently skipping errors."""
    def _noop(k, v): pass
    try:
        store = wrapped_db.get_object_store_by_name(store_name)
        if store is None:
            return
        for rec in store.iterate_records(live_only=True,
                                         bad_deserializer_data_handler=_noop):
            try:
                val = rec.value
                if isinstance(val, dict):
                    yield val
            except Exception:
                continue
    except Exception:
        return


def _to_str(val) -> str:
    """Decode bytes to str. V8 stores one-byte strings as Latin-1."""
    if isinstance(val, bytes):
        try:
            return val.decode("utf-8")
        except UnicodeDecodeError:
            return val.decode("latin-1")
    return str(val) if val is not None else ""


def _strip_html(text) -> str:
    return re.sub(r"<[^>]+>", "", _to_str(text)).strip()


def _sender(from_field) -> str:
    if not from_field:
        return "Unknown"
    if isinstance(from_field, bytes):
        from_field = from_field.decode("utf-8", errors="replace")
    if isinstance(from_field, dict):
        user = from_field.get("user", from_field)
        return (_to_str(user.get("displayName") or user.get("name")
                or user.get("id", "Unknown")))
    raw = _to_str(from_field)
    # "8:orgid:<guid>" or "orgid:<guid>" → strip prefix noise
    parts = raw.split(":")
    return parts[-1] if len(parts) > 1 else raw


def _members_from_conv(conv: dict) -> list[str]:
    members = conv.get("members") or []
    names = []
    for m in members[:8]:
        if not isinstance(m, dict):
            continue
        name = (_to_str(m.get("nameHint") or m.get("displayName") or m.get("friendlyName") or ""))
        if not name:
            # strip "8:orgid:" prefix from id
            raw_id = _to_str(m.get("id") or m.get("mri") or "")
            name = raw_id.split(":")[-1] if ":" in raw_id else raw_id
        if name:
            names.append(name)
    return names


def _display_name(conv: dict) -> str:
    tp = conv.get("threadProperties") or {}
    return _to_str(tp.get("topic") or conv.get("displayName") or conv.get("id", ""))


# ---------------------------------------------------------------------------
# Actions
# ---------------------------------------------------------------------------

def action_list_dbs(widb) -> dict:
    return {"data": [db_id.name for db_id in widb.database_ids]}


def action_get_chats(widb, count: int) -> dict:
    db = _get_wrapped_db(widb, "conversation-manager")
    if db is None:
        return {"error": "conversation-manager database not found"}

    seen = set()
    chats = []
    for conv in _iter_store(db, "conversations"):
        cid = conv.get("id", "")
        if not cid or cid in seen:
            continue
        ttype = conv.get("threadType", "")
        if ttype in ("streamofnotifications",):
            continue
        seen.add(cid)
        chats.append({
            "id": cid,
            "displayName": _display_name(conv),
            "type": conv.get("type", ""),
            "threadType": ttype,
            "members": _members_from_conv(conv),
        })
        if len(chats) >= count:
            break

    return {"data": chats}


def action_get_messages(widb, chat_id: str, count: int) -> dict:
    db = _get_wrapped_db(widb, "replychain-manager")
    if db is None:
        return {"error": "replychain-manager database not found"}

    messages = []
    for chain in _iter_store(db, "replychains"):
        if chain.get("conversationId") != chat_id:
            continue
        msg_map = chain.get("messageMap") or {}
        for msg in msg_map.values():
            if not isinstance(msg, dict):
                continue
            mtype = msg.get("messageType", "")
            if mtype not in ("RichText/Html", "Text", ""):
                continue
            content = _strip_html(msg.get("content", ""))
            if not content:
                continue
            messages.append({
                "id": str(msg.get("id", "")),
                "content": content,
                "from": _sender(msg.get("from") or msg.get("imDisplayName")),
                "time": msg.get("originalArrivalTime") or msg.get("clientArrivalTime", ""),
                "chat_id": chat_id,
                "type": mtype,
            })
            if len(messages) >= count:
                break
        if len(messages) >= count:
            break

    messages.sort(key=lambda m: m["time"])
    return {"data": messages}


def action_search_messages(widb, query: str, count: int) -> dict:
    db = _get_wrapped_db(widb, "replychain-manager")
    if db is None:
        return {"error": "replychain-manager database not found"}

    q = query.lower()
    results = []
    for chain in _iter_store(db, "replychains"):
        msg_map = chain.get("messageMap") or {}
        conv_id = chain.get("conversationId", "")
        for msg in msg_map.values():
            if not isinstance(msg, dict):
                continue
            content = _strip_html(msg.get("content", ""))
            if not content or q not in content.lower():
                continue
            results.append({
                "id": str(msg.get("id", "")),
                "content": content,
                "from": _sender(msg.get("from") or msg.get("imDisplayName")),
                "time": msg.get("originalArrivalTime") or msg.get("clientArrivalTime", ""),
                "chat_id": conv_id,
            })
            if len(results) >= count:
                break
        if len(results) >= count:
            break

    return {"data": results}


def action_get_channels(widb, count: int) -> dict:
    db = _get_wrapped_db(widb, "conversation-manager")
    if db is None:
        return {"error": "conversation-manager database not found"}

    # Classic Teams: channels have type=Topic or threadType in General/Regular/channel
    _CHANNEL_TYPES = {"General", "Regular", "channel", "Topic"}
    seen = set()
    channels = []
    for conv in _iter_store(db, "conversations"):
        ctype = _to_str(conv.get("type", ""))
        ttype = _to_str(conv.get("threadType", ""))
        if ctype not in _CHANNEL_TYPES and ttype not in _CHANNEL_TYPES:
            continue
        cid = _to_str(conv.get("id", ""))
        if not cid or cid in seen:
            continue
        seen.add(cid)
        channels.append({
            "id": cid,
            "channelName": _display_name(conv),
            "teamId": _to_str(conv.get("teamId", "")),
            "type": ctype or ttype,
        })
        if len(channels) >= count:
            break

    return {"data": channels}


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Read Teams local IDB cache")
    parser.add_argument("--action", required=True,
                        choices=["get_chats", "get_messages", "search_messages",
                                 "get_channels", "list_stores"])
    parser.add_argument("--count",   type=int,  default=20)
    parser.add_argument("--chat_id", default=None)
    parser.add_argument("--query",   default=None)
    parser.add_argument("--idb_path", default=None,
                        help="Override auto-detected LevelDB path")
    args = parser.parse_args()

    ldb_path = Path(args.idb_path) if args.idb_path else find_teams_idb()

    if not ldb_path or not ldb_path.exists():
        print(json.dumps({
            "error": (
                "Teams IndexedDB not found. "
                "Teams must be installed and launched at least once. "
                "Use --idb_path to override."
            )
        }))
        sys.exit(1)

    tmp_root = None
    try:
        widb, tmp_root = open_idb(ldb_path)
    except Exception as e:
        print(json.dumps({"error": f"Cannot open Teams database: {e}"}))
        sys.exit(1)

    try:
        if args.action == "list_stores":
            result = action_list_dbs(widb)
        elif args.action == "get_chats":
            result = action_get_chats(widb, args.count)
        elif args.action == "get_messages":
            if not args.chat_id:
                result = {"error": "--chat_id is required"}
            else:
                result = action_get_messages(widb, args.chat_id, args.count)
        elif args.action == "search_messages":
            if not args.query:
                result = {"error": "--query is required"}
            else:
                result = action_search_messages(widb, args.query, args.count)
        elif args.action == "get_channels":
            result = action_get_channels(widb, args.count)
        else:
            result = {"error": f"Unknown action: {args.action}"}
    except Exception as e:
        result = {"error": str(e)}
    finally:
        if tmp_root:
            shutil.rmtree(tmp_root, ignore_errors=True)

    print(json.dumps(result, ensure_ascii=False, default=str))


if __name__ == "__main__":
    main()
