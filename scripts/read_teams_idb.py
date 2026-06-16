#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Read Microsoft Teams local IndexedDB cache (read-only, no API needed).
Requires: pip install git+https://github.com/cclgroupltd/ccl_chrome_indexeddb.git

Works while Teams is running — copies LevelDB to a temp folder before reading.

Performance model
-----------------
Deserializing the Teams LevelDB (V8 records, ~50 MB) costs ~25 s of CPU per
run. Doing that on every tool call made parallel queries time out. Instead we
parse the whole reply-chain + conversation stores ONCE into a small on-disk
SQLite index and serve every subsequent query from it in milliseconds. The
index is reused for `TEAMS_CACHE_TTL` seconds (default 300) and rebuilt under a
lock so concurrent cold calls only pay the parse cost once. Pass --refresh to
force a rebuild.
"""

import sys
import json
import os
import re
import time
import shutil
import sqlite3
import hashlib
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
    Copy LevelDB to a temp folder and open. The original is locked by Teams;
    the copy is not. The .blob sidecar holds attachments/media only (no message
    text), so we skip it — saves a copy and is irrelevant to the text stores.
    Returns (WrappedIndexDB, tmp_root). Caller must shutil.rmtree(tmp_root).
    """
    tmp_root = Path(tempfile.mkdtemp(prefix="teams_idb_"))
    dst_ldb = tmp_root / ldb_path.name
    shutil.copytree(ldb_path, dst_ldb, ignore=shutil.ignore_patterns("LOCK"))
    return idb.WrappedIndexDB(dst_ldb, None), tmp_root


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


def _msg_time(msg: dict):
    """Arrival time as a number when possible (ms epoch), else raw/empty."""
    t = msg.get("originalArrivalTime")
    if t is None:
        t = msg.get("clientArrivalTime", "")
    if isinstance(t, (int, float)):
        return t
    s = _to_str(t)
    try:
        return float(s)
    except (ValueError, TypeError):
        return s


# ---------------------------------------------------------------------------
# SQLite index cache
# ---------------------------------------------------------------------------

_DEFAULT_TTL = int(os.environ.get("TEAMS_CACHE_TTL", "300"))
_STALE_LOCK_S = 120        # steal a build lock older than this (build is ~25s)
_BUILD_WAIT_S = 240        # max time to wait for another process's build

_SCHEMA = """
CREATE TABLE IF NOT EXISTS messages (
    id          TEXT PRIMARY KEY,
    chat_id     TEXT,
    sender      TEXT,
    content     TEXT,
    content_lc  TEXT,
    time        REAL,
    type        TEXT
);
CREATE INDEX IF NOT EXISTS idx_messages_chat ON messages(chat_id, time);
CREATE INDEX IF NOT EXISTS idx_messages_lc   ON messages(content_lc);
CREATE TABLE IF NOT EXISTS conversations (
    id           TEXT PRIMARY KEY,
    display_name TEXT,
    type         TEXT,
    thread_type  TEXT,
    team_id      TEXT,
    members      TEXT
);
"""


def _cache_dir(ldb_path: Path) -> Path:
    key = hashlib.md5(str(ldb_path).encode("utf-8")).hexdigest()[:12]
    return Path(tempfile.gettempdir()) / f"teams_idb_cache_{key}"


def _cache_is_fresh(db_path: Path, stamp_path: Path, ttl: int) -> bool:
    if not db_path.exists():
        return False
    try:
        built_at = float(stamp_path.read_text().strip())
    except (OSError, ValueError):
        return False
    return (time.time() - built_at) < ttl


def _build_index(ldb_path: Path, db_path: Path) -> None:
    """Parse the whole IDB once and write a fresh SQLite index atomically."""
    widb, tmp_root = open_idb(ldb_path)
    tmp_db = db_path.with_name(db_path.name + f".tmp-{os.getpid()}")
    try:
        if tmp_db.exists():
            tmp_db.unlink()
        con = sqlite3.connect(str(tmp_db))
        try:
            con.executescript(_SCHEMA)

            rc_db = _get_wrapped_db(widb, "replychain-manager")
            if rc_db is not None:
                msg_rows = []
                for chain in _iter_store(rc_db, "replychains"):
                    conv_id = _to_str(chain.get("conversationId", ""))
                    for msg in (chain.get("messageMap") or {}).values():
                        if not isinstance(msg, dict):
                            continue
                        mtype = msg.get("messageType", "")
                        if mtype not in ("RichText/Html", "Text", ""):
                            continue
                        content = _strip_html(msg.get("content", ""))
                        if not content:
                            continue
                        msg_rows.append((
                            str(msg.get("id", "")),
                            conv_id,
                            _sender(msg.get("from") or msg.get("imDisplayName")),
                            content,
                            content.lower(),
                            _msg_time(msg),
                            _to_str(mtype),
                        ))
                con.executemany(
                    "INSERT OR REPLACE INTO messages "
                    "(id, chat_id, sender, content, content_lc, time, type) "
                    "VALUES (?,?,?,?,?,?,?)", msg_rows)

            cm_db = _get_wrapped_db(widb, "conversation-manager")
            if cm_db is not None:
                conv_rows = []
                for conv in _iter_store(cm_db, "conversations"):
                    cid = _to_str(conv.get("id", ""))
                    if not cid:
                        continue
                    conv_rows.append((
                        cid,
                        _display_name(conv),
                        _to_str(conv.get("type", "")),
                        _to_str(conv.get("threadType", "")),
                        _to_str(conv.get("teamId", "")),
                        json.dumps(_members_from_conv(conv), ensure_ascii=False),
                    ))
                con.executemany(
                    "INSERT OR REPLACE INTO conversations "
                    "(id, display_name, type, thread_type, team_id, members) "
                    "VALUES (?,?,?,?,?,?)", conv_rows)

            con.commit()
        finally:
            con.close()
        os.replace(str(tmp_db), str(db_path))
    finally:
        if tmp_db.exists():
            tmp_db.unlink()
        shutil.rmtree(tmp_root, ignore_errors=True)


def get_index(ldb_path: Path, ttl: int, refresh: bool) -> Path:
    """
    Return a path to a fresh SQLite index, building it under a lock if needed.
    Concurrent cold callers wait for the in-progress build instead of each
    re-parsing the 50 MB LevelDB.
    """
    cache_dir = _cache_dir(ldb_path)
    cache_dir.mkdir(parents=True, exist_ok=True)
    db_path = cache_dir / "index.sqlite"
    stamp_path = cache_dir / "built_at"
    lock_path = cache_dir / "build.lock"

    if not refresh and _cache_is_fresh(db_path, stamp_path, ttl):
        return db_path

    # Try to become the builder.
    while True:
        try:
            fd = os.open(str(lock_path), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.write(fd, str(os.getpid()).encode())
            os.close(fd)
            break  # we hold the lock
        except FileExistsError:
            # Someone else is building. Wait for their result.
            waited = 0.0
            while waited < _BUILD_WAIT_S:
                if _cache_is_fresh(db_path, stamp_path, ttl):
                    return db_path
                # Steal a stale/abandoned lock.
                try:
                    if (time.time() - lock_path.stat().st_mtime) > _STALE_LOCK_S:
                        lock_path.unlink(missing_ok=True)
                        break
                except OSError:
                    break  # lock vanished — retry acquire
                time.sleep(0.3)
                waited += 0.3
            else:
                # Waited too long; fall back to building ourselves.
                lock_path.unlink(missing_ok=True)
            continue  # retry acquire

    try:
        # Double-check: another builder may have finished between our checks.
        if refresh or not _cache_is_fresh(db_path, stamp_path, ttl):
            _build_index(ldb_path, db_path)
            stamp_path.write_text(str(time.time()))
    finally:
        lock_path.unlink(missing_ok=True)

    return db_path


def _connect_ro(db_path: Path) -> sqlite3.Connection:
    con = sqlite3.connect(f"file:{db_path}?mode=ro", uri=True)
    con.row_factory = sqlite3.Row
    return con


def _like_escape(s: str) -> str:
    return s.replace("\\", "\\\\").replace("%", "\\%").replace("_", "\\_")


# ---------------------------------------------------------------------------
# Actions (served from the SQLite index)
# ---------------------------------------------------------------------------

def action_get_chats(db_path: Path, count: int) -> dict:
    con = _connect_ro(db_path)
    try:
        rows = con.execute(
            "SELECT id, display_name, type, thread_type, members "
            "FROM conversations "
            "WHERE thread_type != 'streamofnotifications' "
            "LIMIT ?", (count,)).fetchall()
    finally:
        con.close()
    chats = []
    for r in rows:
        try:
            members = json.loads(r["members"]) if r["members"] else []
        except (ValueError, TypeError):
            members = []
        chats.append({
            "id": r["id"],
            "displayName": r["display_name"],
            "type": r["type"],
            "threadType": r["thread_type"],
            "members": members,
        })
    return {"data": chats}


def action_get_messages(db_path: Path, chat_id: str, count: int) -> dict:
    con = _connect_ro(db_path)
    try:
        rows = con.execute(
            "SELECT id, chat_id, sender, content, time, type "
            "FROM messages WHERE chat_id = ? "
            "ORDER BY time ASC LIMIT ?", (chat_id, count)).fetchall()
    finally:
        con.close()
    messages = [{
        "id": r["id"],
        "content": r["content"],
        "from": r["sender"],
        "time": r["time"],
        "chat_id": r["chat_id"],
        "type": r["type"],
    } for r in rows]
    return {"data": messages}


def action_search_messages(db_path: Path, query: str, count: int) -> dict:
    pattern = f"%{_like_escape(query.lower())}%"
    con = _connect_ro(db_path)
    try:
        rows = con.execute(
            "SELECT id, chat_id, sender, content, time "
            "FROM messages WHERE content_lc LIKE ? ESCAPE '\\' "
            "ORDER BY time DESC LIMIT ?", (pattern, count)).fetchall()
    finally:
        con.close()
    results = [{
        "id": r["id"],
        "content": r["content"],
        "from": r["sender"],
        "time": r["time"],
        "chat_id": r["chat_id"],
    } for r in rows]
    return {"data": results}


_CHANNEL_TYPES = {"General", "Regular", "channel", "Topic"}


def action_get_channels(db_path: Path, count: int) -> dict:
    con = _connect_ro(db_path)
    try:
        rows = con.execute(
            "SELECT id, display_name, type, thread_type, team_id "
            "FROM conversations").fetchall()
    finally:
        con.close()
    channels = []
    for r in rows:
        ctype, ttype = r["type"], r["thread_type"]
        if ctype not in _CHANNEL_TYPES and ttype not in _CHANNEL_TYPES:
            continue
        channels.append({
            "id": r["id"],
            "channelName": r["display_name"],
            "teamId": r["team_id"],
            "type": ctype or ttype,
        })
        if len(channels) >= count:
            break
    return {"data": channels}


def action_list_dbs(ldb_path: Path) -> dict:
    """Debug: list raw object-store databases. Reads the IDB directly."""
    widb, tmp_root = open_idb(ldb_path)
    try:
        return {"data": [db_id.name for db_id in widb.database_ids]}
    finally:
        shutil.rmtree(tmp_root, ignore_errors=True)


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
    parser.add_argument("--ttl", type=int, default=_DEFAULT_TTL,
                        help="Reuse the SQLite index for this many seconds")
    parser.add_argument("--refresh", action="store_true",
                        help="Force a rebuild of the SQLite index")
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

    try:
        if args.action == "list_stores":
            result = action_list_dbs(ldb_path)
        else:
            db_path = get_index(ldb_path, args.ttl, args.refresh)
            if args.action == "get_chats":
                result = action_get_chats(db_path, args.count)
            elif args.action == "get_messages":
                if not args.chat_id:
                    result = {"error": "--chat_id is required"}
                else:
                    result = action_get_messages(db_path, args.chat_id, args.count)
            elif args.action == "search_messages":
                if not args.query:
                    result = {"error": "--query is required"}
                else:
                    result = action_search_messages(db_path, args.query, args.count)
            elif args.action == "get_channels":
                result = action_get_channels(db_path, args.count)
            else:
                result = {"error": f"Unknown action: {args.action}"}
    except Exception as e:
        result = {"error": str(e)}

    print(json.dumps(result, ensure_ascii=False, default=str))


if __name__ == "__main__":
    main()
