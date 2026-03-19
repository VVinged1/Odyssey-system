import json
import sqlite3
from datetime import datetime, timezone
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any
from urllib.parse import parse_qs, urlparse


BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "bridge_state.db"
HOST = "127.0.0.1"
PORT = 8765


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat()


def json_dumps(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"))


class BridgeDatabase:
    def __init__(self, path: Path):
        self.path = path
        self.connection = sqlite3.connect(path, check_same_thread=False)
        self.connection.row_factory = sqlite3.Row
        self._initialize()

    def _initialize(self) -> None:
        self.connection.executescript(
            """
            PRAGMA journal_mode=WAL;

            CREATE TABLE IF NOT EXISTS token_states (
                room_id TEXT NOT NULL,
                token_id TEXT NOT NULL,
                token_name TEXT,
                player_id TEXT,
                character_id TEXT,
                tracker_data TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                PRIMARY KEY (room_id, token_id)
            );

            CREATE TABLE IF NOT EXISTS roll_events (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                room_id TEXT NOT NULL,
                token_id TEXT NOT NULL,
                actor_name TEXT,
                summary TEXT NOT NULL,
                payload TEXT NOT NULL,
                created_at TEXT NOT NULL,
                applied_at TEXT
            );

            CREATE INDEX IF NOT EXISTS idx_roll_events_room_id_id
            ON roll_events (room_id, id);
            """
        )
        self.connection.commit()

    def upsert_tokens(self, room_id: str, tokens: list[dict[str, Any]]) -> list[dict[str, Any]]:
        now = utc_now()
        saved: list[dict[str, Any]] = []
        with self.connection:
            for token in tokens:
                token_id = str(token["token_id"])
                record = {
                    "room_id": room_id,
                    "token_id": token_id,
                    "token_name": token.get("token_name") or "",
                    "player_id": token.get("player_id") or "",
                    "character_id": token.get("character_id") or "",
                    "tracker_data": json_dumps(token.get("tracker_data") or {}),
                    "updated_at": now,
                }
                self.connection.execute(
                    """
                    INSERT INTO token_states (
                        room_id, token_id, token_name, player_id, character_id, tracker_data, updated_at
                    ) VALUES (
                        :room_id, :token_id, :token_name, :player_id, :character_id, :tracker_data, :updated_at
                    )
                    ON CONFLICT(room_id, token_id) DO UPDATE SET
                        token_name = excluded.token_name,
                        player_id = excluded.player_id,
                        character_id = excluded.character_id,
                        tracker_data = excluded.tracker_data,
                        updated_at = excluded.updated_at
                    """,
                    record,
                )
                saved.append(
                    {
                        "token_id": token_id,
                        "updated_at": now,
                    }
                )
        return saved

    def list_tokens(self, room_id: str) -> list[dict[str, Any]]:
        rows = self.connection.execute(
            """
            SELECT room_id, token_id, token_name, player_id, character_id, tracker_data, updated_at
            FROM token_states
            WHERE room_id = ?
            ORDER BY token_name COLLATE NOCASE, token_id
            """,
            (room_id,),
        ).fetchall()
        return [self._token_row_to_dict(row) for row in rows]

    def create_roll_event(self, room_id: str, event: dict[str, Any]) -> dict[str, Any]:
        token_id = str(event["token_id"])
        actor_name = event.get("actor_name") or ""
        summary = event.get("summary") or "Roll resolved"
        payload = event.get("payload") or {}
        created_at = utc_now()
        with self.connection:
            cursor = self.connection.execute(
                """
                INSERT INTO roll_events (room_id, token_id, actor_name, summary, payload, created_at)
                VALUES (?, ?, ?, ?, ?, ?)
                """,
                (room_id, token_id, actor_name, summary, json_dumps(payload), created_at),
            )
        return {
            "id": cursor.lastrowid,
            "room_id": room_id,
            "token_id": token_id,
            "actor_name": actor_name,
            "summary": summary,
            "payload": payload,
            "created_at": created_at,
        }

    def list_roll_events(self, room_id: str, since_id: int) -> list[dict[str, Any]]:
        rows = self.connection.execute(
            """
            SELECT id, room_id, token_id, actor_name, summary, payload, created_at, applied_at
            FROM roll_events
            WHERE room_id = ? AND id > ?
            ORDER BY id ASC
            """,
            (room_id, since_id),
        ).fetchall()
        return [self._event_row_to_dict(row) for row in rows]

    def mark_events_applied(self, room_id: str, event_ids: list[int]) -> int:
        if not event_ids:
            return 0
        applied_at = utc_now()
        placeholders = ",".join("?" for _ in event_ids)
        params = [applied_at, room_id, *event_ids]
        with self.connection:
            cursor = self.connection.execute(
                f"""
                UPDATE roll_events
                SET applied_at = ?
                WHERE room_id = ? AND id IN ({placeholders})
                """,
                params,
            )
        return cursor.rowcount

    def _token_row_to_dict(self, row: sqlite3.Row) -> dict[str, Any]:
        return {
            "room_id": row["room_id"],
            "token_id": row["token_id"],
            "token_name": row["token_name"],
            "player_id": row["player_id"],
            "character_id": row["character_id"],
            "tracker_data": json.loads(row["tracker_data"]),
            "updated_at": row["updated_at"],
        }

    def _event_row_to_dict(self, row: sqlite3.Row) -> dict[str, Any]:
        return {
            "id": row["id"],
            "room_id": row["room_id"],
            "token_id": row["token_id"],
            "actor_name": row["actor_name"],
            "summary": row["summary"],
            "payload": json.loads(row["payload"]),
            "created_at": row["created_at"],
            "applied_at": row["applied_at"],
        }


database = BridgeDatabase(DB_PATH)


class BridgeHandler(BaseHTTPRequestHandler):
    server_version = "OdysseyBridge/0.1"

    def do_OPTIONS(self) -> None:
        self.send_response(HTTPStatus.NO_CONTENT)
        self._send_cors_headers()
        self.end_headers()

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/health":
            self._send_json({"ok": True, "service": "odyssey-bridge", "time": utc_now()})
            return

        if parsed.path == "/api/tokens":
            query = parse_qs(parsed.query)
            room_id = (query.get("room_id") or [""])[0].strip()
            if not room_id:
                self._send_error(HTTPStatus.BAD_REQUEST, "room_id is required")
                return
            self._send_json({"items": database.list_tokens(room_id)})
            return

        if parsed.path == "/api/events":
            query = parse_qs(parsed.query)
            room_id = (query.get("room_id") or [""])[0].strip()
            since_id_raw = (query.get("since_id") or ["0"])[0].strip()
            if not room_id:
                self._send_error(HTTPStatus.BAD_REQUEST, "room_id is required")
                return
            try:
                since_id = int(since_id_raw or "0")
            except ValueError:
                self._send_error(HTTPStatus.BAD_REQUEST, "since_id must be an integer")
                return
            items = database.list_roll_events(room_id, since_id)
            self._send_json({"items": items})
            return

        self._send_error(HTTPStatus.NOT_FOUND, "Route not found")

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        payload = self._read_json()
        if payload is None:
            return

        if parsed.path == "/api/tokens/upsert":
            room_id = str(payload.get("room_id") or "").strip()
            tokens = payload.get("tokens") or []
            if not room_id:
                self._send_error(HTTPStatus.BAD_REQUEST, "room_id is required")
                return
            if not isinstance(tokens, list) or not tokens:
                self._send_error(HTTPStatus.BAD_REQUEST, "tokens must be a non-empty list")
                return
            for token in tokens:
                if not isinstance(token, dict) or not token.get("token_id"):
                    self._send_error(HTTPStatus.BAD_REQUEST, "each token must include token_id")
                    return
            saved = database.upsert_tokens(room_id, tokens)
            self._send_json({"saved": saved}, HTTPStatus.CREATED)
            return

        if parsed.path == "/api/rolls":
            room_id = str(payload.get("room_id") or "").strip()
            event = payload.get("event")
            if not room_id:
                self._send_error(HTTPStatus.BAD_REQUEST, "room_id is required")
                return
            if not isinstance(event, dict) or not event.get("token_id"):
                self._send_error(HTTPStatus.BAD_REQUEST, "event.token_id is required")
                return
            created = database.create_roll_event(room_id, event)
            self._send_json(created, HTTPStatus.CREATED)
            return

        if parsed.path == "/api/events/ack":
            room_id = str(payload.get("room_id") or "").strip()
            event_ids = payload.get("event_ids") or []
            if not room_id:
                self._send_error(HTTPStatus.BAD_REQUEST, "room_id is required")
                return
            if not isinstance(event_ids, list) or not all(isinstance(item, int) for item in event_ids):
                self._send_error(HTTPStatus.BAD_REQUEST, "event_ids must be a list of integers")
                return
            updated = database.mark_events_applied(room_id, event_ids)
            self._send_json({"updated": updated})
            return

        self._send_error(HTTPStatus.NOT_FOUND, "Route not found")

    def log_message(self, format: str, *args: Any) -> None:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        print(f"[{timestamp}] {self.address_string()} {format % args}")

    def _read_json(self) -> dict[str, Any] | None:
        content_length = int(self.headers.get("Content-Length", "0"))
        raw_body = self.rfile.read(content_length) if content_length > 0 else b"{}"
        try:
            return json.loads(raw_body.decode("utf-8") or "{}")
        except json.JSONDecodeError:
            self._send_error(HTTPStatus.BAD_REQUEST, "invalid JSON body")
            return None

    def _send_json(self, payload: dict[str, Any], status: HTTPStatus = HTTPStatus.OK) -> None:
        body = json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8")
        self.send_response(status)
        self._send_cors_headers()
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _send_error(self, status: HTTPStatus, message: str) -> None:
        self._send_json({"error": message, "status": int(status)}, status)

    def _send_cors_headers(self) -> None:
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET,POST,OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")


def main() -> None:
    print(f"Starting Odyssey bridge on http://{HOST}:{PORT}")
    print(f"SQLite database: {DB_PATH}")
    server = ThreadingHTTPServer((HOST, PORT), BridgeHandler)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("Stopping bridge server")
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
