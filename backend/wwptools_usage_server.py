#!/usr/bin/env python3
import argparse
import csv
import io
import json
import os
import sqlite3
from datetime import datetime, timedelta, timezone
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlparse


def utc_now():
    return datetime.now(timezone.utc)


def utc_now_text():
    return utc_now().strftime("%Y-%m-%dT%H:%M:%SZ")


def parse_utc_timestamp(value):
    text = str(value or "").strip()
    if not text:
        return utc_now()
    try:
        if text.endswith("Z"):
            text = text[:-1] + "+00:00"
        return datetime.fromisoformat(text).astimezone(timezone.utc)
    except Exception:
        return utc_now()


def iso_week_start_text(value):
    current = parse_utc_timestamp(value)
    week_start = (current - timedelta(days=current.weekday())).date()
    return week_start.isoformat()


class TelemetryStore(object):
    def __init__(self, db_path):
        self.db_path = os.path.abspath(db_path)
        db_dir = os.path.dirname(self.db_path)
        if db_dir:
            os.makedirs(db_dir, exist_ok=True)
        self._initialize()

    def _connect(self):
        connection = sqlite3.connect(self.db_path)
        connection.row_factory = sqlite3.Row
        return connection

    def _initialize(self):
        with self._connect() as connection:
            connection.execute(
                """
                CREATE TABLE IF NOT EXISTS telemetry_events (
                    event_id TEXT PRIMARY KEY,
                    event_type TEXT NOT NULL,
                    timestamp_utc TEXT NOT NULL,
                    week_start TEXT NOT NULL,
                    extension_version TEXT,
                    command_name TEXT,
                    command_bundle TEXT,
                    command_extension TEXT,
                    tool_key TEXT,
                    install_id TEXT,
                    user_id TEXT,
                    machine_id TEXT,
                    revit_version TEXT,
                    payload_json TEXT NOT NULL,
                    received_utc TEXT NOT NULL
                )
                """
            )
            connection.execute(
                "CREATE INDEX IF NOT EXISTS idx_events_timestamp ON telemetry_events(timestamp_utc)"
            )
            connection.execute(
                "CREATE INDEX IF NOT EXISTS idx_events_week_tool ON telemetry_events(week_start, tool_key)"
            )

    def ingest_events(self, events):
        accepted = 0
        received_utc = utc_now_text()
        with self._connect() as connection:
            for event in events:
                if not isinstance(event, dict):
                    continue
                event_id = str(event.get("event_id") or "").strip()
                if not event_id:
                    continue
                timestamp_utc = parse_utc_timestamp(event.get("timestamp_utc")).strftime("%Y-%m-%dT%H:%M:%SZ")
                payload_json = json.dumps(event, sort_keys=True)
                cursor = connection.execute(
                    """
                    INSERT OR IGNORE INTO telemetry_events (
                        event_id,
                        event_type,
                        timestamp_utc,
                        week_start,
                        extension_version,
                        command_name,
                        command_bundle,
                        command_extension,
                        tool_key,
                        install_id,
                        user_id,
                        machine_id,
                        revit_version,
                        payload_json,
                        received_utc
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        event_id,
                        str(event.get("event_type") or "").strip(),
                        timestamp_utc,
                        iso_week_start_text(timestamp_utc),
                        str(event.get("extension_version") or "").strip(),
                        str(event.get("command_name") or "").strip(),
                        str(event.get("command_bundle") or "").strip(),
                        str(event.get("command_extension") or "").strip(),
                        str(event.get("tool_key") or "").strip(),
                        str(event.get("install_id") or "").strip(),
                        str(event.get("user_id") or "").strip(),
                        str(event.get("machine_id") or "").strip(),
                        str(event.get("revit_version") or "").strip(),
                        payload_json,
                        received_utc,
                    ),
                )
                accepted += 1 if cursor.rowcount and cursor.rowcount > 0 else 0
        return accepted

    def _cutoff_text(self, days):
        return (utc_now() - timedelta(days=max(1, int(days)))).strftime("%Y-%m-%dT%H:%M:%SZ")

    def summary(self, days=7):
        cutoff = self._cutoff_text(days)
        with self._connect() as connection:
            active_users = connection.execute(
                "SELECT COUNT(DISTINCT user_id) AS value FROM telemetry_events WHERE timestamp_utc >= ?",
                (cutoff,),
            ).fetchone()["value"]
            active_installs = connection.execute(
                "SELECT COUNT(DISTINCT install_id) AS value FROM telemetry_events WHERE timestamp_utc >= ?",
                (cutoff,),
            ).fetchone()["value"]
            tool_users = connection.execute(
                """
                SELECT COUNT(DISTINCT user_id) AS value
                FROM telemetry_events
                WHERE timestamp_utc >= ? AND event_type = 'command-exec'
                """,
                (cutoff,),
            ).fetchone()["value"]
            tool_runs = connection.execute(
                """
                SELECT COUNT(*) AS value
                FROM telemetry_events
                WHERE timestamp_utc >= ? AND event_type = 'command-exec'
                """,
                (cutoff,),
            ).fetchone()["value"]

            versions = [
                dict(row)
                for row in connection.execute(
                    """
                    SELECT extension_version, COUNT(*) AS events
                    FROM telemetry_events
                    WHERE timestamp_utc >= ?
                    GROUP BY extension_version
                    ORDER BY events DESC, extension_version DESC
                    """,
                    (cutoff,),
                ).fetchall()
            ]

            tools = [
                {
                    "tool_key": row["tool_key"],
                    "tool_name": row["tool_name"],
                    "runs": row["runs"],
                    "users": row["users"],
                    "installs": row["installs"],
                }
                for row in connection.execute(
                    """
                    SELECT
                        tool_key,
                        COALESCE(NULLIF(command_name, ''), NULLIF(tool_key, ''), '(unknown)') AS tool_name,
                        COUNT(*) AS runs,
                        COUNT(DISTINCT user_id) AS users,
                        COUNT(DISTINCT install_id) AS installs
                    FROM telemetry_events
                    WHERE timestamp_utc >= ? AND event_type = 'command-exec'
                    GROUP BY tool_key, tool_name
                    ORDER BY runs DESC, tool_name ASC
                    """,
                    (cutoff,),
                ).fetchall()
            ]

        return {
            "window_days": int(days),
            "cutoff_utc": cutoff,
            "active_users": int(active_users or 0),
            "active_installs": int(active_installs or 0),
            "tool_users": int(tool_users or 0),
            "tool_runs": int(tool_runs or 0),
            "versions": versions,
            "tools": tools,
        }

    def weekly_report(self, weeks=8):
        weeks = max(1, int(weeks))
        current_week = utc_now().date() - timedelta(days=utc_now().weekday())
        min_week = (current_week - timedelta(weeks=weeks - 1)).isoformat()
        with self._connect() as connection:
            rows = connection.execute(
                """
                SELECT
                    week_start,
                    tool_key,
                    COALESCE(NULLIF(command_name, ''), NULLIF(tool_key, ''), '(unknown)') AS tool_name,
                    COUNT(*) AS runs,
                    COUNT(DISTINCT user_id) AS users,
                    COUNT(DISTINCT install_id) AS installs
                FROM telemetry_events
                WHERE event_type = 'command-exec' AND week_start >= ?
                GROUP BY week_start, tool_key, tool_name
                ORDER BY week_start DESC, runs DESC, tool_name ASC
                """,
                (min_week,),
            ).fetchall()
        return [dict(row) for row in rows]


def render_summary_html(summary):
    tool_rows = []
    for tool in summary["tools"]:
        tool_rows.append(
            "<tr><td>{}</td><td>{}</td><td>{}</td><td>{}</td></tr>".format(
                escape_html(tool["tool_name"]),
                int(tool["runs"]),
                int(tool["users"]),
                int(tool["installs"]),
            )
        )

    version_rows = []
    for version in summary["versions"]:
        version_rows.append(
            "<tr><td>{}</td><td>{}</td></tr>".format(
                escape_html(version.get("extension_version") or "(unknown)"),
                int(version.get("events") or 0),
            )
        )

    return """<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <title>WWPTools Usage Summary</title>
  <style>
    body {{ font-family: Segoe UI, Arial, sans-serif; margin: 24px; color: #1f2937; }}
    .cards {{ display: flex; gap: 16px; flex-wrap: wrap; margin-bottom: 24px; }}
    .card {{ border: 1px solid #d7dee6; border-radius: 8px; padding: 16px; min-width: 180px; }}
    .value {{ font-size: 28px; font-weight: 700; margin-top: 6px; }}
    table {{ border-collapse: collapse; width: 100%; margin-top: 12px; }}
    th, td {{ border: 1px solid #d7dee6; padding: 8px 10px; text-align: left; }}
    th {{ background: #f3f5f7; }}
    h1, h2 {{ margin-bottom: 8px; }}
    .muted {{ color: #6b7280; }}
  </style>
</head>
<body>
  <h1>WWPTools Usage Summary</h1>
  <div class="muted">Rolling {window_days}-day window ending now</div>
  <div class="cards">
    <div class="card"><div>Active users</div><div class="value">{active_users}</div></div>
    <div class="card"><div>Active installs</div><div class="value">{active_installs}</div></div>
    <div class="card"><div>Tool users</div><div class="value">{tool_users}</div></div>
    <div class="card"><div>Tool runs</div><div class="value">{tool_runs}</div></div>
  </div>
  <h2>Tool Usage</h2>
  <table>
    <thead><tr><th>Tool</th><th>Runs</th><th>Users</th><th>Installs</th></tr></thead>
    <tbody>{tool_rows}</tbody>
  </table>
  <h2>Versions Seen</h2>
  <table>
    <thead><tr><th>Version</th><th>Events</th></tr></thead>
    <tbody>{version_rows}</tbody>
  </table>
</body>
</html>""".format(
        window_days=summary["window_days"],
        active_users=summary["active_users"],
        active_installs=summary["active_installs"],
        tool_users=summary["tool_users"],
        tool_runs=summary["tool_runs"],
        tool_rows="".join(tool_rows) or "<tr><td colspan='4'>No tool usage recorded.</td></tr>",
        version_rows="".join(version_rows) or "<tr><td colspan='2'>No version data recorded.</td></tr>",
    )


def render_weekly_markdown(rows):
    lines = [
        "# WWPTools Weekly Tool Report",
        "",
        "| Week Start | Tool | Runs | Users | Installs |",
        "| --- | --- | ---: | ---: | ---: |",
    ]
    for row in rows:
        lines.append(
            "| {week_start} | {tool_name} | {runs} | {users} | {installs} |".format(
                week_start=row["week_start"],
                tool_name=row["tool_name"],
                runs=row["runs"],
                users=row["users"],
                installs=row["installs"],
            )
        )
    if len(lines) == 4:
        lines.append("| - | No command usage recorded | 0 | 0 | 0 |")
    return "\n".join(lines) + "\n"


def render_weekly_csv(rows):
    stream = io.StringIO()
    writer = csv.writer(stream)
    writer.writerow(["week_start", "tool_key", "tool_name", "runs", "users", "installs"])
    for row in rows:
        writer.writerow(
            [
                row["week_start"],
                row["tool_key"],
                row["tool_name"],
                row["runs"],
                row["users"],
                row["installs"],
            ]
        )
    return stream.getvalue()


def escape_html(value):
    return (
        str(value or "")
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def require_key(handler, api_key):
    if not api_key:
        return True

    parsed = urlparse(handler.path)
    query = parse_qs(parsed.query)
    provided = handler.headers.get("X-WWPTools-Key") or ""
    if not provided:
        values = query.get("key") or []
        if values:
            provided = values[0]
    return provided == api_key


def build_handler(store, api_key):
    class UsageHandler(BaseHTTPRequestHandler):
        def _send(self, status, body, content_type="application/json"):
            if isinstance(body, str):
                body = body.encode("utf-8")
            self.send_response(status)
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)

        def _json(self, status, payload):
            self._send(status, json.dumps(payload, indent=2, sort_keys=True))

        def do_GET(self):
            parsed = urlparse(self.path)
            if parsed.path == "/healthz":
                self._json(200, {"ok": True, "utc": utc_now_text()})
                return

            if parsed.path in ("/admin", "/api/admin/summary", "/api/admin/weekly-report") and not require_key(self, api_key):
                self._json(401, {"error": "missing or invalid key"})
                return

            query = parse_qs(parsed.query)
            if parsed.path == "/admin":
                days = int((query.get("days") or ["7"])[0])
                summary = store.summary(days=days)
                self._send(200, render_summary_html(summary), "text/html; charset=utf-8")
                return

            if parsed.path == "/api/admin/summary":
                days = int((query.get("days") or ["7"])[0])
                self._json(200, store.summary(days=days))
                return

            if parsed.path == "/api/admin/weekly-report":
                weeks = int((query.get("weeks") or ["8"])[0])
                fmt = str((query.get("format") or ["json"])[0]).strip().lower()
                rows = store.weekly_report(weeks=weeks)
                if fmt == "markdown":
                    self._send(200, render_weekly_markdown(rows), "text/markdown; charset=utf-8")
                    return
                if fmt == "csv":
                    self._send(200, render_weekly_csv(rows), "text/csv; charset=utf-8")
                    return
                self._json(200, {"weeks": weeks, "rows": rows})
                return

            self._json(404, {"error": "not found"})

        def do_POST(self):
            parsed = urlparse(self.path)
            if parsed.path != "/api/usage":
                self._json(404, {"error": "not found"})
                return

            if not require_key(self, api_key):
                self._json(401, {"error": "missing or invalid key"})
                return

            try:
                length = int(self.headers.get("Content-Length") or "0")
            except Exception:
                length = 0
            raw_body = self.rfile.read(length)
            try:
                payload = json.loads(raw_body.decode("utf-8") if raw_body else "{}")
            except Exception:
                self._json(400, {"error": "invalid json"})
                return

            events = payload.get("events") if isinstance(payload, dict) else None
            if not isinstance(events, list):
                self._json(400, {"error": "expected events list"})
                return

            accepted = store.ingest_events(events)
            self._json(200, {"accepted": accepted, "received": len(events)})

        def log_message(self, format_text, *args):
            return

    return UsageHandler


def main():
    parser = argparse.ArgumentParser(description="WWPTools telemetry backend")
    parser.add_argument("--host", default="0.0.0.0")
    parser.add_argument("--port", type=int, default=8787)
    parser.add_argument("--db", default=os.path.join("data", "wwptools-usage.db"))
    parser.add_argument("--api-key", default=os.environ.get("WWPTOOLS_TELEMETRY_KEY", ""))
    args = parser.parse_args()

    store = TelemetryStore(args.db)
    server = ThreadingHTTPServer((args.host, args.port), build_handler(store, args.api_key))
    print("WWPTools telemetry backend listening on http://{}:{}/".format(args.host, args.port))
    print("Database: {}".format(os.path.abspath(args.db)))
    if args.api_key:
        print("API key protection: enabled")
    else:
        print("API key protection: disabled")
    server.serve_forever()


if __name__ == "__main__":
    main()
