#!/usr/bin/env python3
"""
Attendance Tracker - Lightweight HTTP Server

Simple HTTP server for team attendance tracking using local JSON file storage.
"""

from http.server import HTTPServer, SimpleHTTPRequestHandler
import json
import os
import sys
import socket
from urllib.parse import urlparse

PORT      = 8080
DATA_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'attendance_data.json')


class AttendanceHandler(SimpleHTTPRequestHandler):
    """HTTP handler for attendance data - uses JSON file storage."""

    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path == '/api/attendance':
            data = self._load_json()
            self._json_response(200, data)
            return
        if parsed.path == '/':
            self.path = '/team-attendance-tracker-sharepoint.html'
        return SimpleHTTPRequestHandler.do_GET(self)

    def do_POST(self):
        parsed = urlparse(self.path)
        if parsed.path == '/api/attendance':
            length   = int(self.headers.get('Content-Length', 0))
            raw      = self.rfile.read(length)
            try:
                new_data = json.loads(raw.decode('utf-8'))
                existing = self._load_json()
                existing.update(new_data)
                self._save_json(existing)
                self._json_response(200, {'status': 'success', 'message': 'Data saved'})
            except Exception as e:
                self._json_response(500, {'status': 'error', 'message': str(e)})
            return

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()

    def _json_response(self, code: int, payload):
        body = json.dumps(payload, ensure_ascii=False).encode('utf-8')
        self.send_response(code)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(body)))
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        self.wfile.write(body)

    def _load_json(self) -> dict:
        if os.path.exists(DATA_FILE):
            try:
                with open(DATA_FILE, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception:
                pass
        return {}

    def _save_json(self, data: dict):
        with open(DATA_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)

    def log_message(self, fmt, *args):
        print(f'[{self.log_date_time_string()}] {fmt % args}')


def _get_local_ip() -> str:
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(('8.8.8.8', 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return 'localhost'


def main():
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    local_ip = _get_local_ip()
    httpd = HTTPServer(('', PORT), AttendanceHandler)

    print('=' * 70)
    print('Team Attendance Tracker Server')
    print('=' * 70)
    print(f'\nRunning on:')
    print(f'   Local:   http://localhost:{PORT}')
    print(f'   Network: http://{local_ip}:{PORT}')
    print(f'\nStorage: Local JSON file ({DATA_FILE})')
    print('\nPress Ctrl+C to stop\n')
    print('=' * 70 + '\n')
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print('\nServer stopped.')


if __name__ == '__main__':
    main()
