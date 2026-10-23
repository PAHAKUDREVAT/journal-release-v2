import sqlite3
import os
import sys
import socket
import webbrowser
import threading
from datetime import datetime
from io import BytesIO
from flask import Flask, request, redirect, url_for, render_template_string, send_file, make_response, jsonify
from openpyxl import Workbook

if getattr(sys, 'frozen', False):
BASE_DIR = os.path.dirname(sys.executable)
else:
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DB_PATH = os.path.join(BASE_DIR, 'journal.db')
app = Flask(__name__)

def init_db():
conn = sqlite3.connect(DB_PATH)
c = conn.cursor()
c.execute('''CREATE TABLE IF NOT EXISTS requests (
id INTEGER PRIMARY KEY AUTOINCREMENT,
area TEXT,
created_at TEXT NOT NULL,
project_no TEXT,
horizon TEXT,
work_type TEXT,
created_by TEXT NOT NULL,
phone TEXT,
safety TEXT,
priority TEXT,
executor TEXT,
completion_date TEXT,
status TEXT,
comment TEXT,
updated_at TEXT)''')
c.execute('''CREATE TRIGGER IF NOT EXISTS prevent_created_at_update BEFORE UPDATE ON requests
FOR EACH ROW WHEN OLD.created_at <> NEW.created_at BEGIN SELECT RAISE(ABORT, 'immutable'); END''')
c.execute('''CREATE TRIGGER IF NOT EXISTS prevent_delete BEFORE DELETE ON requests
FOR EACH ROW BEGIN SELECT RAISE(ABORT, 'delete forbidden'); END''')
conn.commit()
conn.close()

def get_local_ip():
s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
try: s.connect(("8.8.8.8", 80)); return s.getsockname()[0]
except: return "127.0.0.1"
finally: s.close()

HTML_PAGE = """
<!DOCTYPE html><html lang="ru"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Журнал заявок</title>
<style>
body{font-family:Arial,sans-serif;background:#f4f6f8;margin:0;color:#222;padding:20px}
.card{background:#fff;padding:20px;border-radius:8px;box-shadow:0 2px 5px rgba(0,0,0,0.1);max-width:100%
