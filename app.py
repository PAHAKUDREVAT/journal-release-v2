import sqlite3
import os
import sys
import socket
import webbrowser
import threading
from datetime import datetime
from io import BytesIO
from flask import Flask, request, redirect, url_for, render_template_string, send_file, make_response
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
        id INTEGER PRIMARY KEY AUTOINCREMENT, created_at TEXT NOT NULL, created_by TEXT NOT NULL,
        phone TEXT, horizon TEXT, area TEXT, safety TEXT, description TEXT, executor TEXT, 
        priority TEXT, status TEXT, comment TEXT, updated_at TEXT)''')
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
.card{background:#fff;padding:20px;border-radius:8px;box-shadow:0 2px 5px rgba(0,0,0,0.1);max-width:1600px;margin:0 auto}
table{width:100%;border-collapse:collapse;margin-top:15px;font-size:12px;table-layout:fixed}
th,td{padding:6px;border-bottom:1px solid #eee;text-align:left;word-wrap:break-word;vertical-align:top;}
th{background:#f8f9fa} input,select,textarea{width:100%;padding:8px;margin:5px 0;box-sizing:border-box;font-family:inherit;font-size:13px;}
.btn{background:#1e3a5f;color:#fff;padding:10px;border:none;border-radius:4px;cursor:pointer;width:100%;margin-top:10px;font-weight:bold;}
.high{background-color:#ffe6e6;color:#b03030;font-weight:bold;}
.col-form{display:flex;gap:10px;margin-bottom:10px;}
.col{flex:1;} label{font-size:12px;color:#555;font-weight:bold;} a{color:#1e3a5f;}
</style></head><body><div class="card">
<h2 style="margin-top:0;">Журнал заявок маркшейдерской службы</h2>
{% if edit_mode %}<h3>Редактирование заявки №{{ r['id'] }}</h3>{% endif %}
<form method="POST" action="{% if edit_mode %}/update/{{ r['id'] }}{% else %}/create{% endif %}">
<div class="col-form">
    <div class="col">
        <label>Подал (ФИО):</label>
        <input name="author" required value="{{ r['created_by'] if r else saved_author }}">
        <label>Телефон:</label>
        <input name="phone" value="{{ r['phone'] if r else saved_phone }}">
        <label>Дата подачи (не меняется):</label>
        <input value="{{ r['created_at'] if r else now }}" readonly style="background:#eee;">
    </div>
    <div class="col">
        <label>Горизонт:</label>
        <input name="horizon" required value="{{ r['horizon'] if r else '' }}">
        <label>Участок:</label>
        <input name="area" value="{{ r['area'] if r else '' }}">
        <label>Безопасность:</label>
        <select name="safety">
            <option value="Да" {% if r and r['safety'] == 'Да' %}selected{% endif %}>Да (обезопасено)</option>
            <option value="Нет" {% if r and r['safety'] == 'Нет' %}selected{% endif %}>Нет (опасно)</option>
            <option value="Не требуется" {% if r and r['safety'] == 'Не требуется' %}selected{% endif %}>Не требуется</option>
        </select>
    </div>
    <div class="col">
        <label>Приоритет:</label>
        <select name="priority">
            <option value="Низкий" {% if r and r['priority'] == 'Низкий' %}selected{% endif %}>Низкий</option>
            <option value="Обычный" {% if not r or r['priority'] == 'Обычный' %}selected{% endif %}>Обычный</option>
            <option value="Высокий" {% if r and r['priority'] == 'Высокий' %}selected{% endif %}>Высокий</option>
            <option value="Срочный" {% if r and r['priority'] == 'Срочный' %}selected{% endif %}>Срочный</option>
        </select>
        <label>Исполнитель:</label>
        <input name="executor" value="{{ r['executor'] if r else '' }}">
        <label>Статус:</label>
        <select name="status">
            <option value="Новая" {% if not r or r['status'] == 'Новая' %}selected{% endif %}>Новая</option>
            <option value="Принята" {% if r and r['status'] == 'Принята' %}selected{% endif %}>Принята</option>
            <option value="В работе" {% if r and r['status'] == 'В работе' %}selected{% endif %}>В работе</option>
            <option value="Выполнена" {% if r and r['status'] == 'Выполнена' %}selected{% endif %}>Выполнена</option>
            <option value="Отклонена" {% if r and r['status'] == 'Отклонена' %}selected{% endif %}>Отклонена</option>
        </select>
    </div>
</div>
<label>Описание:</label>
<textarea name="desc" rows="2">{{ r['description'] if r else '' }}</textarea>
<label>Комментарий:</label>
<textarea name="comment" rows="1">{{ r['comment'] if r else '' }}</textarea>
<button type="submit" class="btn">{{ 'Сохранить изменения' if edit_mode else 'Подать заявку' }}</button>
</form>
{% if edit_mode %}<div style="text-align:center;margin-top:10px;"><a href="/">Отмена</a></div>{% endif %}
{% if not edit_mode %}
<hr style="margin:20px 0; border:0; border-top:1px solid #ddd;">
<table>
<tr><th style="width:3%">№</th><th style="width:8%">Дата</th><th style="width:8%">Подал</th><th style="width:7%">Тел</th>
<th style="width:7%">Горизонт</th><th style="width:7%">Участок</th><th style="width:8%">Безопасность</th><th style="width:7%">Приоритет</th>
<th style="width:8%">Статус</th><th style="width:8%">Исполнитель</th><th style="width:15%">Описание</th><th style="width:8%">Коммент.</th><th style="width:6%">Действие</th></tr>
{% for req in requests %}
<tr class="{% if req['priority'] in ['Высокий','Срочный'] and req['status'] not in ['Выполнена', 'Отклонена'] %}high{% endif %}">
    <td>{{req['id']}}</td><td>{{req['created_at']}}</td><td>{{req['created_by']}}</td><td>{{req['phone']}}</td><td>{{req['horizon']}}</td><td>{{req['area']}}</td><td>{{req['safety']}}</td>
    <td>{{req['priority']}}</td><td><b>{{req['status']}}</b></td><td>{{req['executor']}}</td><td>{{req['description']}}</td><td>{{req['comment']}}</td><td><a href="/edit/{{req['id']}}">Ред.</a></td>
</tr>
{% endfor %}
</table>
<div style="margin-top:20px;"><a href="/export" style="background:#2a6e2a;color:#fff;padding:10px 15px;text-decoration:none;border-radius:4px;font-weight:bold;">Скачать Excel (для Р7)</a></div>
{% endif %}
</div></body></html>
"""
@app.route('/')
def index():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    rows = conn.execute("SELECT * FROM requests ORDER BY id DESC").fetchall()
    conn.close()
    now = datetime.now().strftime("%d.%m.%Y %H:%M:%S")
    saved_author = request.cookies.get('saved_author', '')
    saved_phone = request.cookies.get('saved_phone', '')
    return render_template_string(HTML_PAGE, requests=rows, edit_mode=False, now=now, r=None, saved_author=saved_author, saved_phone=saved_phone)
@app.route('/create', methods=['POST'])
def create():
    now = datetime.now().strftime("%d.%m.%Y %H:%M:%S")
    conn = sqlite3.connect(DB_PATH)
    conn.execute("INSERT INTO requests (created_at,created_by,phone,horizon,area,safety,priority,description,comment,executor,status,updated_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
                 (now, request.form['author'], request.form.get('phone',''), request.form['horizon'], request.form.get('area',''), request.form['safety'], request.form['priority'], request.form.get('desc',''), request.form.get('comment',''), request.form.get('executor',''), 'Новая', now))
    conn.commit(); conn.close()
    
    resp = make_response(redirect(url_for('index')))
    resp.set_cookie('saved_author', request.form['author'], max_age=365*24*60*60)
    resp.set_cookie('saved_phone', request.form.get('phone',''), max_age=365*24*60*60)
    return resp
@app.route('/edit/<int:id>')
def edit(id):
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    r = conn.execute("SELECT * FROM requests WHERE id=?", (id,)).fetchone()
    conn.close()
    if not r: return "Заявка не найдена", 404
    return render_template_string(HTML_PAGE, r=r, edit_mode=True)
@app.route('/update/<int:id>', methods=['POST'])
def update(id):
    now = datetime.now().strftime("%d.%m.%Y %H:%M:%S")
    conn = sqlite3.connect(DB_PATH)
    conn.execute("UPDATE requests SET created_by=?, phone=?, horizon=?, area=?, safety=?, priority=?, description=?, comment=?, executor=?, status=?, updated_at=? WHERE id=?",
                 (request.form['author'], request.form.get('phone',''), request.form['horizon'], request.form.get('area',''), request.form['safety'], request.form['priority'], request.form.get('desc',''), request.form.get('comment',''), request.form.get('executor',''), request.form['status'], now, id))
    conn.commit(); conn.close()
    return redirect(url_for('index'))
@app.route('/export')
def export():
    conn = sqlite3.connect(DB_PATH)
    rows = conn.execute("SELECT id,created_at,created_by,phone,horizon,area,safety,priority,status,executor,description,comment FROM requests ORDER BY id").fetchall()
    conn.close()
    wb = Workbook(); ws = wb.active
    ws.append(["№","Дата","Подал","Телефон","Горизонт","Участок","Безопасность","Приоритет","Статус","Исполнитель","Описание","Комментарий"])
    for r in rows: ws.append(list(r))
    buf = BytesIO(); wb.save(buf); buf.seek(0)
    return send_file(buf, as_attachment=True, download_name="Журнал.xlsx")
def open_browser():
    import time; time.sleep(1.5); webbrowser.open("http://127.0.0.1:5001")
if __name__ == '__main__':
    init_db()
    ip = get_local_ip()
    print(f"СЕРВЕР РАБОТАЕТ. Адрес: http://{ip}:5001")
    threading.Thread(target=open_browser, daemon=True).start()
    app.run(host='0.0.0.0', port=5001, debug=False, use_reloader=False)
