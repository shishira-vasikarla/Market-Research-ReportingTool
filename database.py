import sqlite3, json, uuid
from datetime import datetime
import pandas as pd
from io import StringIO
DB_PATH = 'project_history.db'
def connect():
    c=sqlite3.connect(DB_PATH); c.row_factory=sqlite3.Row; c.execute('PRAGMA foreign_keys=ON'); return c
def init_db():
    with connect() as c:
        c.executescript('''
        CREATE TABLE IF NOT EXISTS projects(project_id TEXT PRIMARY KEY,name TEXT NOT NULL,company_data TEXT NOT NULL,questions_data TEXT NOT NULL,template_data TEXT,created_at TEXT NOT NULL,updated_at TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS answers(answer_id TEXT PRIMARY KEY,project_id TEXT NOT NULL,company_url TEXT,company_info TEXT,question TEXT NOT NULL,selected_version_id TEXT,created_at TEXT NOT NULL,FOREIGN KEY(project_id) REFERENCES projects(project_id) ON DELETE CASCADE);
        CREATE TABLE IF NOT EXISTS answer_versions(version_id TEXT PRIMARY KEY,answer_id TEXT NOT NULL,version_number INTEGER NOT NULL,answer_text TEXT NOT NULL,note TEXT DEFAULT '',status TEXT NOT NULL DEFAULT 'Draft',created_at TEXT NOT NULL,FOREIGN KEY(answer_id) REFERENCES answers(answer_id) ON DELETE CASCADE,UNIQUE(answer_id,version_number));
        CREATE TABLE IF NOT EXISTS reports(report_id TEXT PRIMARY KEY,project_id TEXT NOT NULL,rows_json TEXT NOT NULL,source_versions TEXT NOT NULL,created_at TEXT NOT NULL,FOREIGN KEY(project_id) REFERENCES projects(project_id) ON DELETE CASCADE);
        ''')
def touch(pid):
    with connect() as c: c.execute('UPDATE projects SET updated_at=? WHERE project_id=?',(datetime.now().isoformat(timespec='seconds'),pid))
def create_project(name, company, questions, template):
    pid=str(uuid.uuid4()); now=datetime.now().isoformat(timespec='seconds')
    with connect() as c: c.execute('INSERT INTO projects VALUES(?,?,?,?,?,?,?)',(pid,name,company.to_json(orient='records'),questions.to_json(orient='records'),template.to_json(orient='records') if template is not None else None,now,now))
    return pid
def list_projects():
    with connect() as c: return [dict(x) for x in c.execute('SELECT project_id,name,created_at,updated_at FROM projects ORDER BY updated_at DESC')]
def get_project(pid):
    with connect() as c: r=c.execute('SELECT * FROM projects WHERE project_id=?',(pid,)).fetchone()
    if not r:return None
    return {'name':r['name'],'company':pd.read_json(StringIO(r['company_data'])),'questions':pd.read_json(StringIO(r['questions_data'])),'template':pd.read_json(StringIO(r['template_data'])) if r['template_data'] else None}
def create_answer(pid, aid, url, info, question, answer):
    now=datetime.now().isoformat(timespec='seconds'); vid=str(uuid.uuid4())
    with connect() as c:
        c.execute('INSERT INTO answers VALUES(?,?,?,?,?,?,?)',(aid,pid,url,info,question,vid,now))
        c.execute('INSERT INTO answer_versions VALUES(?,?,?,?,?,?,?)',(vid,aid,1,answer,'Original answer','Draft',now))
    touch(pid)
def add_version(aid, text, note):
    now=datetime.now().isoformat(timespec='seconds'); vid=str(uuid.uuid4())
    with connect() as c:
        n=c.execute('SELECT COALESCE(MAX(version_number),0)+1 n FROM answer_versions WHERE answer_id=?',(aid,)).fetchone()['n']
        c.execute('INSERT INTO answer_versions VALUES(?,?,?,?,?,?,?)',(vid,aid,n,text,note,'Draft',now))
        c.execute('UPDATE answers SET selected_version_id=? WHERE answer_id=?',(vid,aid))
        pid=c.execute('SELECT project_id FROM answers WHERE answer_id=?',(aid,)).fetchone()['project_id']
    touch(pid)
def get_answers(pid):
    with connect() as c: rows=c.execute('''SELECT a.*,v.version_id latest_id,v.version_number latest_number,v.answer_text latest_text FROM answers a JOIN answer_versions v ON v.answer_id=a.answer_id AND v.version_number=(SELECT MAX(version_number) FROM answer_versions WHERE answer_id=a.answer_id) WHERE a.project_id=? ORDER BY a.created_at''',(pid,)).fetchall()
    return [dict(r) for r in rows]
def versions(aid):
    with connect() as c: return [dict(r) for r in c.execute('SELECT * FROM answer_versions WHERE answer_id=? ORDER BY version_number',(aid,))]
def choose(aid,vid):
    with connect() as c:
        if not c.execute('SELECT 1 FROM answer_versions WHERE answer_id=? AND version_id=?',(aid,vid)).fetchone(): raise ValueError('Version does not belong to this answer')
        c.execute('UPDATE answers SET selected_version_id=? WHERE answer_id=?',(vid,aid)); pid=c.execute('SELECT project_id FROM answers WHERE answer_id=?',(aid,)).fetchone()['project_id']
    touch(pid)
def approve(vid):
    with connect() as c: c.execute("UPDATE answer_versions SET status='Approved' WHERE version_id=?",(vid,))
def selected_answers(pid):
    with connect() as c: return [dict(r) for r in c.execute('''SELECT a.answer_id,a.company_url,a.question,v.version_id,v.version_number,v.answer_text,v.status FROM answers a JOIN answer_versions v ON v.version_id=a.selected_version_id WHERE a.project_id=? ORDER BY a.created_at''',(pid,))]
def save_report(pid, rows, sources):
    rid=str(uuid.uuid4()); now=datetime.now().isoformat(timespec='seconds')
    with connect() as c:c.execute('INSERT INTO reports VALUES(?,?,?,?,?)',(rid,pid,json.dumps(rows),json.dumps(sources),now))
    touch(pid); return rid
def reports(pid):
    with connect() as c:return [dict(r) for r in c.execute('SELECT report_id,created_at FROM reports WHERE project_id=? ORDER BY created_at DESC',(pid,))]
def get_report(rid):
    with connect() as c:r=c.execute('SELECT * FROM reports WHERE report_id=?',(rid,)).fetchone()
    return {'rows':json.loads(r['rows_json']),'sources':json.loads(r['source_versions'])} if r else None
