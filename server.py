import os, json, sqlite3, base64, uuid
from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler
from urllib.parse import urlparse
from pathlib import Path

ROOT = Path(__file__).resolve().parent
DB_PATH = Path(os.getenv('DB_PATH', ROOT / 'data' / 'truckcontrol.sqlite3'))
DB_PATH.parent.mkdir(parents=True, exist_ok=True)
API_KEY = os.getenv('API_KEY', '')
PORT = int(os.getenv('PORT', '8080'))

TABLES = {
    'trips': ['id','no','date','from','to','weight','status','startKm','endKm','driver','truck','trailer'],
    'drivers': ['id','name','phone'],
    'trucks': ['id','plate','model'],
    'trailers': ['id','plate','type'],
    'fuel': ['id','type','date','liters','cost','currency','station','tripId','trip'],
    'expenses': ['id','tripId','trip','date','category','amount','currency','payment','note'],
    'route': ['id','city','type','note','date'],
    'docs': ['id','tripId','trip','type','name','note','data','mime'],
    'gps': ['id','date','lat','lon','accuracy'],
    'maintenance': ['id','date','km','truck','type','note','cost'],
}

SCHEMA = '''
CREATE TABLE IF NOT EXISTS company (id INTEGER PRIMARY KEY CHECK(id=1), name TEXT NOT NULL DEFAULT '', server TEXT NOT NULL DEFAULT '');
CREATE TABLE IF NOT EXISTS users (id INTEGER PRIMARY KEY, name TEXT NOT NULL, role TEXT NOT NULL);
CREATE TABLE IF NOT EXISTS trips (id TEXT PRIMARY KEY, no TEXT, date TEXT, "from" TEXT, "to" TEXT, weight TEXT, status TEXT, startKm TEXT, endKm TEXT, driver TEXT, truck TEXT, trailer TEXT);
CREATE TABLE IF NOT EXISTS drivers (id TEXT PRIMARY KEY, name TEXT, phone TEXT);
CREATE TABLE IF NOT EXISTS trucks (id TEXT PRIMARY KEY, plate TEXT, model TEXT);
CREATE TABLE IF NOT EXISTS trailers (id TEXT PRIMARY KEY, plate TEXT, type TEXT);
CREATE TABLE IF NOT EXISTS fuel (id TEXT PRIMARY KEY, type TEXT, date TEXT, liters TEXT, cost TEXT, currency TEXT, station TEXT, tripId TEXT, trip TEXT);
CREATE TABLE IF NOT EXISTS expenses (id TEXT PRIMARY KEY, tripId TEXT, trip TEXT, date TEXT, category TEXT, amount TEXT, currency TEXT, payment TEXT, note TEXT);
CREATE TABLE IF NOT EXISTS route (id TEXT PRIMARY KEY, city TEXT, type TEXT, note TEXT, date TEXT);
CREATE TABLE IF NOT EXISTS docs (id TEXT PRIMARY KEY, tripId TEXT, trip TEXT, type TEXT, name TEXT, note TEXT, data TEXT, mime TEXT);
CREATE TABLE IF NOT EXISTS gps (id TEXT PRIMARY KEY, date TEXT, lat REAL, lon REAL, accuracy REAL);
CREATE TABLE IF NOT EXISTS maintenance (id TEXT PRIMARY KEY, date TEXT, km TEXT, truck TEXT, type TEXT, note TEXT, cost TEXT);
'''

def conn():
    c = sqlite3.connect(DB_PATH)
    c.row_factory = sqlite3.Row
    c.execute('PRAGMA journal_mode=WAL')
    c.execute('PRAGMA foreign_keys=ON')
    c.executescript(SCHEMA)
    c.execute("INSERT OR IGNORE INTO company(id,name,server) VALUES(1,'Моя транспортная компания','')")
    return c

def auth(handler):
    return not API_KEY or handler.headers.get('X-API-Key') == API_KEY

def clean(v):
    if v is None: return ''
    if isinstance(v, (dict,list)): return json.dumps(v, ensure_ascii=False)
    return v

def row_to_dict(row):
    return dict(row)

def get_state():
    c=conn()
    try:
        company=c.execute('SELECT name,server FROM company WHERE id=1').fetchone()
        user=c.execute('SELECT name,role FROM users ORDER BY id DESC LIMIT 1').fetchone()
        state={'company':dict(company) if company else {'name':'','server':''},
               'user':dict(user) if user else {'name':'Пользователь','role':'dispatcher'}}
        for name, cols in TABLES.items():
            rows=c.execute('SELECT * FROM '+name).fetchall()
            state[name]=[row_to_dict(r) for r in rows]
        return state
    finally: c.close()

def export_excel(state):
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill
        path = ROOT / 'data' / 'TruckControl_Отчёт.xlsx'
        path.parent.mkdir(parents=True, exist_ok=True)
        wb = Workbook()
        wb.remove(wb.active)
        labels = {'trips':'Рейсы','drivers':'Водители','trucks':'Тягачи','trailers':'Прицепы','fuel':'Топливо','expenses':'Расходы','maintenance':'ТО','route':'Маршрут','docs':'Документы','gps':'GPS'}
        for key, title in labels.items():
            ws = wb.create_sheet(title)
            rows = state.get(key, []) or []
            if key == 'trips':
                rows = [dict(row, **{'Километраж отчёта': max(0, float(row.get('endKm') or 0)-float(row.get('startKm') or 0))}) for row in rows]
            cols = (['no','date','from','to','weight','status','startKm','endKm','Километраж отчёта','driver','truck','trailer'] if key == 'trips' else (list(dict.fromkeys(c for row in rows for c in row.keys() if c != 'data')) if rows else TABLES[key]))
            ws.append(cols)
            for cell in ws[1]:
                cell.font = Font(bold=True, color='FFFFFF')
                cell.fill = PatternFill('solid', fgColor='0F4C5C')
            for row in rows:
                ws.append([str(row.get(c, '') if row.get(c) is not None else '') for c in cols])
            ws.freeze_panes = 'A2'
            ws.auto_filter.ref = ws.dimensions
            for column in ws.columns:
                letter = column[0].column_letter
                ws.column_dimensions[letter].width = min(max(max(len(str(c.value or '')) for c in column)+2, 12), 36)
        # Summary worksheet with live-calculated mileage and totals
        ws = wb.create_sheet('Сводка', 0)
        ws.append(['Показатель','Значение'])
        ws.append(['Всего пройдено, км', '=SUM(Рейсы!I:I)'])
        ws.append(['Количество рейсов', '=COUNTA(Рейсы!A:A)-1'])
        ws.append(['Всего заправок', '=COUNTA(Топливо!A:A)-1'])
        ws.append(['Количество записей ТО', '=COUNTA(ТО!A:A)-1'])
        for c in ws[1]:
            c.font = Font(bold=True, color='FFFFFF'); c.fill = PatternFill('solid', fgColor='0F4C5C')
        ws.column_dimensions['A'].width=28; ws.column_dimensions['B'].width=22
        wb.save(path)
    except Exception as exc:
        print('Excel export skipped:', exc)

def replace_state(state):
    c=conn()
    try:
        c.execute('BEGIN')
        company=state.get('company') or {}
        c.execute('UPDATE company SET name=?, server=? WHERE id=1',(clean(company.get('name')),clean(company.get('server'))))
        c.execute('DELETE FROM users')
        user=state.get('user') or {}
        c.execute('INSERT INTO users(id,name,role) VALUES(1,?,?)',(clean(user.get('name')),clean(user.get('role','dispatcher'))))
        for name, cols in TABLES.items():
            c.execute('DELETE FROM '+name)
            placeholders=','.join('?' for _ in cols)
            quoted_cols=','.join('"from"' if x=='from' else '"to"' if x=='to' else x for x in cols)
            sql=f'INSERT INTO {name} ({quoted_cols}) VALUES ({placeholders})'
            for item in state.get(name,[]) or []:
                vals=[clean(item.get(col)) for col in cols]
                c.execute(sql, vals)
        c.commit()
        export_excel(state)
    except Exception:
        c.rollback(); raise
    finally: c.close()

def json_response(handler, status, payload):
    data=json.dumps(payload, ensure_ascii=False).encode('utf-8')
    handler.send_response(status)
    handler.send_header('Content-Type','application/json; charset=utf-8')
    handler.send_header('Content-Length',str(len(data)))
    handler.send_header('Access-Control-Allow-Origin','*')
    handler.send_header('Access-Control-Allow-Headers','Content-Type, X-API-Key')
    handler.send_header('Access-Control-Allow-Methods','GET, PUT, POST, OPTIONS')
    handler.end_headers(); handler.wfile.write(data)

class Handler(SimpleHTTPRequestHandler):
    def __init__(self,*args,**kwargs): super().__init__(*args,directory=str(ROOT),**kwargs)
    def do_OPTIONS(self):
        self.send_response(204)
        self.send_header('Access-Control-Allow-Origin','*')
        self.send_header('Access-Control-Allow-Headers','Content-Type, X-API-Key')
        self.send_header('Access-Control-Allow-Methods','GET, PUT, POST, OPTIONS')
        self.end_headers()
    def do_GET(self):
        path=urlparse(self.path).path
        if path=='/api/health':
            json_response(self,200,{'ok':True,'database':str(DB_PATH.name)})
        elif path=='/api/state':
            if not auth(self): json_response(self,401,{'error':'Unauthorized'}); return
            try: json_response(self,200,get_state())
            except Exception as e: json_response(self,500,{'error':str(e)})
        else: super().do_GET()
    def do_PUT(self):
        path=urlparse(self.path).path
        if path!='/api/state': json_response(self,404,{'error':'Not found'}); return
        if not auth(self): json_response(self,401,{'error':'Unauthorized'}); return
        try:
            length=int(self.headers.get('Content-Length','0'))
            body=self.rfile.read(length)
            state=json.loads(body.decode('utf-8'))
            replace_state(state)
            json_response(self,200,{'ok':True,'state':get_state()})
        except Exception as e:
            json_response(self,400,{'error':str(e)})

if __name__=='__main__':
    print(f'TruckControl server listening on 0.0.0.0:{PORT}')
    ThreadingHTTPServer(('0.0.0.0',PORT),Handler).serve_forever()
