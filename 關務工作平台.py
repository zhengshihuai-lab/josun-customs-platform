#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
關務工作平台 v3.0 — Ticket-Based Customs Work Management Platform
A real working platform where customs staff can fill in data, attach files,
and manage each shipment ticket through the entire workflow.

Usage:  python3 關務工作平台.py
Then open: http://localhost:8899
"""

import http.server, json, subprocess, os, sys, threading, webbrowser
import urllib.parse, cgi, shutil, uuid, mimetypes
from datetime import datetime
from http.server import HTTPServer
from socketserver import ThreadingMixIn

PORT = int(os.environ.get('PORT', 8899))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, '平台數據')
IS_CLOUD = os.environ.get('RENDER') or os.environ.get('RAILWAY_ENV') or os.environ.get('DYNO')

# Auto-create data directories
for d in ['step_data', 'uploads', 'outputs']:
    os.makedirs(os.path.join(DATA_DIR, d), exist_ok=True)

# ─────────────────────── Data Management Functions ───────────────────────

def load_tickets():
    path = os.path.join(DATA_DIR, 'tickets.json')
    if os.path.exists(path):
        with open(path, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {'tickets': []}

def save_tickets(data):
    path = os.path.join(DATA_DIR, 'tickets.json')
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def create_ticket(data):
    tickets = load_tickets()
    n = len(tickets['tickets']) + 1
    ticket_id = f"TICKET-{n:03d}"
    ticket = {
        'id': ticket_id,
        'name': data.get('name', ''),
        'po_number': data.get('po_number', ''),
        'type': data.get('type', 'export'),
        'country': data.get('country', 'thailand'),
        'created': datetime.now().isoformat(),
        'status': 'active',
        'completed_steps': [],
        'notes': data.get('notes', '')
    }
    tickets['tickets'].append(ticket)
    save_tickets(tickets)
    os.makedirs(os.path.join(DATA_DIR, 'uploads', ticket_id), exist_ok=True)
    os.makedirs(os.path.join(DATA_DIR, 'outputs', ticket_id), exist_ok=True)
    return ticket

def update_ticket(ticket_id, data):
    tickets = load_tickets()
    for t in tickets['tickets']:
        if t['id'] == ticket_id:
            for k, v in data.items():
                if k not in ('id', 'created'):
                    t[k] = v
            break
    save_tickets(tickets)
    return {'success': True}

def load_ticket(ticket_id):
    tickets = load_tickets()
    for t in tickets['tickets']:
        if t['id'] == ticket_id:
            return t
    return {}

def load_step_data(ticket_id, step_id):
    path = os.path.join(DATA_DIR, 'step_data', f'{ticket_id}_{step_id}.json')
    if os.path.exists(path):
        with open(path, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}

def save_step_data(ticket_id, step_id, data):
    path = os.path.join(DATA_DIR, 'step_data', f'{ticket_id}_{step_id}.json')
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def mark_step_complete(ticket_id, step_id):
    tickets = load_tickets()
    for t in tickets['tickets']:
        if t['id'] == ticket_id:
            if step_id not in t['completed_steps']:
                t['completed_steps'].append(step_id)
            # Check if all steps are done
            all_steps = get_all_step_ids()
            if all(s in t['completed_steps'] for s in all_steps):
                t['status'] = 'completed'
            break
    save_tickets(tickets)

def get_all_step_ids():
    return [
        'm1s1','m1s2','m1s3',
        'm2s1','m2s2','m2s3','m2s4','m2s5','m2s6','m2s7','m2s8',
        'm3s1','m3s2','m3s3','m3s4','m3s5','m3s6',
        'm4s1','m4s2','m4s3','m4s4',
        'm5s1','m5s2','m5s3','m5s4','m5s5',
        'm6s1','m6s2','m6s3','m6s4',
        'm7s1','m7s2','m7s3','m7s4','m7s5','m7s6',
        'm8s1','m8s2','m8s3','m8s4','m8s5',
        'm9s1','m9s2','m9s3',
        'm10s1','m10s2','m10s3',
        'm11s1','m11s2','m11s3',
        'm12s1','m12s2','m12s3','m12s4',
    ]

def get_system_status():
    tickets = load_tickets()
    active = sum(1 for t in tickets['tickets'] if t['status'] == 'active')
    completed = sum(1 for t in tickets['tickets'] if t['status'] == 'completed')
    return {
        'server': 'running',
        'time': datetime.now().isoformat(),
        'total_tickets': len(tickets['tickets']),
        'active_tickets': active,
        'completed_tickets': completed,
        'tools_available': len(TOOL_FILES),
        'data_dir': DATA_DIR
    }

TOOL_FILES = {
    'hs-classify': 'HS編碼歸類助手.py',
    'shipping-mark': '嘜頭裝箱單生成器.py',
    'doc-generator': '關務單據自動生成器.py',
    'doc-validator': '報關單據校驗助手.py',
    'erp-import': 'ERP批量錄入工具.py',
    'tax-calc': '稅費核算引擎.py',
    'indonesia-collab': '印尼協同自動化工具.py',
    'freight-booking': '訂艙比價工具.py',
    'brand-auth': '品牌授權管理系統.py',
    'policy-tracker': '政策追蹤工具.py',
    'data-archive': '資料歸檔工具.py',
}

def get_tool_list():
    tools = []
    for key, fname in TOOL_FILES.items():
        fpath = os.path.join(BASE_DIR, fname)
        tools.append({
            'id': key,
            'filename': fname,
            'available': os.path.exists(fpath)
        })
    return {'tools': tools}

def get_workflows():
    return {
        'modules': MODULES_DEF
    }

MODULES_DEF = [
    {'id': 1, 'name': '出貨準備', 'icon': '📦', 'steps': ['ERP導出','確認品項','缺貨確認']},
    {'id': 2, 'name': '包裝與標識管理', 'icon': '📐', 'steps': ['IPPC檢查','散件並箱','稱重','裝箱清單','紙箱嘜頭','櫃型核算','托盤嘜頭','PACKING LIST']},
    {'id': 3, 'name': '報關資料準備', 'icon': '📋', 'steps': ['HS歸類','型錄收集','工程師確認','申報要素','代理核驗','FTA確認']},
    {'id': 4, 'name': '單據製作與審核', 'icon': '📄', 'steps': ['全套單據','校驗','提單確認','泰國清關']},
    {'id': 5, 'name': '訂艙與物流', 'icon': '🚢', 'steps': ['預估','櫃型','詢價','比價','船期']},
    {'id': 6, 'name': '稅費核算', 'icon': '💰', 'steps': ['查詢','計算','FTA比對','報表']},
    {'id': 7, 'name': 'ERP錄入', 'icon': '💻', 'steps': ['數據驗證','採購訂單','銷售訂單','料號主數據','產品主數據','批量錄入']},
    {'id': 8, 'name': '印尼協同', 'icon': '🇮🇩', 'steps': ['進度','出貨通知','清關同步','日報週報','翻譯']},
    {'id': 9, 'name': '資料歸檔', 'icon': '🗃️', 'steps': ['歸檔','微盤','報告']},
    {'id': 10, 'name': '品牌授權', 'icon': '®️', 'steps': ['預警','總覽','續約']},
    {'id': 11, 'name': '政策監控', 'icon': '📡', 'steps': ['泰國政策','印尼政策','綜合動態']},
    {'id': 12, 'name': '每日運營', 'icon': '📊', 'steps': ['進度總覽','異常追蹤','KPI看板','日誌']},
]

def execute_tool(tool_name, args):
    filename = TOOL_FILES.get(tool_name)
    if not filename:
        return {'success': False, 'error': f'未知工具: {tool_name}'}
    filepath = os.path.join(BASE_DIR, filename)
    if not os.path.exists(filepath):
        return {'success': False, 'error': f'檔案不存在: {filename}'}
    try:
        cmd = [sys.executable, filepath] + args.split() if args else [sys.executable, filepath]
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=120, cwd=BASE_DIR)
        stdout = result.stdout
        if len(stdout) > 4000:
            stdout = '...(截斷)...\n' + stdout[-4000:]
        return {
            'success': result.returncode == 0,
            'stdout': stdout,
            'stderr': result.stderr[-1000:] if result.stderr else '',
            'returncode': result.returncode
        }
    except subprocess.TimeoutExpired:
        return {'success': False, 'error': '執行逾時 (120秒)'}
    except Exception as e:
        return {'success': False, 'error': str(e)}

def log_history(ticket_id, action, detail=''):
    path = os.path.join(DATA_DIR, 'history.json')
    history = []
    if os.path.exists(path):
        with open(path, 'r', encoding='utf-8') as f:
            history = json.load(f)
    history.append({
        'time': datetime.now().isoformat(),
        'ticket': ticket_id,
        'action': action,
        'detail': detail
    })
    if len(history) > 5000:
        history = history[-5000:]
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(history, f, ensure_ascii=False, indent=2)


# ─────────────────────── ERP Connector Functions ───────────────────────

ERP_CONNECTOR = '用友ERP連接器.py'

def erp_run(args_list):
    """Run the ERP connector with given arguments"""
    filepath = os.path.join(BASE_DIR, ERP_CONNECTOR)
    if not os.path.exists(filepath):
        return {'success': False, 'error': f'ERP連接器不存在: {ERP_CONNECTOR}'}
    try:
        cmd = [sys.executable, filepath] + args_list
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=60, cwd=BASE_DIR)
        stdout = result.stdout[-3000:] if len(result.stdout) > 3000 else result.stdout
        return {'success': result.returncode == 0, 'stdout': stdout,
                'stderr': result.stderr[-500:] if result.stderr else '', 'returncode': result.returncode}
    except subprocess.TimeoutExpired:
        return {'success': False, 'error': 'ERP操作逾時(60秒)'}
    except Exception as e:
        return {'success': False, 'error': str(e)}

def get_erp_status():
    """Get ERP connection and sync status"""
    config_path = os.path.join(DATA_DIR, 'erp_config.json')
    config = {}
    if os.path.exists(config_path):
        with open(config_path, 'r', encoding='utf-8') as f:
            config = json.load(f)
    cache_db = os.path.join(DATA_DIR, 'erp_cache.db')
    cache_exists = os.path.exists(cache_db)
    # Count records in cache if DB exists
    cache_stats = {}
    if cache_exists:
        try:
            import sqlite3
            conn = sqlite3.connect(cache_db)
            cur = conn.cursor()
            for table in ['erp_purchase_orders','erp_materials','erp_suppliers','erp_inventory','erp_invoices','erp_customers','erp_sales_orders']:
                try:
                    cur.execute(f'SELECT COUNT(*) FROM {table}')
                    cache_stats[table.replace('erp_','')] = {'count': cur.fetchone()[0]}
                except:
                    cache_stats[table.replace('erp_','')] = {'count': 0}
            conn.close()
        except:
            pass
    pending = 0
    pending_dir = os.path.join(DATA_DIR, 'erp_queue')
    if os.path.exists(pending_dir):
        pending = len([f for f in os.listdir(pending_dir) if f.endswith('.json')])
    return {
        'connected': True,
        'mode': config.get('mode', 'mock'),
        'url': config.get('base_url', ''),
        'last_sync': config.get('last_sync', None),
        'pending_queue': pending,
        'unresolved_conflicts': 0,
        'cache_stats': cache_stats,
        'sync_history': []
    }

def erp_pull(data_type=None):
    """Pull data from ERP"""
    args = ['--pull']
    if data_type:
        args += ['--pull-type', data_type]
    result = erp_run(args)
    if result.get('success'):
        # Refresh status to get updated cache counts
        status = get_erp_status()
        result['cache_stats'] = status['cache_stats']
    return result

def erp_push():
    """Push pending operations to ERP"""
    return erp_run(['--push'])

def erp_sync():
    """Full bidirectional sync"""
    return erp_run(['--sync'])

def erp_search(search_type, keyword):
    """Search cached ERP data"""
    return erp_run(['--search', search_type, keyword])

def erp_test():
    """Test ERP connection"""
    return erp_run(['--test'])


# ─────────────────────── HTTP Request Handler ───────────────────────

class PlatformHandler(http.server.BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        pass  # suppress default logging

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path.rstrip('/')

        if path == '' or path == '/':
            self.serve_html()
        elif path == '/api/tickets':
            self.json_response(load_tickets())
        elif path.startswith('/api/ticket/'):
            ticket_id = path.split('/')[-1]
            self.json_response(load_ticket(ticket_id))
        elif path.startswith('/api/step_data/'):
            parts = path.strip('/').split('/')
            if len(parts) >= 4:
                ticket_id, step_id = parts[2], parts[3]
                self.json_response(load_step_data(ticket_id, step_id))
            else:
                self.send_error(400)
        elif path == '/api/status':
            self.json_response(get_system_status())
        elif path == '/api/tools':
            self.json_response(get_tool_list())
        elif path == '/api/workflows':
            self.json_response(get_workflows())
        elif path == '/api/history':
            p = os.path.join(DATA_DIR, 'history.json')
            if os.path.exists(p):
                with open(p, 'r', encoding='utf-8') as f:
                    self.json_response(json.load(f))
            else:
                self.json_response([])
        elif path.startswith('/api/execute'):
            params = urllib.parse.parse_qs(parsed.query)
            tool = params.get('tool', [''])[0]
            args = params.get('args', [''])[0]
            result = execute_tool(tool, args)
            log_history(params.get('ticket', [''])[0], 'execute_tool', f'{tool} {args}')
            self.json_response(result)
        elif path.startswith('/api/erp/'):
            self.handle_erp_get(path, parsed)
        elif path.startswith('/uploads/') or path.startswith('/outputs/'):
            self.serve_file(os.path.join(DATA_DIR, path.lstrip('/')))
        else:
            self.send_error(404)

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path.rstrip('/')

        if path == '/api/ticket/create':
            data = self.read_json_body()
            result = create_ticket(data)
            log_history(result['id'], 'create_ticket', result.get('name', ''))
            self.json_response(result)
        elif path.startswith('/api/ticket/') and path.endswith('/update'):
            ticket_id = path.split('/')[-2]
            data = self.read_json_body()
            result = update_ticket(ticket_id, data)
            self.json_response(result)
        elif path.startswith('/api/step_data/'):
            parts = path.strip('/').split('/')
            if len(parts) >= 4:
                ticket_id, step_id = parts[2], parts[3]
                data = self.read_json_body()
                save_step_data(ticket_id, step_id, data)
                log_history(ticket_id, 'save_step', step_id)
                self.json_response({'success': True})
            else:
                self.send_error(400)
        elif path.startswith('/api/upload/'):
            parts = path.strip('/').split('/')
            if len(parts) >= 4:
                ticket_id, step_id = parts[2], parts[3]
                result = self.handle_file_upload(ticket_id, step_id)
                log_history(ticket_id, 'upload_file', step_id)
                self.json_response(result)
            else:
                self.send_error(400)
        elif path.startswith('/api/step/') and path.endswith('/complete'):
            parts = path.strip('/').split('/')
            if len(parts) >= 4:
                ticket_id, step_id = parts[2], parts[3]
                mark_step_complete(ticket_id, step_id)
                log_history(ticket_id, 'complete_step', step_id)
                self.json_response({'success': True})
            else:
                self.send_error(400)
        elif path == '/api/ticket/delete':
            data = self.read_json_body()
            tid = data.get('id', '')
            tickets = load_tickets()
            tickets['tickets'] = [t for t in tickets['tickets'] if t['id'] != tid]
            save_tickets(tickets)
            self.json_response({'success': True})
        else:
            self.send_error(404)

    def read_json_body(self):
        length = int(self.headers.get('Content-Length', 0))
        body = self.rfile.read(length)
        return json.loads(body.decode('utf-8'))

    def json_response(self, data):
        body = json.dumps(data, ensure_ascii=False, indent=2).encode('utf-8')
        self.send_response(200)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def serve_html(self):
        html = DASHBOARD_HTML.encode('utf-8')
        self.send_response(200)
        self.send_header('Content-Type', 'text/html; charset=utf-8')
        self.send_header('Content-Length', str(len(html)))
        self.end_headers()
        self.wfile.write(html)

    def serve_file(self, filepath):
        if os.path.exists(filepath):
            self.send_response(200)
            mime, _ = mimetypes.guess_type(filepath)
            self.send_header('Content-Type', mime or 'application/octet-stream')
            self.end_headers()
            with open(filepath, 'rb') as f:
                shutil.copyfileobj(f, self.wfile)
        else:
            self.send_error(404)

    def handle_file_upload(self, ticket_id, step_id):
        content_type = self.headers.get('Content-Type', '')
        if 'multipart/form-data' not in content_type:
            return {'success': False, 'error': 'Expected multipart/form-data'}
        upload_dir = os.path.join(DATA_DIR, 'uploads', ticket_id)
        os.makedirs(upload_dir, exist_ok=True)
        try:
            form = cgi.FieldStorage(
                fp=self.rfile,
                headers=self.headers,
                environ={'REQUEST_METHOD': 'POST', 'CONTENT_TYPE': content_type}
            )
            uploaded = []
            for key in form.keys():
                item = form[key]
                if item.filename:
                    safe_name = f"{step_id}_{item.filename}"
                    filepath = os.path.join(upload_dir, safe_name)
                    with open(filepath, 'wb') as f:
                        f.write(item.file.read())
                    uploaded.append({'name': safe_name, 'path': f'/uploads/{ticket_id}/{safe_name}'})
            return {'success': True, 'files': uploaded}
        except Exception as e:
            return {'success': False, 'error': str(e)}


# ─────────────────────── Threaded Server ───────────────────────

class ThreadedHTTPServer(ThreadingMixIn, HTTPServer):
    daemon_threads = True


# ═══════════════════════════════════════════════════════════════════════
#                         HTML DASHBOARD
# ═══════════════════════════════════════════════════════════════════════

DASHBOARD_HTML = r"""<!DOCTYPE html>
<html lang="zh-Hant">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>關務工作平台 — JOSUN</title>
<style>
@import url('https://fonts.googleapis.com/css2?family=Nunito+Sans:wght@400;500;600;700&family=Noto+Sans+TC:wght@400;500;600;700&display=swap');

/* ── Reset & Base ── */
*{margin:0;padding:0;box-sizing:border-box}
:root{
  --bg:#F5F7F5; --card:#FFFFFF; --card2:#F8F9F8;
  --border:#E0E5E0; --text:#2D3748; --text2:#6B7280;
  --accent:#77AE40; --accent2:#77AE40; --green:#77AE40;
  --yellow:#1F492B; --red:#EF4444; --orange:#f97316;
  --sidebar:#FFFFFF; --hover:#BDEAD4;
  --josun-green:#77AE40; --dark-green:#1F492B;
  --light-blue:#9ADCF7; --light-teal:#BDEAD4;
}
html,body{height:100%;font-family:'Nunito Sans','Noto Sans TC',Arial,sans-serif;
  background:var(--bg);color:var(--text);font-size:14px;overflow:hidden}
a{color:var(--accent);text-decoration:none}
::-webkit-scrollbar{width:6px}
::-webkit-scrollbar-track{background:var(--bg)}
::-webkit-scrollbar-thumb{background:#C5CCC5;border-radius:3px}

/* ── Top Bar ── */
.topbar{display:flex;align-items:center;height:52px;background:#FFFFFF;
  border-bottom:2px solid var(--josun-green);padding:0 20px;gap:14px;z-index:100}
.topbar .logo{font-size:18px;font-weight:700;color:var(--josun-green);white-space:nowrap;
  font-family:'Nunito Sans','Noto Sans TC',sans-serif;display:flex;align-items:center;gap:4px}
.topbar .logo .logo-circle{font-size:22px;line-height:1}
.topbar .logo .logo-text{font-weight:700;letter-spacing:1px}
.topbar .logo .logo-sep{color:var(--border);margin:0 6px;font-weight:300}
.topbar .logo .logo-sub{color:var(--dark-green);font-weight:500;font-size:14px}
.topbar .logo .logo-ver{color:var(--text2);font-weight:400;font-size:11px;margin-left:4px}
.ticket-select{background:#FFFFFF;color:var(--text);border:1px solid var(--border);
  border-radius:6px;padding:6px 12px;font-size:13px;min-width:260px;max-width:400px;
  font-family:'Nunito Sans','Noto Sans TC',sans-serif;transition:border-color .2s}
.ticket-select:focus{outline:none;border-color:var(--josun-green);
  box-shadow:0 0 0 3px rgba(119,174,64,.12)}
.topbar .spacer{flex:1}
.topbar .clock{color:var(--text2);font-size:12px;font-family:'Nunito Sans',monospace}
.topbar .status-dot{width:8px;height:8px;border-radius:50%;background:var(--green);display:inline-block;margin-right:4px}
.topbar .status-label{font-size:12px;color:var(--text2);margin-right:8px}

/* ── Layout ── */
.main-layout{display:flex;height:calc(100vh - 88px)}

/* ── Sidebar ── */
.sidebar{width:210px;min-width:210px;background:#FFFFFF;border-right:1px solid var(--border);
  overflow-y:auto;padding:8px 0}
.sidebar .mod-item{display:flex;align-items:center;padding:10px 14px;cursor:pointer;
  border-left:3px solid transparent;transition:all .15s;gap:8px;font-size:13px;color:var(--text)}
.sidebar .mod-item:hover{background:rgba(189,234,212,.35)}
.sidebar .mod-item.active{background:rgba(189,234,212,.25);border-left-color:var(--josun-green);
  color:var(--dark-green);font-weight:600}
.sidebar .mod-item .icon{font-size:16px;width:22px;text-align:center}
.sidebar .mod-item .progress-mini{margin-left:auto;font-size:11px;color:var(--josun-green);font-weight:600}
.sidebar .section-label{padding:14px 14px 6px;font-size:10px;text-transform:uppercase;
  color:var(--dark-green);letter-spacing:1.5px;font-weight:600}

/* ── Content Area ── */
.content{flex:1;overflow-y:auto;padding:20px 24px;background:var(--bg)}

/* ── Ticket Info Bar ── */
.ticket-bar{background:#FFFFFF;border:1px solid var(--border);border-left:4px solid var(--josun-green);
  border-radius:8px;padding:14px 18px;margin-bottom:18px;display:flex;align-items:center;gap:16px;
  flex-wrap:wrap;box-shadow:0 1px 3px rgba(0,0,0,.04)}
.ticket-bar .tb-title{font-size:16px;font-weight:700;color:var(--dark-green);
  font-family:'Nunito Sans','Noto Sans TC',sans-serif}
.ticket-bar .tb-meta{font-size:12px;color:var(--text2);margin-top:2px}
.ticket-bar .tb-progress{margin-left:auto;display:flex;align-items:center;gap:10px}
.progress-bar{width:160px;height:8px;background:#E8ECE8;border-radius:4px;overflow:hidden}
.progress-bar .fill{height:100%;background:var(--josun-green);border-radius:4px;transition:width .3s}
.ticket-bar .tb-pct{font-size:13px;font-weight:700;color:var(--josun-green)}

/* ── Step Card ── */
.step-card{background:#FFFFFF;border:1px solid var(--border);border-radius:8px;
  margin-bottom:12px;overflow:hidden;transition:all .2s;box-shadow:0 1px 2px rgba(0,0,0,.03)}
.step-card.completed{border-color:var(--josun-green);opacity:.85}
.step-header{display:flex;align-items:center;padding:11px 16px;cursor:pointer;
  background:rgba(189,234,212,.18);gap:10px;user-select:none;transition:background .15s}
.step-header:hover{background:rgba(189,234,212,.32)}
.step-num{background:var(--josun-green);color:#fff;border-radius:4px;padding:2px 9px;
  font-size:11px;font-weight:700;white-space:nowrap}
.step-name{font-weight:600;font-size:14px;color:var(--dark-green)}
.step-status{margin-left:auto;font-size:12px;white-space:nowrap}
.step-status.pending{color:var(--text2)}
.step-status.done{color:var(--josun-green)}
.step-body{padding:0 16px 16px;display:none}
.step-card.open .step-body{display:block}
.step-card .chevron{color:var(--text2);transition:transform .2s;font-size:12px}
.step-card.open .chevron{transform:rotate(90deg)}

/* ── Panels inside step ── */
.panel{margin-top:12px;padding:14px;background:var(--card2);border:1px solid var(--border);border-radius:6px}
.panel h4{font-size:13px;margin-bottom:8px;color:var(--dark-green);display:flex;align-items:center;gap:6px;
  font-family:'Nunito Sans','Noto Sans TC',sans-serif;font-weight:600}
.panel label{display:block;font-size:12px;color:var(--text2);margin-bottom:3px}

/* ── Form Elements ── */
.form-row{margin-bottom:10px}
.form-row label{display:block;font-size:12px;color:var(--dark-green);margin-bottom:4px;font-weight:500}
.form-row input[type="text"],.form-row input[type="number"],.form-row input[type="date"],
.form-row select,.form-row textarea{
  width:100%;background:#FFFFFF;border:1px solid var(--border);border-radius:6px;
  color:var(--text);padding:7px 11px;font-size:13px;font-family:'Nunito Sans','Noto Sans TC',sans-serif;
  transition:border-color .2s,box-shadow .2s}
.form-row textarea{min-height:60px;resize:vertical}
.form-row input:focus,.form-row select:focus,.form-row textarea:focus{
  outline:none;border-color:var(--josun-green);box-shadow:0 0 0 3px rgba(119,174,64,.1)}
.form-grid{display:grid;grid-template-columns:1fr 1fr;gap:10px}
.form-grid-3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px}
.checkbox-row{display:flex;align-items:center;gap:6px;margin-bottom:6px;font-size:13px}
.checkbox-row input[type="checkbox"]{accent-color:var(--josun-green)}

/* ── Buttons ── */
.btn{display:inline-flex;align-items:center;gap:6px;padding:7px 16px;border:none;
  border-radius:6px;cursor:pointer;font-size:13px;font-family:'Nunito Sans','Noto Sans TC',sans-serif;
  transition:all .15s;font-weight:500}
.btn-save{background:var(--josun-green);color:#fff}
.btn-save:hover{background:#68993A}
.btn-execute{background:var(--josun-green);color:#fff}
.btn-execute:hover{background:#68993A}
.btn-sm{padding:5px 12px;font-size:12px}
.btn-danger{background:var(--red);color:#fff}
.btn-danger:hover{background:#dc2626}
.btn-outline{background:#FFFFFF;border:1px solid var(--josun-green);color:var(--josun-green)}
.btn-outline:hover{background:rgba(119,174,64,.06);border-color:#68993A;color:#68993A}

/* ── File Upload ── */
.upload-area{border:2px dashed var(--border);border-radius:6px;padding:18px;text-align:center;
  color:var(--text2);cursor:pointer;transition:border-color .2s,background .2s;background:#FFFFFF}
.upload-area:hover{border-color:var(--josun-green);background:rgba(189,234,212,.1)}
.upload-area input[type="file"]{display:none}
.file-list{margin-top:8px}
.file-item{display:flex;align-items:center;gap:8px;padding:6px 10px;background:#FFFFFF;
  border:1px solid var(--border);border-radius:6px;margin-bottom:4px;font-size:12px}
.file-item a{color:var(--josun-green)}
.file-item::before{content:'';display:inline-block;width:6px;height:6px;border-radius:50%;background:var(--josun-green)}

/* ── Terminal Output ── */
.terminal{background:var(--dark-green);border:1px solid rgba(189,234,212,.2);border-radius:6px;padding:12px;
  font-family:'Cascadia Code','Fira Code',monospace;font-size:12px;color:var(--light-teal);
  max-height:300px;overflow-y:auto;white-space:pre-wrap;word-break:break-all;margin-top:8px}

/* ── Command Display ── */
.cmd-display{background:var(--card2);border:1px solid var(--border);border-radius:6px;
  padding:8px 12px;font-family:'Cascadia Code','Fira Code',monospace;font-size:12px;
  color:var(--dark-green);margin-bottom:8px;font-weight:500}

/* ── Dynamic Rows ── */
.dyn-row{display:flex;gap:8px;align-items:flex-end;margin-bottom:6px;padding:10px;
  background:#FFFFFF;border:1px solid var(--border);border-radius:6px;flex-wrap:wrap}
.dyn-row .form-row{flex:1;min-width:100px;margin-bottom:0}
.dyn-row .btn-remove{background:var(--red);color:#fff;border:none;border-radius:4px;
  padding:4px 8px;cursor:pointer;font-size:12px;align-self:flex-end}

/* ── Notes ── */
.notes-area{width:100%;min-height:50px;background:#FFFFFF;border:1px solid var(--border);
  border-radius:6px;color:var(--text);padding:8px 11px;font-size:13px;
  font-family:'Nunito Sans','Noto Sans TC',sans-serif;resize:vertical;transition:border-color .2s}
.notes-area:focus{outline:none;border-color:var(--josun-green);box-shadow:0 0 0 3px rgba(119,174,64,.1)}

/* ── Complete Panel ── */
.complete-panel{display:flex;align-items:center;gap:12px;margin-top:12px;padding-top:12px;
  border-top:2px solid var(--light-teal)}
.complete-panel .complete-time{font-size:12px;color:var(--text2)}

/* ── Modal ── */
.modal-overlay{position:fixed;inset:0;background:rgba(0,0,0,.35);z-index:1000;
  display:none;align-items:center;justify-content:center}
.modal-overlay.show{display:flex}
.modal{background:#FFFFFF;border:1px solid var(--border);border-radius:12px;
  padding:0;width:500px;max-width:90vw;max-height:80vh;overflow-y:auto;
  box-shadow:0 8px 30px rgba(0,0,0,.12)}
.modal .modal-accent{height:4px;background:var(--josun-green);border-radius:12px 12px 0 0}
.modal .modal-body{padding:24px}
.modal h2{font-size:18px;margin-bottom:16px;color:var(--dark-green);
  font-family:'Nunito Sans','Noto Sans TC',sans-serif;font-weight:700}
.modal .btn-row{display:flex;gap:8px;justify-content:flex-end;margin-top:16px}

/* ── Bottom Bar ── */
.bottombar{height:40px;background:#FFFFFF;border-top:1px solid var(--border);
  display:flex;align-items:center;padding:0 16px;gap:12px;font-size:12px;color:var(--text2)}
.bottombar .btn{padding:4px 10px;font-size:11px}

/* ── Toast ── */
.toast{position:fixed;top:64px;right:20px;background:var(--josun-green);color:#fff;
  padding:10px 22px;border-radius:6px;font-size:13px;z-index:2000;
  transform:translateX(120%);transition:transform .3s;font-weight:500;
  box-shadow:0 4px 12px rgba(119,174,64,.3)}
.toast.show{transform:translateX(0)}
.toast.error{background:var(--red);box-shadow:0 4px 12px rgba(239,68,68,.3)}

/* ── Module Title ── */
.module-title{font-size:18px;font-weight:700;margin-bottom:4px;display:flex;align-items:center;gap:8px;
  color:var(--dark-green);font-family:'Nunito Sans','Noto Sans TC',sans-serif}
.module-subtitle{font-size:12px;color:var(--text2);margin-bottom:16px}

/* ── Empty State ── */
.empty-state{text-align:center;padding:80px 20px;color:var(--text2)}
.empty-state .big-icon{font-size:48px;margin-bottom:14px}
.empty-state p{font-size:14px;margin-bottom:18px;color:var(--text2)}

/* ── Execute Panel ── */
.execute-panel{background:var(--card2)}

/* ── Print ── */
@media print{
  .sidebar,.topbar,.bottombar,.btn,.upload-area,.execute-panel{display:none!important}
  .step-body{display:block!important}
  .step-card{break-inside:avoid;border:1px solid #ccc;box-shadow:none}
  body{background:#fff;color:#000}
  .ticket-bar{box-shadow:none;border:1px solid #ccc}
  .content{background:#fff}
}

/* Notification Center */
.notif-bell{position:relative;background:none;border:none;font-size:20px;cursor:pointer;padding:6px 8px;border-radius:8px;transition:background .2s}
.notif-bell:hover{background:rgba(119,174,64,.1)}
.notif-badge{position:absolute;top:2px;right:2px;background:#EF4444;color:#fff;font-size:10px;font-weight:700;min-width:16px;height:16px;border-radius:8px;display:flex;align-items:center;justify-content:center;padding:0 4px}
.notif-overlay{position:fixed;inset:0;background:rgba(0,0,0,.2);z-index:900;display:none}
.notif-overlay.show{display:block}
.notif-panel{position:fixed;top:52px;right:16px;width:420px;max-height:calc(100vh - 80px);background:#fff;border:1px solid var(--border,#e0e0e0);border-radius:12px;box-shadow:0 8px 32px rgba(0,0,0,.12);z-index:901;display:none;flex-direction:column;overflow:hidden}
.notif-panel.show{display:flex}
.notif-header{display:flex;align-items:center;justify-content:space-between;padding:16px 20px;border-bottom:1px solid #eee}
.notif-header h3{margin:0;font-size:16px;color:#1F492B;font-weight:700;font-family:'Nunito Sans','Noto Sans TC',sans-serif}
.notif-actions{display:flex;gap:6px}
.notif-tabs{display:flex;border-bottom:1px solid #eee;padding:0 12px}
.notif-tab{padding:10px 14px;font-size:13px;color:#666;cursor:pointer;border:none;background:none;border-bottom:2px solid transparent;font-family:'Nunito Sans','Noto Sans TC',sans-serif;transition:all .2s}
.notif-tab:hover{color:#1F492B}
.notif-tab.active{color:#77AE40;border-bottom-color:#77AE40;font-weight:600}
.notif-list{flex:1;overflow-y:auto;padding:8px 0;max-height:450px}
.notif-empty{text-align:center;color:#999;padding:40px 20px;font-size:14px}
.notif-item{padding:12px 20px;border-bottom:1px solid #f5f5f5;cursor:pointer;transition:background .15s}
.notif-item:hover{background:#f9faf9}
.notif-item.unread{background:#f0f7e8}
.notif-item-head{display:flex;align-items:center;gap:8px;margin-bottom:4px}
.notif-icon{font-size:16px;flex-shrink:0}
.notif-title{font-size:13px;font-weight:600;color:#1F492B;flex:1}
.notif-time{font-size:11px;color:#999;flex-shrink:0}
.notif-detail{font-size:12px;color:#666;line-height:1.5;display:none;padding-top:6px;border-top:1px dashed #eee;margin-top:6px}
.notif-item.expanded .notif-detail{display:block}
.notif-tag{display:inline-block;font-size:10px;padding:2px 6px;border-radius:4px;font-weight:600}
.notif-tag.tool{background:#E8F5E9;color:#2E7D32}
.notif-tag.step{background:#E3F2FD;color:#1565C0}
.notif-tag.alert{background:#FFF3E0;color:#E65100}
.notif-tag.success{background:#E8F5E9;color:#2E7D32}
.notif-tag.error{background:#FFEBEE;color:#C62828}
.notif-footer{padding:10px 20px;border-top:1px solid #eee;text-align:center}
.notif-footer span{font-size:12px;color:#999}
.notif-item .notif-expand-hint{font-size:10px;color:#bbb;margin-top:2px}
.notif-item.expanded .notif-expand-hint{display:none}
</style>
</head>
<body>

<!-- ═══ Top Bar ═══ -->
<div class="topbar">
  <div class="logo">
    <span class="logo-circle">&#9679;</span>
    <span class="logo-text">JOSUN</span>
    <span class="logo-sep">|</span>
    <span class="logo-sub">關務工作平台</span>
    <span class="logo-ver">v3.0</span>
  </div>
  <select class="ticket-select" id="ticketSelect" onchange="selectTicket(this.value)">
    <option value="">— 選擇工作票 —</option>
  </select>
  <button class="btn btn-sm btn-save" onclick="showNewTicketModal()">+ 新建工作票</button>
  <div class="spacer"></div>
  <button class="notif-bell" id="notifBell" onclick="toggleNotificationCenter()">
    &#128276;
    <span class="notif-badge" id="notifBadge" style="display:none">0</span>
  </button>
  <span class="status-label"><span class="status-dot"></span>系統運行中</span>
  <span class="clock" id="clock"></span>
</div>

<!-- ═══ Notification Center ═══ -->
<div class="notif-overlay" id="notifOverlay" onclick="closeNotificationCenter()"></div>
<div class="notif-panel" id="notifPanel">
  <div class="notif-header">
    <h3>&#128276; 通知中心</h3>
    <div class="notif-actions">
      <button class="btn btn-sm btn-outline" onclick="markAllNotifRead()">全部已讀</button>
      <button class="btn btn-sm btn-outline" onclick="clearAllNotif()">清除全部</button>
    </div>
  </div>
  <div class="notif-tabs">
    <button class="notif-tab active" onclick="switchNotifTab('all',this)">全部</button>
    <button class="notif-tab" onclick="switchNotifTab('tool',this)">工具執行</button>
    <button class="notif-tab" onclick="switchNotifTab('step',this)">步驟完成</button>
    <button class="notif-tab" onclick="switchNotifTab('alert',this)">提醒</button>
  </div>
  <div class="notif-list" id="notifList">
    <div class="notif-empty">尚無通知</div>
  </div>
  <div class="notif-footer">
    <span id="notifCount">0 則通知</span>
  </div>
</div>

<!-- ═══ Main Layout ═══ -->
<div class="main-layout">

  <!-- Sidebar -->
  <div class="sidebar" id="sidebar"></div>

  <!-- Content -->
  <div class="content" id="content">
    <div class="empty-state" id="emptyState">
      <div class="big-icon">&#128203;</div>
      <p>請選擇或建立工作票以開始作業</p>
      <button class="btn btn-save" onclick="showNewTicketModal()">+ 建立新工作票</button>
    </div>
    <div id="ticketContent" style="display:none"></div>
  </div>
</div>

<!-- ═══ Bottom Bar ═══ -->
<div class="bottombar">
  <button class="btn btn-sm btn-outline" onclick="saveAllStepsInModule()">&#128190; 全部儲存 (Ctrl+S)</button>
  <button class="btn btn-sm btn-outline" onclick="expandAllSteps()">&#9660; 展開全部</button>
  <button class="btn btn-sm btn-outline" onclick="collapseAllSteps()">&#9650; 收合全部</button>
  <button class="btn btn-sm btn-outline" onclick="nextIncompleteStep()">&#10145; 下一步未完成</button>
  <div style="flex:1"></div>
  <button class="btn btn-sm btn-outline" onclick="exportProgressReport()">&#128202; 匯出報告</button>
  <button class="btn btn-sm btn-outline" onclick="exportReport()">&#128424; 列印</button>
  <span id="bottomTicketInfo"></span>
</div>

<!-- ═══ New Ticket Modal ═══ -->
<div class="modal-overlay" id="newTicketModal">
  <div class="modal">
    <div class="modal-accent"></div>
    <div class="modal-body">
    <h2>建立新工作票</h2>
    <div class="form-row">
      <label>工作票名稱</label>
      <input type="text" id="nt_name" placeholder="例：PO-2026-TH-0055 筆記型電腦出口">
    </div>
    <div class="form-grid">
      <div class="form-row">
        <label>PO / 合同號</label>
        <input type="text" id="nt_po" placeholder="PO-2026-0055">
      </div>
      <div class="form-row">
        <label>類型</label>
        <select id="nt_type">
          <option value="export">出口</option>
          <option value="import">進口</option>
        </select>
      </div>
    </div>
    <div class="form-row">
      <label>目的國</label>
      <select id="nt_country">
        <option value="thailand">泰國</option>
        <option value="indonesia">印尼</option>
      </select>
    </div>
    <div class="form-row">
      <label>備註</label>
      <textarea id="nt_notes" rows="3" placeholder="選填備註..."></textarea>
    </div>
    <div class="btn-row">
      <button class="btn btn-outline" onclick="closeModal('newTicketModal')">取消</button>
      <button class="btn btn-save" onclick="createTicket()">建立</button>
    </div>
    </div>
  </div>
</div>

<!-- Toast -->
<div class="toast" id="toast"></div>

<script>
// ═══════════════════════ State ═══════════════════════
let currentTicket = null;
let currentModule = 1;
let allTickets = [];
let ticketData = {};  // cached step data

// ═══════════════════════ Module Definitions ═══════════════════════
const MODULES = [
  {id:1, name:'出貨準備', icon:'📦', steps:['ERP導出','確認品項','缺貨確認']},
  {id:2, name:'包裝與標識管理', icon:'📐', steps:['IPPC檢查','散件並箱','稱重','裝箱清單','紙箱嘜頭','櫃型核算','托盤嘜頭','PACKING LIST']},
  {id:3, name:'報關資料準備', icon:'📋', steps:['HS歸類','型錄收集','工程師確認','申報要素','代理核驗','FTA確認']},
  {id:4, name:'單據製作與審核', icon:'📄', steps:['全套單據','校驗','提單確認','泰國清關']},
  {id:5, name:'訂艙與物流', icon:'🚢', steps:['預估','櫃型','詢價','比價','船期']},
  {id:6, name:'稅費核算', icon:'💰', steps:['查詢','計算','FTA比對','報表']},
  {id:7, name:'ERP錄入', icon:'💻', steps:['數據驗證','採購訂單','銷售訂單','料號主數據','產品主數據','批量錄入']},
  {id:8, name:'印尼協同', icon:'🇮🇩', steps:['進度查詢','出貨通知','清關同步','日報週報','翻譯']},
  {id:9, name:'資料歸檔', icon:'🗃️', steps:['歸檔','微盤','報告']},
  {id:10, name:'品牌授權', icon:'®', steps:['預警','總覽','續約']},
  {id:11, name:'政策監控', icon:'📡', steps:['泰國政策','印尼政策','綜合動態']},
  {id:12, name:'每日運營', icon:'📊', steps:['進度總覽','異常追蹤','KPI看板','日誌']},
];

const TOOL_MAP = {
  'm1s1':{tool:'hs-classify',args:'--export-po'},
  'm2s5':{tool:'shipping-mark',args:'--carton-mark'},
  'm2s7':{tool:'shipping-mark',args:'--pallet-mark'},
  'm2s8':{tool:'shipping-mark',args:'--packing-list'},
  'm3s1':{tool:'hs-classify',args:'--classify'},
  'm4s1':{tool:'doc-generator',args:'--full-set'},
  'm4s2':{tool:'doc-validator',args:'--check'},
  'm6s1':{tool:'tax-calc',args:'--query'},
  'm6s2':{tool:'tax-calc',args:'--calc'},
  'm7s1':{tool:'erp-import',args:'--validate'},
  'm7s6':{tool:'erp-import',args:'--batch'},
  'm8s1':{tool:'indonesia-collab',args:'--status'},
  'm8s2':{tool:'indonesia-collab',args:'--notify'},
  'm9s1':{tool:'data-archive',args:'--archive'},
  'm10s1':{tool:'brand-auth',args:'--alert'},
  'm10s2':{tool:'brand-auth',args:'--overview'},
  'm11s1':{tool:'policy-tracker',args:'--thailand'},
  'm11s2':{tool:'policy-tracker',args:'--indonesia'},
  'm11s3':{tool:'policy-tracker',args:'--all'},
};

// ═══════════════════════ Init ═══════════════════════
document.addEventListener('DOMContentLoaded', () => {
  updateClock();
  setInterval(updateClock, 1000);
  renderSidebar();
  loadAllTickets();
});

function updateClock() {
  const now = new Date();
  document.getElementById('clock').textContent =
    now.toLocaleDateString('zh-Hant') + ' ' + now.toLocaleTimeString('zh-Hant');
}

// ═══════════════════════ Sidebar ═══════════════════════
function renderSidebar() {
  const sb = document.getElementById('sidebar');
  let html = '<div class="section-label">作業模組</div>';
  MODULES.forEach(m => {
    const pct = getModuleProgress(m.id);
    html += `<div class="mod-item ${currentModule===m.id?'active':''}" onclick="switchModule(${m.id})">
      <span class="icon">${m.icon}</span>
      <span>${m.name}</span>
      <span class="progress-mini">${pct}%</span>
    </div>`;
  });
  sb.innerHTML = html;
}

function switchModule(modId) {
  currentModule = modId;
  renderSidebar();
  if (currentTicket) renderModuleContent(modId);
}

function getModuleProgress(modId) {
  if (!currentTicket) return 0;
  const m = MODULES.find(x => x.id === modId);
  if (!m) return 0;
  const completed = currentTicket.completed_steps || [];
  let done = 0;
  m.steps.forEach((_, i) => {
    if (completed.includes(`m${modId}s${i+1}`)) done++;
  });
  return Math.round((done / m.steps.length) * 100);
}

function getTotalProgress() {
  if (!currentTicket) return 0;
  const allSteps = [];
  MODULES.forEach(m => m.steps.forEach((_, i) => allSteps.push(`m${m.id}s${i+1}`)));
  const completed = currentTicket.completed_steps || [];
  const done = allSteps.filter(s => completed.includes(s)).length;
  return Math.round((done / allSteps.length) * 100);
}

// ═══════════════════════ Tickets ═══════════════════════
function loadAllTickets() {
  fetch('/api/tickets').then(r=>r.json()).then(data => {
    allTickets = data.tickets || [];
    const sel = document.getElementById('ticketSelect');
    sel.innerHTML = '<option value="">— 選擇工作票 —</option>';
    allTickets.forEach(t => {
      const pct = t.completed_steps ? t.completed_steps.length : 0;
      sel.innerHTML += `<option value="${t.id}">${t.id} | ${t.name} (${t.status})</option>`;
    });
    if (currentTicket) {
      sel.value = currentTicket.id;
    }
  });
}

function showNewTicketModal() {
  document.getElementById('newTicketModal').classList.add('show');
}

function closeModal(id) {
  document.getElementById(id).classList.remove('show');
}

function createTicket() {
  const data = {
    name: document.getElementById('nt_name').value,
    po_number: document.getElementById('nt_po').value,
    type: document.getElementById('nt_type').value,
    country: document.getElementById('nt_country').value,
    notes: document.getElementById('nt_notes').value
  };
  if (!data.name) { toast('請輸入工作票名稱', true); return; }
  fetch('/api/ticket/create', {
    method: 'POST',
    headers: {'Content-Type':'application/json'},
    body: JSON.stringify(data)
  }).then(r=>r.json()).then(t => {
    closeModal('newTicketModal');
    toast(`工作票 ${t.id} 已建立`);
    currentTicket = t;
    loadAllTickets();
    setTimeout(() => {
      document.getElementById('ticketSelect').value = t.id;
      selectTicket(t.id);
    }, 300);
  });
}

function selectTicket(ticketId) {
  if (!ticketId) {
    currentTicket = null;
    document.getElementById('emptyState').style.display = '';
    document.getElementById('ticketContent').style.display = 'none';
    document.getElementById('bottomTicketInfo').textContent = '';
    renderSidebar();
    return;
  }
  fetch('/api/ticket/' + ticketId).then(r=>r.json()).then(t => {
    currentTicket = t;
    ticketData = {};
    document.getElementById('emptyState').style.display = 'none';
    document.getElementById('ticketContent').style.display = '';
    document.getElementById('bottomTicketInfo').textContent = `${t.id} | ${t.name}`;
    renderSidebar();
    renderModuleContent(currentModule);
    const completedSteps = (t.completed_steps || []).length;
    const totalSteps = MODULES.reduce((acc, m) => acc + m.steps.length, 0);
    addNotification('alert', '已載入工作票: ' + t.name, '共 ' + completedSteps + '/' + totalSteps + ' 步驟已完成', 'alert');
  });
}

// ═══════════════════════ Module Rendering ═══════════════════════
function renderModuleContent(modId) {
  const m = MODULES.find(x => x.id === modId);
  if (!m || !currentTicket) return;
  const container = document.getElementById('ticketContent');
  const pct = getTotalProgress();

  let html = `
    <div class="ticket-bar">
      <div>
        <div class="tb-title">${currentTicket.name}</div>
        <div class="tb-meta">${currentTicket.id} | PO: ${currentTicket.po_number||'-'} |
          ${currentTicket.type==='export'?'出口':'進口'} |
          ${currentTicket.country==='thailand'?'泰國':'印尼'} |
          建立: ${currentTicket.created ? currentTicket.created.substring(0,10) : '-'}</div>
      </div>
      <div class="tb-progress">
        <div class="progress-bar"><div class="fill" style="width:${pct}%"></div></div>
        <span class="tb-pct">${pct}%</span>
      </div>
      <button class="btn btn-sm btn-outline" onclick="toggleTicketStatus('${currentTicket.id}')">${currentTicket.status==='active'?'標記完成':'重新開啟'}</button>
      <button class="btn btn-sm btn-danger" onclick="deleteTicket('${currentTicket.id}')">刪除</button>
    </div>
    <div class="module-title">${m.icon} ${m.name}</div>
    <div class="module-subtitle">模組 ${m.id} — 共 ${m.steps.length} 個步驟</div>
  `;

  m.steps.forEach((stepName, i) => {
    const stepId = `m${modId}s${i+1}`;
    const isDone = (currentTicket.completed_steps || []).includes(stepId);
    html += renderStepCard(modId, i+1, stepId, stepName, isDone);
  });

  container.innerHTML = html;

  // Load saved data for each step
  m.steps.forEach((_, i) => {
    const stepId = `m${modId}s${i+1}`;
    loadStepDataToForm(stepId);
  });
}

function renderStepCard(modId, stepNum, stepId, stepName, isDone) {
  const statusClass = isDone ? 'done' : 'pending';
  const statusText = isDone ? '已完成' : '待處理';
  const cardClass = isDone ? 'completed' : '';
  const tool = TOOL_MAP[stepId];
  const cmdStr = tool ? `python3 ${TOOL_MAP[stepId] ? getToolFilename(tool.tool) : ''} ${tool.args}` : '';

  return `
  <div class="step-card ${cardClass}" id="card-${stepId}">
    <div class="step-header" onclick="toggleCard('${stepId}')">
      <span class="chevron">&#9654;</span>
      <span class="step-num">Step ${stepNum}</span>
      <span class="step-name">${stepName}</span>
      <span class="step-status ${statusClass}" id="status-${stepId}">${isDone ? '&#9989;' : '&#11036;'} ${statusText}</span>
    </div>
    <div class="step-body">
      ${getPrereqHTML(modId, stepNum)}
      ${getFormHTML(modId, stepNum, stepId)}
      ${getUploadHTML(stepId)}
      ${tool ? getExecuteHTML(stepId, tool, cmdStr) : ''}
      <div id="output-${stepId}" style="display:none">
        <div class="panel"><h4>&#128228; 執行結果</h4>
          <div class="terminal" id="term-${stepId}"></div>
        </div>
      </div>
      <div class="panel">
        <h4>&#128221; 備註</h4>
        <textarea class="notes-area" id="notes-${stepId}" placeholder="輸入備註..." onchange="saveNotes('${stepId}')"></textarea>
      </div>
      <div class="complete-panel">
        <label class="checkbox-row" style="margin:0">
          <input type="checkbox" id="chk-${stepId}" ${isDone?'checked':''} onchange="markComplete('${stepId}', this.checked)">
          <span>標記完成</span>
        </label>
        <span class="complete-time" id="ctime-${stepId}"></span>
      </div>
    </div>
  </div>`;
}

function getToolFilename(toolId) {
  const map = {
    'hs-classify':'HS編碼歸類助手.py','shipping-mark':'嘜頭裝箱單生成器.py',
    'doc-generator':'關務單據自動生成器.py','doc-validator':'報關單據校驗助手.py',
    'erp-import':'ERP批量錄入工具.py','tax-calc':'稅費核算引擎.py',
    'indonesia-collab':'印尼協同自動化工具.py','freight-booking':'訂艙比價工具.py',
    'brand-auth':'品牌授權管理系統.py','policy-tracker':'政策追蹤工具.py',
    'data-archive':'資料歸檔工具.py'
  };
  return map[toolId] || '';
}

function toggleCard(stepId) {
  const card = document.getElementById('card-' + stepId);
  card.classList.toggle('open');
}

// ═══════════════════════ Prerequisites ═══════════════════════
function getPrereqHTML(modId, stepNum) {
  const prereqs = {
    '1-1': ['採購訂單/合同','ERP帳號權限'],
    '1-2': ['ERP導出品項清單','供應商回覆確認'],
    '1-3': ['品項清單','供應商交期回覆'],
    '2-1': ['包裝材料清單','IPPC認證供應商資料'],
    '2-2': ['品項裝箱規劃','散件清單'],
    '2-3': ['電子秤校正','各箱品項明細'],
    '2-4': ['稱重記錄','箱號清單'],
    '2-5': ['中英文品名','箱號及數量'],
    '2-6': ['總體積(CBM)','總毛重','托盤數'],
    '2-7': ['托盤編號','各托盤內容物'],
    '2-8': ['發票號','裝箱清單數據'],
    '3-1': ['產品中英文型錄','產品照片（含型號/品牌/產地銘牌）','採購訂單/合同'],
    '3-2': ['產品型錄','規格書','產品照片'],
    '3-3': ['HS歸類結果','工程師聯絡資訊'],
    '3-4': ['HS歸類結果','申報要素模板'],
    '3-5': ['報關代理聯絡資訊','全套報關資料'],
    '3-6': ['HS編碼','FTA原產地證明'],
    '4-1': ['合同號','PO號','發票號','前模組數據'],
    '4-2': ['全套單據檔案'],
    '4-3': ['提單草稿','船公司資訊'],
    '4-4': ['清關代理資訊','清關文件清單'],
    '5-1': ['總毛重','總體積','總箱數'],
    '5-2': ['包裝模組數據'],
    '5-3': ['起運港','目的港','預計出貨日'],
    '5-4': ['各家貨代報價'],
    '5-5': ['選定貨代','船期確認'],
    '6-1': ['HS編碼','進口國'],
    '6-2': ['CIF價值','FOB價格'],
    '6-3': ['HS編碼','FTA协定'],
    '6-4': ['稅費計算結果'],
    '7-1': ['ERP模板檔案'],
    '8-1': ['訂單號','印尼合作方資訊'],
    '9-1': ['合同號','所有文件'],
    '10-1': ['品牌清單'],
    '11-1': ['目標國家'],
  };
  const key = `${modId}-${stepNum}`;
  const items = prereqs[key];
  if (!items) return '';
  let html = '<div class="panel"><h4>&#128203; 資料準備清單</h4>';
  items.forEach(p => {
    html += `<label class="checkbox-row"><input type="checkbox"> ${p}</label>`;
  });
  html += '</div>';
  return html;
}

// ═══════════════════════ Form Definitions ═══════════════════════
function getFormHTML(modId, stepNum, stepId) {
  const forms = {
    // Module 1: 出貨準備
    '1-1': `<div class="form-grid">
      <div class="form-row"><label>PO號碼</label><input type="text" name="po_number" data-step="${stepId}" placeholder="PO-2026-0055"></div>
      <div class="form-row"><label>供應商</label><select name="supplier" data-step="${stepId}"><option value="">請選擇</option><option>供應商A</option><option>供應商B</option><option>供應商C</option></select></div>
    </div>
    <div class="form-grid">
      <div class="form-row"><label>起始日期</label><input type="date" name="date_from" data-step="${stepId}"></div>
      <div class="form-row"><label>結束日期</label><input type="date" name="date_to" data-step="${stepId}"></div>
    </div>`,

    '1-2': `<div class="form-row"><label>品項清單</label><textarea name="item_list" data-step="${stepId}" rows="4" placeholder="品名 / 型號 / 數量&#10;筆記型電腦 / MBP-16 / 50&#10;充電器 / CHG-65W / 50"></textarea></div>
    <div class="form-row"><label>數量確認（總數）</label><input type="number" name="total_qty" data-step="${stepId}" placeholder="0"></div>`,

    '1-3': `<div class="form-row"><label>缺貨品項</label><textarea name="shortage_items" data-step="${stepId}" rows="3" placeholder="品名 / 型號 / 缺貨數量"></textarea></div>
    <div class="form-grid">
      <div class="form-row"><label>預計到貨日</label><input type="date" name="eta_date" data-step="${stepId}"></div>
      <div class="form-row"><label>處理方式</label><select name="handle_method" data-step="${stepId}"><option value="wait">等貨</option><option value="partial">分批出</option><option value="substitute">替代</option></select></div>
    </div>`,

    // Module 2: 包裝與標識管理
    '2-1': `<div class="form-grid-3">
      <div class="form-row"><label>托盤/木箱數量</label><input type="number" name="pallet_count" data-step="${stepId}" placeholder="0"></div>
      <div class="form-row"><label>有IPPC標識數</label><input type="number" name="ippc_yes" data-step="${stepId}" placeholder="0"></div>
      <div class="form-row"><label>無標識需更換數</label><input type="number" name="ippc_no" data-step="${stepId}" placeholder="0"></div>
    </div>
    <div class="form-row"><label>檢查結果</label><textarea name="inspect_result" data-step="${stepId}" rows="3" placeholder="記錄檢查結果..."></textarea></div>`,

    '2-2': `<div class="form-row"><label>箱號</label><input type="text" name="box_number" data-step="${stepId}" placeholder="BOX-001"></div>
    <div class="form-row"><label>箱內品項（品名/型號/數量）</label><textarea name="box_contents" data-step="${stepId}" rows="3" placeholder="品名 / 型號 / 數量"></textarea></div>
    <div class="form-row"><label>放置托盤號</label><input type="text" name="pallet_number" data-step="${stepId}" placeholder="PLT-01"></div>`,

    '2-3': `<div id="weightRows-${stepId}">
      <div class="dyn-row">
        <div class="form-row"><label>型號</label><input type="text" name="model" data-step="${stepId}" data-dyn="weight"></div>
        <div class="form-row"><label>單箱淨重(kg)</label><input type="number" name="net_weight" data-step="${stepId}" data-dyn="weight" step="0.01"></div>
        <div class="form-row"><label>單箱毛重(kg)</label><input type="number" name="gross_weight" data-step="${stepId}" data-dyn="weight" step="0.01"></div>
        <div class="form-row"><label>箱數</label><input type="number" name="box_count" data-step="${stepId}" data-dyn="weight"></div>
      </div>
    </div>
    <button class="btn btn-sm btn-outline" onclick="addWeightRow('${stepId}')">+ 新增一行</button>`,

    '2-4': `<div class="form-row"><label>發票號</label><input type="text" name="invoice_no" data-step="${stepId}" placeholder="INV-2026-001"></div>
    <div class="form-row"><label>自動彙整數據</label><textarea name="consolidated" data-step="${stepId}" rows="4" placeholder="（將自動從稱重步驟帶入）" id="consolidated-${stepId}"></textarea></div>`,

    '2-5': `<div class="form-grid">
      <div class="form-row"><label>中文品名</label><input type="text" name="name_cn" data-step="${stepId}" placeholder="筆記型電腦"></div>
      <div class="form-row"><label>英文品名</label><input type="text" name="name_en" data-step="${stepId}" placeholder="Laptop Computer"></div>
    </div>
    <div class="form-grid">
      <div class="form-row"><label>數量</label><input type="number" name="quantity" data-step="${stepId}" placeholder="0"></div>
      <div class="form-row"><label>箱號</label><input type="text" name="box_no" data-step="${stepId}" placeholder="CTN-001"></div>
    </div>`,

    '2-6': `<div class="form-grid-3">
      <div class="form-row"><label>總體積 CBM</label><input type="number" name="total_cbm" data-step="${stepId}" step="0.01" placeholder="0.00" onchange="calcContainer()"></div>
      <div class="form-row"><label>總毛重 kg</label><input type="number" name="total_gw" data-step="${stepId}" step="0.1" placeholder="0.0" onchange="calcContainer()"></div>
      <div class="form-row"><label>托盤數</label><input type="number" name="pallet_num" data-step="${stepId}" placeholder="0"></div>
    </div>
    <div class="form-row"><label>推薦櫃型</label><div id="containerResult-${stepId}" style="padding:8px;background:var(--card2);border-radius:4px;font-size:13px;color:var(--accent)">請輸入數據後自動計算</div></div>`,

    '2-7': `<div class="form-row"><label>托盤號</label><input type="text" name="pallet_id" data-step="${stepId}" placeholder="PLT-01"></div>
    <div class="form-row"><label>托盤內容</label><textarea name="pallet_content" data-step="${stepId}" rows="3" placeholder="品名 / 箱號 / 數量"></textarea></div>`,

    '2-8': `<div class="form-row"><label>發票號</label><input type="text" name="invoice_no" data-step="${stepId}" placeholder="INV-2026-001"></div>
    <div class="form-row"><label>自動生成</label><div style="padding:8px;background:var(--card2);border-radius:4px;font-size:12px;color:var(--text2)">點擊「執行工具」自動從前面數據生成 PACKING LIST</div></div>`,

    // Module 3: 報關資料準備
    '3-1': `<div class="form-grid">
      <div class="form-row"><label>品名（中文）</label><input type="text" name="product_cn" data-step="${stepId}" placeholder="例：筆記型電腦"></div>
      <div class="form-row"><label>品名（英文）</label><input type="text" name="product_en" data-step="${stepId}" placeholder="e.g. Laptop Computer"></div>
    </div>
    <div class="form-grid-3">
      <div class="form-row"><label>型號</label><input type="text" name="model" data-step="${stepId}" placeholder="MacBook Pro 16"></div>
      <div class="form-row"><label>品牌</label><input type="text" name="brand" data-step="${stepId}" placeholder="Apple"></div>
      <div class="form-row"><label>材質</label><input type="text" name="material" data-step="${stepId}" placeholder="鋁合金/塑膠"></div>
    </div>
    <div class="form-row"><label>用途</label><textarea name="usage" data-step="${stepId}" rows="2" placeholder="商務辦公、程式開發等"></textarea></div>
    <div class="form-row"><label>工作原理</label><textarea name="principle" data-step="${stepId}" rows="2" placeholder="電子計算機原理..."></textarea></div>
    <div class="form-row"><label>主要組成</label><textarea name="composition" data-step="${stepId}" rows="2" placeholder="主機板、CPU、記憶體、硬碟、螢幕..."></textarea></div>`,

    '3-2': `<div class="form-row">
      <label>檢查清單</label>
      <label class="checkbox-row"><input type="checkbox" name="chk_catalog" data-step="${stepId}"> 產品型錄已收集</label>
      <label class="checkbox-row"><input type="checkbox" name="chk_spec" data-step="${stepId}"> 規格書已收集</label>
      <label class="checkbox-row"><input type="checkbox" name="chk_photo" data-step="${stepId}"> 產品照片已收集</label>
      <label class="checkbox-row"><input type="checkbox" name="chk_label" data-step="${stepId}"> 銘牌照片已收集</label>
    </div>`,

    '3-3': `<div class="form-grid">
      <div class="form-row"><label>確認人</label><input type="text" name="engineer" data-step="${stepId}" placeholder="王工程師"></div>
      <div class="form-row"><label>確認日期</label><input type="date" name="confirm_date" data-step="${stepId}"></div>
    </div>
    <div class="form-row"><label>功能描述</label><textarea name="func_desc" data-step="${stepId}" rows="2" placeholder="產品功能描述..."></textarea></div>
    <div class="form-row"><label>用途</label><textarea name="usage" data-step="${stepId}" rows="2" placeholder="產品用途..."></textarea></div>
    <div class="form-row"><label>原理</label><textarea name="principle" data-step="${stepId}" rows="2" placeholder="工作原理..."></textarea></div>`,

    '3-4': `<div class="form-row"><label>申報要素（自動從Step 1生成）</label>
      <textarea name="declaration" data-step="${stepId}" rows="5" placeholder="自動帶入：品名、品牌、型號、用途、材質、原理..." id="declaration-${stepId}"></textarea></div>
    <button class="btn btn-sm btn-outline" onclick="autoFillDeclaration('${stepId}')">自動填入</button>`,

    '3-5': `<div class="form-row"><label>代理名稱</label><input type="text" name="agent_name" data-step="${stepId}" placeholder="報關行名稱"></div>
    <div class="form-row"><label>核驗結果</label><select name="verify_result" data-step="${stepId}"><option value="confirmed">確認</option><option value="adjust">需調整</option></select></div>
    <div class="form-row"><label>備註</label><textarea name="notes" data-step="${stepId}" rows="2" placeholder="代理核驗備註..."></textarea></div>`,

    '3-6': `<div class="form-grid">
      <div class="form-row"><label>HS編碼（自動帶入）</label><input type="text" name="hs_code" data-step="${stepId}" placeholder="8471.30" id="hscode-${stepId}"></div>
      <div class="form-row"><label>進口國</label><select name="import_country" data-step="${stepId}"><option value="thailand">泰國</option><option value="indonesia">印尼</option></select></div>
    </div>
    <div class="form-row"><label>FTA查詢結果</label><div id="ftaResult-${stepId}" style="padding:8px;background:var(--card2);border-radius:4px;font-size:12px;color:var(--text2)">點擊執行後顯示</div></div>`,

    // Module 4: 單據製作與審核
    '4-1': `<div class="form-grid-3">
      <div class="form-row"><label>合同號</label><input type="text" name="contract_no" data-step="${stepId}" placeholder="CT-2026-001"></div>
      <div class="form-row"><label>PO號</label><input type="text" name="po_no" data-step="${stepId}" placeholder="PO-2026-0055"></div>
      <div class="form-row"><label>發票號</label><input type="text" name="invoice_no" data-step="${stepId}" placeholder="INV-2026-001"></div>
    </div>
    <div class="form-row"><label>自動帶入數據</label><div style="padding:8px;background:var(--card2);border-radius:4px;font-size:12px;color:var(--text2)">將自動從前面模組帶入品項、數量、金額等數據</div></div>`,

    '4-2': `<div class="form-row"><label>選擇要校驗的檔案</label>
      <select name="file_select" data-step="${stepId}"><option value="all">全套單據</option><option value="invoice">發票</option><option value="packing">裝箱單</option><option value="co">產地證</option></select></div>`,

    '4-3': `<div class="form-grid">
      <div class="form-row"><label>提單號</label><input type="text" name="bl_no" data-step="${stepId}" placeholder="BL-2026-001"></div>
      <div class="form-row"><label>船名</label><input type="text" name="vessel" data-step="${stepId}" placeholder="EVER GOLDEN"></div>
    </div>
    <div class="form-grid-3">
      <div class="form-row"><label>航次</label><input type="text" name="voyage" data-step="${stepId}" placeholder="V.001E"></div>
      <div class="form-row"><label>開船日</label><input type="date" name="etd" data-step="${stepId}"></div>
      <div class="form-row"><label>到港日</label><input type="date" name="eta" data-step="${stepId}"></div>
    </div>`,

    '4-4': `<div class="form-row"><label>清關代理</label><input type="text" name="clear_agent" data-step="${stepId}" placeholder="泰國清關代理名稱"></div>
    <div class="form-row"><label>清關文件清單</label>
      <label class="checkbox-row"><input type="checkbox" name="doc_invoice" data-step="${stepId}"> 商業發票 Commercial Invoice</label>
      <label class="checkbox-row"><input type="checkbox" name="doc_packing" data-step="${stepId}"> 裝箱單 Packing List</label>
      <label class="checkbox-row"><input type="checkbox" name="doc_bl" data-step="${stepId}"> 提單 Bill of Lading</label>
      <label class="checkbox-row"><input type="checkbox" name="doc_co" data-step="${stepId}"> 產地證 Certificate of Origin</label>
      <label class="checkbox-row"><input type="checkbox" name="doc_insurance" data-step="${stepId}"> 保險單 Insurance Policy</label>
    </div>`,

    // Module 5: 訂艙與物流
    '5-1': `<div class="form-grid-3">
      <div class="form-row"><label>總毛重 (kg)</label><input type="number" name="total_gw" data-step="${stepId}" step="0.1" placeholder="0.0"></div>
      <div class="form-row"><label>總體積 (CBM)</label><input type="number" name="total_cbm" data-step="${stepId}" step="0.01" placeholder="0.00"></div>
      <div class="form-row"><label>總箱數</label><input type="number" name="total_boxes" data-step="${stepId}" placeholder="0"></div>
    </div>`,

    '5-2': `<div class="form-row"><label>櫃型（自動從包裝模組帶入）</label>
      <div id="autoContainer-${stepId}" style="padding:8px;background:var(--card2);border-radius:4px;font-size:12px;color:var(--text2)">自動帶入櫃型核算結果</div></div>`,

    '5-3': `<div class="form-grid">
      <div class="form-row"><label>起運港</label><select name="origin_port" data-step="${stepId}"><option value="CNSHA">上海</option><option value="CNNGB">寧波</option><option value="CNSHE">蛇口</option><option value="CNTXG">天津</option></select></div>
      <div class="form-row"><label>目的港</label><select name="dest_port" data-step="${stepId}"><option value="THBKK">曼谷(BKK)</option><option value="THLCH">林查班(Laem Chabang)</option><option value="IDJKT">雅加達(Jakarta)</option><option value="IDSUB">泗水(Surabaya)</option></select></div>
    </div>
    <div class="form-grid">
      <div class="form-row"><label>櫃型</label><select name="container_type" data-step="${stepId}"><option value="20GP">20'GP</option><option value="40GP">40'GP</option><option value="40HQ">40'HQ</option></select></div>
      <div class="form-row"><label>預計出貨日</label><input type="date" name="ship_date" data-step="${stepId}"></div>
    </div>`,

    '5-4': `<div id="quoteRows-${stepId}">
      <div class="dyn-row">
        <div class="form-row"><label>貨代名</label><input type="text" name="forwarder" data-step="${stepId}" data-dyn="quote"></div>
        <div class="form-row"><label>海運費(USD)</label><input type="number" name="ocean_freight" data-step="${stepId}" data-dyn="quote" step="0.01"></div>
        <div class="form-row"><label>THC(USD)</label><input type="number" name="thc" data-step="${stepId}" data-dyn="quote" step="0.01"></div>
        <div class="form-row"><label>文件費(USD)</label><input type="number" name="doc_fee" data-step="${stepId}" data-dyn="quote" step="0.01"></div>
        <div class="form-row"><label>天數</label><input type="number" name="transit_days" data-step="${stepId}" data-dyn="quote"></div>
        <button class="btn-remove" onclick="removeRow(this)">&#10005;</button>
      </div>
    </div>
    <button class="btn btn-sm btn-outline" onclick="addQuoteRow('${stepId}')">+ 新增報價</button>
    <div class="form-row" style="margin-top:8px"><label>自動推薦</label><div id="quoteRecommend-${stepId}" style="padding:8px;background:var(--card2);border-radius:4px;font-size:12px;color:var(--accent)">計算中...</div></div>`,

    '5-5': `<div class="form-grid">
      <div class="form-row"><label>選定貨代</label><input type="text" name="selected_forwarder" data-step="${stepId}" placeholder="貨代名稱"></div>
      <div class="form-row"><label>確認船期</label><input type="date" name="confirmed_date" data-step="${stepId}"></div>
    </div>
    <div class="form-row"><label>提單號</label><input type="text" name="bl_no" data-step="${stepId}" placeholder="BL-2026-001"></div>`,

    // Module 6: 稅費核算
    '6-1': `<div class="form-grid">
      <div class="form-row"><label>HS編碼</label><input type="text" name="hs_code" data-step="${stepId}" placeholder="8471.30" id="tax-hscode-${stepId}"></div>
      <div class="form-row"><label>進口國</label><select name="country" data-step="${stepId}"><option value="thailand">泰國</option><option value="indonesia">印尼</option></select></div>
    </div>`,

    '6-2': `<div class="form-grid">
      <div class="form-row"><label>CIF價值</label><input type="number" name="cif_value" data-step="${stepId}" step="0.01" placeholder="0.00" onchange="calcTax('${stepId}')"></div>
      <div class="form-row"><label>FOB價值</label><input type="number" name="fob_value" data-step="${stepId}" step="0.01" placeholder="0.00"></div>
    </div>
    <div class="form-grid">
      <div class="form-row"><label>運費</label><input type="number" name="freight" data-step="${stepId}" step="0.01" placeholder="0.00" onchange="calcTax('${stepId}')"></div>
      <div class="form-row"><label>保險費</label><input type="number" name="insurance" data-step="${stepId}" step="0.01" placeholder="0.00" onchange="calcTax('${stepId}')"></div>
    </div>
    <div class="form-row"><label>幣別</label><select name="currency" data-step="${stepId}"><option value="USD">USD</option><option value="THB">THB</option><option value="IDR">IDR</option><option value="CNY">CNY</option></select></div>
    <div class="form-row"><label>稅費估算結果</label><div id="taxResult-${stepId}" style="padding:8px;background:var(--card2);border-radius:4px;font-size:12px;color:var(--accent)">請輸入數據後自動計算</div></div>`,

    '6-3': `<div class="form-row"><label>FTA比對</label><div style="padding:8px;background:var(--card2);border-radius:4px;font-size:12px;color:var(--text2)">點擊執行工具自動比對FTA優惠稅率</div></div>`,
    '6-4': `<div class="form-row"><label>報表生成</label><div style="padding:8px;background:var(--card2);border-radius:4px;font-size:12px;color:var(--text2)">點擊執行工具自動生成稅費報表</div></div>`,

    // Module 7: ERP錄入
    '7-1': `<div class="form-row"><label>選擇模板檔案</label><input type="text" name="template_file" data-step="${stepId}" placeholder="檔案路徑或名稱"></div>`,
    '7-2': `<div class="form-row"><label>選擇已驗證的檔案</label><input type="text" name="verified_file" data-step="${stepId}" placeholder="已驗證的採購訂單檔案"></div>`,
    '7-3': `<div class="form-row"><label>選擇已驗證的檔案</label><input type="text" name="verified_file" data-step="${stepId}" placeholder="已驗證的銷售訂單檔案"></div>`,
    '7-4': `<div class="form-row"><label>選擇已驗證的檔案</label><input type="text" name="verified_file" data-step="${stepId}" placeholder="已驗證的料號主數據檔案"></div>`,
    '7-5': `<div class="form-row"><label>選擇已驗證的檔案</label><input type="text" name="verified_file" data-step="${stepId}" placeholder="已驗證的產品主數據檔案"></div>`,
    '7-6': `<div class="form-row"><label>批量錄入設定</label><div style="padding:8px;background:var(--card2);border-radius:4px;font-size:12px;color:var(--text2)">將自動錄入前面步驟驗證過的檔案</div></div>`,

    // Module 8: 印尼協同
    '8-1': `<div class="form-row"><label>訂單號</label><input type="text" name="order_no" data-step="${stepId}" placeholder="ORD-2026-001"></div>`,
    '8-2': `<div class="form-row"><label>出貨通知（自動帶入數據生成雙語訊息）</label>
      <textarea name="shipment_notice" data-step="${stepId}" rows="5" placeholder="自動從前面模組帶入數據..."></textarea></div>`,
    '8-3': `<div class="form-row"><label>文件清單</label>
      <label class="checkbox-row"><input type="checkbox" name="doc1" data-step="${stepId}"> 商業發票</label>
      <label class="checkbox-row"><input type="checkbox" name="doc2" data-step="${stepId}"> 裝箱單</label>
      <label class="checkbox-row"><input type="checkbox" name="doc3" data-step="${stepId}"> 提單</label>
      <label class="checkbox-row"><input type="checkbox" name="doc4" data-step="${stepId}"> 產地證</label>
    </div>
    <div class="form-row"><label>接收人</label><input type="text" name="receiver" data-step="${stepId}" placeholder="印尼方聯絡人"></div>`,
    '8-4': `<div class="form-grid">
      <div class="form-row"><label>起始日期</label><input type="date" name="date_from" data-step="${stepId}"></div>
      <div class="form-row"><label>結束日期</label><input type="date" name="date_to" data-step="${stepId}"></div>
    </div>`,
    '8-5': `<div class="form-row"><label>原文</label><textarea name="source_text" data-step="${stepId}" rows="4" placeholder="輸入需要翻譯的文字..."></textarea></div>
    <div class="form-row"><label>翻譯方向</label><select name="direction" data-step="${stepId}"><option value="cn-id">中文 → 印尼文</option><option value="id-cn">印尼文 → 中文</option></select></div>`,

    // Module 9: 資料歸檔
    '9-1': `<div class="form-grid">
      <div class="form-row"><label>合同號</label><input type="text" name="contract_no" data-step="${stepId}" placeholder="CT-2026-001"></div>
      <div class="form-row"><label>文件類型</label><select name="doc_type" data-step="${stepId}"><option value="invoice">發票</option><option value="packing">裝箱單</option><option value="bl">提單</option><option value="co">產地證</option><option value="all">全套</option></select></div>
    </div>
    <div class="form-row"><label>自動命名規則預覽</label><div style="padding:8px;background:var(--card2);border-radius:4px;font-size:12px;color:var(--text2)">格式: [合同號]_[文件類型]_[日期]</div></div>`,
    '9-2': `<div class="form-row"><label>微盤更新</label><div style="padding:8px;background:var(--card2);border-radius:4px;font-size:12px;color:var(--text2)">點擊執行自動生成微盤更新清單</div></div>`,
    '9-3': `<div class="form-grid">
      <div class="form-row"><label>起始日期</label><input type="date" name="date_from" data-step="${stepId}"></div>
      <div class="form-row"><label>結束日期</label><input type="date" name="date_to" data-step="${stepId}"></div>
    </div>`,

    // Module 10: 品牌授權
    '10-1': `<div class="form-row"><label>品牌授權預警</label><div style="padding:8px;background:var(--card2);border-radius:4px;font-size:12px;color:var(--text2)">點擊執行工具自動掃描即將到期的品牌授權</div></div>`,
    '10-2': `<div class="form-row"><label>品牌授權總覽</label><div style="padding:8px;background:var(--card2);border-radius:4px;font-size:12px;color:var(--text2)">點擊執行工具顯示所有品牌授權狀態</div></div>`,
    '10-3': `<div class="form-grid">
      <div class="form-row"><label>品牌名</label><input type="text" name="brand_name" data-step="${stepId}" placeholder="品牌名稱"></div>
      <div class="form-row"><label>到期日</label><input type="date" name="expiry_date" data-step="${stepId}"></div>
    </div>
    <div class="form-row"><label>負責人</label><input type="text" name="responsible" data-step="${stepId}" placeholder="負責人姓名"></div>`,

    // Module 11: 政策監控
    '11-1': `<div class="form-row"><label>泰國政策追蹤</label><div style="padding:8px;background:var(--card2);border-radius:4px;font-size:12px;color:var(--text2)">點擊執行工具查詢泰國最新關務政策</div></div>`,
    '11-2': `<div class="form-row"><label>印尼政策追蹤</label><div style="padding:8px;background:var(--card2);border-radius:4px;font-size:12px;color:var(--text2)">點擊執行工具查詢印尼最新關務政策</div></div>`,
    '11-3': `<div class="form-row"><label>綜合政策動態</label><div style="padding:8px;background:var(--card2);border-radius:4px;font-size:12px;color:var(--text2)">點擊執行工具顯示所有政策動態</div></div>`,

    // Module 12: 每日運營
    '12-1': `<div class="form-row"><label>進度總覽</label><div style="padding:8px;background:var(--card2);border-radius:4px;font-size:12px;color:var(--text2)">點擊執行顯示所有工作票進度</div></div>`,
    '12-2': `<div class="form-row"><label>異常追蹤</label><div style="padding:8px;background:var(--card2);border-radius:4px;font-size:12px;color:var(--text2)">點擊執行顯示異常事項</div></div>`,
    '12-3': `<div class="form-row"><label>KPI看板</label><div style="padding:8px;background:var(--card2);border-radius:4px;font-size:12px;color:var(--text2)">點擊執行顯示KPI數據</div></div>`,
    '12-4': `<div class="form-row"><label>工作日誌</label><textarea name="daily_log" data-step="${stepId}" rows="5" placeholder="記錄今日工作事項..."></textarea></div>`,
  };

  const key = `${modId}-${stepNum}`;
  const formHTML = forms[key];
  if (!formHTML) return '';
  return `<div class="panel"><h4>&#128221; 資料輸入</h4>${formHTML}
    <div style="margin-top:10px"><button class="btn btn-save" onclick="saveStepData('${stepId}')">&#128190; 儲存資料</button></div></div>`;
}

// ═══════════════════════ Upload HTML ═══════════════════════
function getUploadHTML(stepId) {
  return `<div class="panel"><h4>&#128206; 附件上傳</h4>
    <div class="upload-area" onclick="document.getElementById('upload-${stepId}').click()" id="uploadArea-${stepId}">
      <input type="file" multiple id="upload-${stepId}" onchange="updateUploadLabel('${stepId}')">
      <div id="uploadLabel-${stepId}">點擊選擇檔案或拖曳至此處</div>
    </div>
    <button class="btn btn-sm btn-outline" style="margin-top:6px" onclick="uploadFile('${stepId}')">上傳</button>
    <div class="file-list" id="fileList-${stepId}"></div>
  </div>`;
}

function updateUploadLabel(stepId) {
  const input = document.getElementById('upload-' + stepId);
  const label = document.getElementById('uploadLabel-' + stepId);
  if (input.files.length > 0) {
    label.textContent = `已選擇 ${input.files.length} 個檔案`;
  }
}

// ═══════════════════════ Execute HTML ═══════════════════════
function getExecuteHTML(stepId, tool, cmdStr) {
  return `<div class="panel execute-panel"><h4>&#9889; 執行工具</h4>
    <div class="cmd-display"><code>${cmdStr}</code></div>
    <button class="btn btn-execute" onclick="executeStep('${stepId}','${tool.tool}','${tool.args}')">&#9654; 執行工具</button>
  </div>`;
}

// ═══════════════════════ Data Operations ═══════════════════════
function saveStepData(stepId) {
  if (!currentTicket) { toast('請先選擇工作票', true); return; }
  const card = document.getElementById('card-' + stepId);
  if (!card) return;
  const inputs = card.querySelectorAll('[data-step]');
  const data = {};
  inputs.forEach(inp => {
    const name = inp.name || inp.id;
    if (inp.type === 'checkbox') {
      if (!data[name]) data[name] = [];
      if (inp.checked) data[name].push(true); else data[name].push(false);
    } else {
      data[name] = inp.value;
    }
  });
  // Also save notes
  const notes = document.getElementById('notes-' + stepId);
  if (notes) data._notes = notes.value;

  fetch(`/api/step_data/${currentTicket.id}/${stepId}`, {
    method: 'POST',
    headers: {'Content-Type':'application/json'},
    body: JSON.stringify(data)
  }).then(r=>r.json()).then(() => {
    toast('資料已儲存');
    addNotification('step', '資料已儲存', stepId + ' 表單資料儲存成功', 'success');
  });
}

function loadStepDataToForm(stepId) {
  if (!currentTicket) return;
  fetch(`/api/step_data/${currentTicket.id}/${stepId}`).then(r=>r.json()).then(data => {
    if (!data || Object.keys(data).length === 0) return;
    ticketData[stepId] = data;
    const card = document.getElementById('card-' + stepId);
    if (!card) return;
    Object.keys(data).forEach(key => {
      if (key === '_notes') {
        const notes = document.getElementById('notes-' + stepId);
        if (notes) notes.value = data[key];
        return;
      }
      const inputs = card.querySelectorAll(`[name="${key}"]`);
      if (inputs.length > 0) {
        if (Array.isArray(data[key])) {
          inputs.forEach((inp, i) => {
            if (inp.type === 'checkbox' && data[key][i] !== undefined) inp.checked = data[key][i];
          });
        } else {
          inputs[0].value = data[key];
        }
      }
    });
  });
}

function saveNotes(stepId) {
  saveStepData(stepId);
}

function markComplete(stepId, checked) {
  if (!currentTicket) return;
  if (checked) {
    fetch(`/api/step/${currentTicket.id}/${stepId}/complete`, {method:'POST', headers:{'Content-Type':'application/json'}})
      .then(r=>r.json()).then(() => {
        if (!currentTicket.completed_steps) currentTicket.completed_steps = [];
        currentTicket.completed_steps.push(stepId);
        const card = document.getElementById('card-' + stepId);
        if (card) card.classList.add('completed');
        const status = document.getElementById('status-' + stepId);
        if (status) { status.className = 'step-status done'; status.innerHTML = '&#9989; 已完成'; }
        const ctime = document.getElementById('ctime-' + stepId);
        if (ctime) ctime.textContent = new Date().toLocaleString('zh-Hant');
        renderSidebar();
        updateProgressBar();
        toast('步驟已完成');
        addNotification('step', '步驟完成: ' + stepId, '工作票 ' + currentTicket.id + ' 的步驟已標記完成', 'step');
      });
  }
}

function updateProgressBar() {
  const pct = getTotalProgress();
  const fills = document.querySelectorAll('.progress-bar .fill');
  const pcts = document.querySelectorAll('.tb-pct');
  fills.forEach(f => f.style.width = pct + '%');
  pcts.forEach(p => p.textContent = pct + '%');
}

// ═══════════════════════ File Upload ═══════════════════════
function uploadFile(stepId) {
  if (!currentTicket) { toast('請先選擇工作票', true); return; }
  const input = document.getElementById('upload-' + stepId);
  if (!input.files.length) { toast('請先選擇檔案', true); return; }
  const formData = new FormData();
  for (let file of input.files) formData.append('file', file);
  fetch(`/api/upload/${currentTicket.id}/${stepId}`, {
    method: 'POST',
    body: formData
  }).then(r=>r.json()).then(data => {
    if (data.success) {
      const list = document.getElementById('fileList-' + stepId);
      data.files.forEach(f => {
        list.innerHTML += `<div class="file-item"><span>&#128196;</span><a href="${f.path}" target="_blank">${f.name}</a></div>`;
      });
      toast('檔案上傳成功');
      input.value = '';
      document.getElementById('uploadLabel-' + stepId).textContent = '點擊選擇檔案或拖曳至此處';
    } else {
      toast('上傳失敗: ' + (data.error || ''), true);
    }
  });
}

// ═══════════════════════ Tool Execution ═══════════════════════
function executeStep(stepId, tool, args) {
  if (!currentTicket) { toast('請先選擇工作票', true); return; }
  const outPanel = document.getElementById('output-' + stepId);
  const term = document.getElementById('term-' + stepId);
  outPanel.style.display = '';
  term.textContent = '⏳ 執行中...';

  // Collect form data as extra args
  const card = document.getElementById('card-' + stepId);
  const mainInput = card.querySelector('[data-step]');
  let extraArg = '';
  if (mainInput && mainInput.value) {
    extraArg = ' "' + mainInput.value.replace(/"/g, '\\"') + '"';
  }

  const fullArgs = args + extraArg;
  const url = `/api/execute?tool=${encodeURIComponent(tool)}&args=${encodeURIComponent(fullArgs)}&ticket=${encodeURIComponent(currentTicket.id)}`;

  fetch(url).then(r=>r.json()).then(data => {
    if (data.success) {
      term.textContent = data.stdout || '(無輸出)';
      term.style.color = '#7ee787';
      addNotification('tool', '工具執行完成: ' + tool, data.stdout ? data.stdout.substring(0, 200) : '', 'tool');
    } else {
      term.textContent = '錯誤: ' + (data.error || data.stderr || '未知錯誤');
      term.style.color = '#f87171';
      addNotification('tool', '工具執行失敗: ' + tool, data.error || data.stderr, 'error');
    }
  }).catch(err => {
    term.textContent = '請求失敗: ' + err.message;
    term.style.color = '#f87171';
    addNotification('tool', '工具執行失敗: ' + tool, err.message, 'error');
  });
}

// ═══════════════════════ Dynamic Rows ═══════════════════════
function addWeightRow(stepId) {
  const container = document.getElementById('weightRows-' + stepId);
  const row = document.createElement('div');
  row.className = 'dyn-row';
  row.innerHTML = `
    <div class="form-row"><label>型號</label><input type="text" name="model" data-step="${stepId}" data-dyn="weight"></div>
    <div class="form-row"><label>單箱淨重(kg)</label><input type="number" name="net_weight" data-step="${stepId}" data-dyn="weight" step="0.01"></div>
    <div class="form-row"><label>單箱毛重(kg)</label><input type="number" name="gross_weight" data-step="${stepId}" data-dyn="weight" step="0.01"></div>
    <div class="form-row"><label>箱數</label><input type="number" name="box_count" data-step="${stepId}" data-dyn="weight"></div>
    <button class="btn-remove" onclick="removeRow(this)">&#10005;</button>`;
  container.appendChild(row);
}

function addQuoteRow(stepId) {
  const container = document.getElementById('quoteRows-' + stepId);
  const row = document.createElement('div');
  row.className = 'dyn-row';
  row.innerHTML = `
    <div class="form-row"><label>貨代名</label><input type="text" name="forwarder" data-step="${stepId}" data-dyn="quote"></div>
    <div class="form-row"><label>海運費(USD)</label><input type="number" name="ocean_freight" data-step="${stepId}" data-dyn="quote" step="0.01"></div>
    <div class="form-row"><label>THC(USD)</label><input type="number" name="thc" data-step="${stepId}" data-dyn="quote" step="0.01"></div>
    <div class="form-row"><label>文件費(USD)</label><input type="number" name="doc_fee" data-step="${stepId}" data-dyn="quote" step="0.01"></div>
    <div class="form-row"><label>天數</label><input type="number" name="transit_days" data-step="${stepId}" data-dyn="quote"></div>
    <button class="btn-remove" onclick="removeRow(this)">&#10005;</button>`;
  container.appendChild(row);
}

function removeRow(btn) {
  btn.parentElement.remove();
}

// ═══════════════════════ Auto-Calculations ═══════════════════════
function calcContainer() {
  const mod = MODULES.find(x => x.id === 2);
  const step6Id = 'm2s6';
  const card = document.getElementById('card-' + step6Id);
  if (!card) return;
  const cbm = parseFloat(card.querySelector('[name="total_cbm"]')?.value) || 0;
  const gw = parseFloat(card.querySelector('[name="total_gw"]')?.value) || 0;
  const result = document.getElementById('containerResult-' + step6Id);
  if (!result) return;

  let rec = '';
  if (cbm <= 0 && gw <= 0) { rec = '請輸入數據'; }
  else if (cbm <= 25 && gw <= 18000) { rec = '推薦: 20\'GP (內容積 ~33 CBM, 載重 ~18 噸)'; }
  else if (cbm <= 55 && gw <= 26000) { rec = '推薦: 40\'GP (內容積 ~67 CBM, 載重 ~26 噸)'; }
  else if (cbm <= 68 && gw <= 26000) { rec = '推薦: 40\'HQ (內容積 ~76 CBM, 載重 ~26 噸)'; }
  else { rec = `需要多個貨櫃。CBM: ${cbm}, 毛重: ${gw}kg`; }

  if (cbm > 0 && gw > 0) {
    rec += `\n體積: ${cbm} CBM | 毛重: ${gw} kg`;
  }
  result.textContent = rec;
}

function calcTax(stepId) {
  const card = document.getElementById('card-' + stepId);
  if (!card) return;
  const cif = parseFloat(card.querySelector('[name="cif_value"]')?.value) || 0;
  const fob = parseFloat(card.querySelector('[name="fob_value"]')?.value) || 0;
  const freight = parseFloat(card.querySelector('[name="freight"]')?.value) || 0;
  const insurance = parseFloat(card.querySelector('[name="insurance"]')?.value) || 0;
  const result = document.getElementById('taxResult-' + stepId);
  if (!result) return;

  let cifCalc = cif || (fob + freight + insurance);
  let duty = cifCalc * 0.05;  // Default 5% duty
  let vat = (cifCalc + duty) * 0.07;  // 7% VAT
  let total = duty + vat;

  result.innerHTML = `CIF價值: ${cifCalc.toFixed(2)}\n進口關稅(5%): ${duty.toFixed(2)}\n增值稅(7%): ${vat.toFixed(2)}\n預計總稅費: <b>${total.toFixed(2)}</b>`;
}

function autoFillDeclaration(stepId) {
  // Pull data from m3s1 (HS classification step)
  const hsData = ticketData['m3s1'] || {};
  const decl = document.getElementById('declaration-' + stepId);
  if (!decl) return;
  const lines = [];
  if (hsData.product_cn) lines.push(`品名：${hsData.product_cn}`);
  if (hsData.product_en) lines.push(`英文品名：${hsData.product_en}`);
  if (hsData.brand) lines.push(`品牌：${hsData.brand}`);
  if (hsData.model) lines.push(`型號：${hsData.model}`);
  if (hsData.material) lines.push(`材質：${hsData.material}`);
  if (hsData.usage) lines.push(`用途：${hsData.usage}`);
  if (hsData.principle) lines.push(`工作原理：${hsData.principle}`);
  if (hsData.composition) lines.push(`主要組成：${hsData.composition}`);
  decl.value = lines.length > 0 ? lines.join('\n') : '（無數據可帶入，請先完成Step 1 HS歸類）';
}

// ═══════════════════════ Toast ═══════════════════════
function toast(msg, isError) {
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.className = 'toast show' + (isError ? ' error' : '');
  setTimeout(() => t.className = 'toast', 3000);
  if (isError) {
    addNotification('alert', msg, '', 'error');
  } else {
    addNotification('step', msg, '', 'success');
  }
}

// ═══════════════════════ Export & Reports ═══════════════════════
function exportReport() {
  if (!currentTicket) { toast('請先選擇工作票', true); return; }
  window.print();
}

function exportProgressReport() {
  if (!currentTicket) { toast('請先選擇工作票', true); return; }
  const t = currentTicket;
  const allSteps = [];
  MODULES.forEach(m => m.steps.forEach((s, i) => allSteps.push({
    module: m.name, step: s, id: `m${m.id}s${i+1}`,
    done: (t.completed_steps || []).includes(`m${m.id}s${i+1}`)
  })));
  const done = allSteps.filter(s => s.done).length;
  let report = `關務工作平台 — 工作票進度報告\n`;
  report += `${'═'.repeat(50)}\n`;
  report += `工作票: ${t.id} | ${t.name}\n`;
  report += `PO號: ${t.po_number || '-'} | 類型: ${t.type==='export'?'出口':'進口'}\n`;
  report += `目的國: ${t.country==='thailand'?'泰國':'印尼'}\n`;
  report += `建立日期: ${t.created ? t.created.substring(0,10) : '-'}\n`;
  report += `總進度: ${Math.round(done/allSteps.length*100)}% (${done}/${allSteps.length})\n`;
  report += `${'─'.repeat(50)}\n\n`;
  let currentMod = '';
  allSteps.forEach(s => {
    if (s.module !== currentMod) {
      currentMod = s.module;
      report += `\n【${s.module}】\n`;
    }
    report += `  ${s.done ? '✅' : '⬜'} ${s.step}\n`;
  });
  report += `\n${'═'.repeat(50)}\n`;
  report += `報告生成時間: ${new Date().toLocaleString('zh-Hant')}\n`;
  const blob = new Blob([report], {type:'text/plain;charset=utf-8'});
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `${t.id}_進度報告_${new Date().toISOString().substring(0,10)}.txt`;
  a.click();
  toast('進度報告已匯出');
}

// ═══════════════════════ Ticket Management ═══════════════════════
function deleteTicket(ticketId) {
  if (!confirm(`確定要刪除工作票 ${ticketId} 嗎？此操作不可逆。`)) return;
  fetch('/api/ticket/delete', {
    method: 'POST',
    headers: {'Content-Type':'application/json'},
    body: JSON.stringify({id: ticketId})
  }).then(r=>r.json()).then(() => {
    if (currentTicket && currentTicket.id === ticketId) {
      currentTicket = null;
      document.getElementById('emptyState').style.display = '';
      document.getElementById('ticketContent').style.display = 'none';
    }
    loadAllTickets();
    renderSidebar();
    toast(`工作票 ${ticketId} 已刪除`);
  });
}

function toggleTicketStatus(ticketId) {
  if (!currentTicket || currentTicket.id !== ticketId) return;
  const newStatus = currentTicket.status === 'active' ? 'completed' : 'active';
  fetch(`/api/ticket/${ticketId}/update`, {
    method: 'POST',
    headers: {'Content-Type':'application/json'},
    body: JSON.stringify({status: newStatus})
  }).then(r=>r.json()).then(() => {
    currentTicket.status = newStatus;
    toast(`工作票狀態已更新為: ${newStatus === 'active' ? '進行中' : '已完成'}`);
    loadAllTickets();
    renderModuleContent(currentModule);
  });
}

// ═══════════════════════ Drag & Drop Upload ═══════════════════════
document.addEventListener('DOMContentLoaded', () => {
  document.addEventListener('dragover', e => { e.preventDefault(); });
  document.addEventListener('drop', e => {
    e.preventDefault();
    const area = e.target.closest('.upload-area');
    if (!area) return;
    const stepId = area.id.replace('uploadArea-', '');
    const input = document.getElementById('upload-' + stepId);
    if (e.dataTransfer.files.length > 0 && input) {
      input.files = e.dataTransfer.files;
      updateUploadLabel(stepId);
    }
  });
});

// ═══════════════════════ Auto-Save on Blur ═══════════════════════
document.addEventListener('focusout', e => {
  const el = e.target;
  if (el.hasAttribute('data-step') && currentTicket) {
    const stepId = el.getAttribute('data-step');
    clearTimeout(el._saveTimer);
    el._saveTimer = setTimeout(() => saveStepDataSilent(stepId), 1000);
  }
});

function saveStepDataSilent(stepId) {
  if (!currentTicket) return;
  const card = document.getElementById('card-' + stepId);
  if (!card) return;
  const inputs = card.querySelectorAll('[data-step]');
  const data = {};
  inputs.forEach(inp => {
    const name = inp.name || inp.id;
    if (inp.type === 'checkbox') {
      if (!data[name]) data[name] = [];
      if (inp.checked) data[name].push(true); else data[name].push(false);
    } else {
      data[name] = inp.value;
    }
  });
  const notes = document.getElementById('notes-' + stepId);
  if (notes) data._notes = notes.value;
  fetch(`/api/step_data/${currentTicket.id}/${stepId}`, {
    method: 'POST',
    headers: {'Content-Type':'application/json'},
    body: JSON.stringify(data)
  });
}

// ═══════════════════════ Cross-Step Auto-Fill ═══════════════════════
function autoFillStep(sourceId, targetId, fieldMap) {
  const srcData = ticketData[sourceId] || {};
  if (Object.keys(srcData).length === 0) {
    toast('來源步驟無數據，請先完成並儲存來源步驟', true);
    return;
  }
  const card = document.getElementById('card-' + targetId);
  if (!card) return;
  let filled = 0;
  fieldMap.forEach(([srcField, tgtField]) => {
    if (srcData[srcField]) {
      const target = card.querySelector(`[name="${tgtField}"]`);
      if (target) {
        target.value = srcData[srcField];
        filled++;
      }
    }
  });
  if (filled > 0) {
    toast(`已從 ${sourceId} 自動填入 ${filled} 個欄位`);
    saveStepData(targetId);
  } else {
    toast('無可填入的數據', true);
  }
}

// Auto-fill from packaging (m2s6) to shipping (m5s1, m5s2)
function autoFillShipping() {
  autoFillStep('m2s6', 'm5s1', [
    ['total_cbm', 'total_cbm'],
    ['total_gw', 'total_gw'],
  ]);
}

// Auto-fill from HS (m3s1) to tax (m6s1)
function autoFillTaxHS() {
  autoFillStep('m3s1', 'm6s1', [
    ['product_cn', 'hs_code'],  // HS code might be stored in product fields
  ]);
}

// Auto-fill from order (m4s3) to clearance (m4s4)
function autoFillClearance() {
  autoFillStep('m4s3', 'm4s4', [
    ['bl_no', 'clear_agent'],
  ]);
}

// ═══════════════════════ Step Navigation ═══════════════════════
function scrollToStep(stepId) {
  const card = document.getElementById('card-' + stepId);
  if (card) {
    card.scrollIntoView({behavior:'smooth', block:'start'});
    if (!card.classList.contains('open')) card.classList.add('open');
  }
}

function nextIncompleteStep() {
  if (!currentTicket) return;
  const m = MODULES.find(x => x.id === currentModule);
  if (!m) return;
  for (let i = 0; i < m.steps.length; i++) {
    const stepId = `m${currentModule}s${i+1}`;
    if (!(currentTicket.completed_steps || []).includes(stepId)) {
      scrollToStep(stepId);
      return;
    }
  }
  toast('此模組所有步驟已完成');
}

// ═══════════════════════ Bulk Operations ═══════════════════════
function saveAllStepsInModule() {
  if (!currentTicket) { toast('請先選擇工作票', true); return; }
  const m = MODULES.find(x => x.id === currentModule);
  if (!m) return;
  let saved = 0;
  m.steps.forEach((_, i) => {
    const stepId = `m${currentModule}s${i+1}`;
    const card = document.getElementById('card-' + stepId);
    if (card && card.querySelector('[data-step]')) {
      saveStepDataSilent(stepId);
      saved++;
    }
  });
  toast(`已儲存 ${saved} 個步驟的資料`);
}

function expandAllSteps() {
  document.querySelectorAll('.step-card').forEach(c => c.classList.add('open'));
}

function collapseAllSteps() {
  document.querySelectorAll('.step-card').forEach(c => c.classList.remove('open'));
}

// ═══════════════════════ Keyboard Shortcuts ═══════════════════════
document.addEventListener('keydown', e => {
  if (e.ctrlKey || e.metaKey) {
    if (e.key === 's') {
      e.preventDefault();
      saveAllStepsInModule();
    }
  }
});

// ═══════════════════════ Load Uploaded Files List ═══════════════════════
function loadUploadedFiles(stepId) {
  if (!currentTicket) return;
  // Check if there are already files shown; the upload handler appends them
  // For initial load, we rely on the upload response to populate
}

// ═══════════════════════ Smart Container Recommendation ═══════════════════════
function calcContainerDetailed() {
  const step6Id = 'm2s6';
  const card = document.getElementById('card-' + step6Id);
  if (!card) return;
  const cbm = parseFloat(card.querySelector('[name="total_cbm"]')?.value) || 0;
  const gw = parseFloat(card.querySelector('[name="total_gw"]')?.value) || 0;
  const pallets = parseInt(card.querySelector('[name="pallet_num"]')?.value) || 0;
  const result = document.getElementById('containerResult-' + step6Id);
  if (!result) return;

  if (cbm <= 0 && gw <= 0) {
    result.textContent = '請輸入數據後自動計算';
    return;
  }

  // Container specs
  const containers = [
    {type:'20\'GP', maxCBM:33.2, maxGW:18000, maxPallets:10, cost:1},
    {type:'40\'GP', maxCBM:67.7, maxGW:26000, maxPallets:20, cost:1.6},
    {type:'40\'HQ', maxCBM:76.3, maxGW:26000, maxPallets:22, cost:1.7},
  ];

  let html = '';
  containers.forEach(c => {
    const needByCBM = Math.ceil(cbm / c.maxCBM);
    const needByGW = Math.ceil(gw / c.maxGW);
    const needByPallets = pallets > 0 ? Math.ceil(pallets / c.maxPallets) : 0;
    const needed = Math.max(needByCBM, needByGW, needByPallets, 1);
    const usedCBM = (cbm / (needed * c.maxCBM) * 100).toFixed(0);
    html += `${c.type}: ${needed} 個 (使用率 ${usedCBM}%)\n`;
  });

  const rec = containers.find(c => {
    const n = Math.max(Math.ceil(cbm/c.maxCBM), Math.ceil(gw/c.maxGW), pallets>0?Math.ceil(pallets/c.maxPallets):0, 1);
    return n === 1;
  });
  if (rec) html += `\n🏆 推薦: ${rec.type}`;

  result.innerHTML = html.replace(/\n/g, '<br>');
}

// Override calcContainer with detailed version
function calcContainer() { calcContainerDetailed(); }

// ═══════════════════════ Enhanced Tax Calculator ═══════════════════════
function calcTaxEnhanced(stepId) {
  const card = document.getElementById('card-' + stepId);
  if (!card) return;
  const cif = parseFloat(card.querySelector('[name="cif_value"]')?.value) || 0;
  const fob = parseFloat(card.querySelector('[name="fob_value"]')?.value) || 0;
  const freight = parseFloat(card.querySelector('[name="freight"]')?.value) || 0;
  const insurance = parseFloat(card.querySelector('[name="insurance"]')?.value) || 0;
  const currency = card.querySelector('[name="currency"]')?.value || 'USD';
  const result = document.getElementById('taxResult-' + stepId);
  if (!result) return;

  // Determine country from ticket
  const country = currentTicket?.country || 'thailand';
  let dutyRate, vatRate;
  if (country === 'thailand') {
    dutyRate = 0.05; vatRate = 0.07;
  } else {
    dutyRate = 0.05; vatRate = 0.11; // Indonesia VAT is 11%
  }

  const cifCalc = cif || (fob + freight + insurance);
  const duty = cifCalc * dutyRate;
  const vat = (cifCalc + duty) * vatRate;
  const total = duty + vat;

  let html = `<table style="width:100%;font-size:12px">`;
  html += `<tr><td>CIF價值</td><td style="text-align:right">${cifCalc.toFixed(2)} ${currency}</td></tr>`;
  html += `<tr><td>進口關稅 (${(dutyRate*100).toFixed(0)}%)</td><td style="text-align:right">${duty.toFixed(2)} ${currency}</td></tr>`;
  html += `<tr><td>增值稅 (${(vatRate*100).toFixed(0)}%)</td><td style="text-align:right">${vatRate.toFixed(2)} ${currency}</td></tr>`;
  html += `<tr style="border-top:1px solid var(--border);font-weight:bold"><td>預計總稅費</td><td style="text-align:right;color:var(--accent)">${total.toFixed(2)} ${currency}</td></tr>`;
  html += `</table>`;
  result.innerHTML = html;
}

function calcTax(stepId) { calcTaxEnhanced(stepId); }

// ═══ Notification Center ═══
let notifications = [];
let notifFilter = 'all';

function addNotification(type, title, detail, tag) {
  const n = {
    id: Date.now() + Math.random(),
    type: type,
    title: title,
    detail: detail || '',
    tag: tag || type,
    time: new Date(),
    read: false,
    expanded: false
  };
  notifications.unshift(n);
  if (notifications.length > 100) notifications = notifications.slice(0, 100);
  updateNotifBadge();
  if (document.getElementById('notifPanel').classList.contains('show')) {
    renderNotifList();
  }
}

function updateNotifBadge() {
  const unread = notifications.filter(n => !n.read).length;
  const badge = document.getElementById('notifBadge');
  if (unread > 0) {
    badge.textContent = unread > 99 ? '99+' : unread;
    badge.style.display = 'flex';
  } else {
    badge.style.display = 'none';
  }
  document.getElementById('notifCount').textContent = notifications.length + ' 則通知';
}

function toggleNotificationCenter() {
  const panel = document.getElementById('notifPanel');
  const overlay = document.getElementById('notifOverlay');
  const isOpen = panel.classList.contains('show');
  if (isOpen) {
    closeNotificationCenter();
  } else {
    panel.classList.add('show');
    overlay.classList.add('show');
    renderNotifList();
  }
}

function closeNotificationCenter() {
  document.getElementById('notifPanel').classList.remove('show');
  document.getElementById('notifOverlay').classList.remove('show');
}

function switchNotifTab(filter, btn) {
  notifFilter = filter;
  document.querySelectorAll('.notif-tab').forEach(t => t.classList.remove('active'));
  if (btn) btn.classList.add('active');
  renderNotifList();
}

function renderNotifList() {
  const list = document.getElementById('notifList');
  let filtered = notifFilter === 'all' ? notifications : notifications.filter(n => n.type === notifFilter);

  if (filtered.length === 0) {
    list.innerHTML = '<div class="notif-empty">尚無通知</div>';
    return;
  }

  list.innerHTML = filtered.map(n => {
    const timeStr = formatNotifTime(n.time);
    const tagClass = n.tag || n.type;
    const tagLabel = {tool:'工具',step:'步驟',alert:'提醒',success:'成功',error:'失敗'}[tagClass] || tagClass;
    const icon = {tool:'⚙️',step:'✅',alert:'⚠️'}[n.type] || '📌';
    return `<div class="notif-item ${n.read?'':'unread'} ${n.expanded?'expanded':''}" onclick="toggleNotifItem('${n.id}')">
      <div class="notif-item-head">
        <span class="notif-icon">${icon}</span>
        <span class="notif-title">${escHtml(n.title)}</span>
        <span class="notif-tag ${tagClass}">${tagLabel}</span>
        <span class="notif-time">${timeStr}</span>
      </div>
      <div class="notif-expand-hint">點擊展開詳情</div>
      ${n.detail ? `<div class="notif-detail">${escHtml(n.detail)}</div>` : ''}
    </div>`;
  }).join('');
}

function toggleNotifItem(id) {
  const n = notifications.find(x => String(x.id) === String(id));
  if (!n) return;
  n.read = true;
  n.expanded = !n.expanded;
  updateNotifBadge();
  renderNotifList();
}

function markAllNotifRead() {
  notifications.forEach(n => n.read = true);
  updateNotifBadge();
  renderNotifList();
}

function clearAllNotif() {
  notifications = [];
  updateNotifBadge();
  renderNotifList();
}

function formatNotifTime(date) {
  if (!(date instanceof Date)) date = new Date(date);
  const now = new Date();
  const diff = Math.floor((now - date) / 1000);
  if (diff < 60) return '剛剛';
  if (diff < 3600) return Math.floor(diff/60) + '分鐘前';
  if (diff < 86400) return Math.floor(diff/3600) + '小時前';
  return date.toLocaleDateString('zh-TW', {month:'short',day:'numeric'});
}

function escHtml(str) {
  if (!str) return '';
  return str.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/\n/g,'<br>');
}
</script>
</body>
</html>"""


# ═══════════════════════ Server Startup ═══════════════════════

def main():
    server = ThreadedHTTPServer(('0.0.0.0', PORT), PlatformHandler)
    print(f"""
╔══════════════════════════════════════════════════════════╗
║                                                          ║
║   關務工作平台 v3.0  已啟動                              ║
║                                                          ║
║   請開啟瀏覽器: http://localhost:{PORT}                  ║
║                                                          ║
║   數據目錄: {DATA_DIR}                                   ║
║   按 Ctrl+C 停止伺服器                                   ║
║                                                          ║
╚══════════════════════════════════════════════════════════╝
""")
    # Open browser after a short delay (skip in cloud)
    if not IS_CLOUD:
        threading.Timer(1.0, lambda: webbrowser.open(f'http://localhost:{PORT}')).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print('\n伺服器已停止。')
        server.shutdown()

if __name__ == '__main__':
    main()
