"""
Lojinha Interna - Genomma Lab  |  Backend Python/Flask
"""
import os, json, threading, smtplib, logging
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from pathlib import Path

from flask import Flask, request, jsonify, send_from_directory, Response
import openpyxl
import functools

app   = Flask(__name__, static_folder='public')
BASE  = Path(__file__).parent

# No Render, usa /tmp para arquivos mutáveis (stock, orders, uploads)
# Localmente, usa a própria pasta do projeto
IS_RENDER = os.getenv('RENDER', '') != ''
TMP       = Path('/tmp/lojinha') if IS_RENDER else BASE

DATA  = BASE / 'data'          # Excel fica sempre junto ao código
UPLOAD= TMP  / 'uploads'
STOCK = TMP  / 'stock.json'
ORDERS= TMP  / 'orders.json'
EXCEL = DATA / 'Estoque_Lojinha_jun26.xlsx'
LOCK  = threading.Lock()

UPLOAD.mkdir(parents=True, exist_ok=True)
(TMP).mkdir(parents=True, exist_ok=True)

logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')
log = logging.getLogger(__name__)

# ── Config ────────────────────────────────────────────────────────────────────
SMTP_HOST  = 'smtp.gmail.com'
SMTP_PORT  = 587
SMTP_USER  = ''
SMTP_PASS  = ''
DEST_EMAIL = 'maycon.silva@contractor.genommalab.com'
PORT       = 3000
ADMIN_USER = 'admin'
ADMIN_PASS = '5827'

env_file = BASE / '.env'
if env_file.exists():
    for line in env_file.read_text().splitlines():
        line = line.strip()
        if line and not line.startswith('#') and '=' in line:
            k, v = line.split('=', 1)
            os.environ.setdefault(k.strip(), v.strip())
SMTP_HOST  = os.getenv('SMTP_HOST', SMTP_HOST)
SMTP_PORT  = int(os.getenv('SMTP_PORT', str(SMTP_PORT)))
SMTP_USER  = os.getenv('SMTP_USER', SMTP_USER)
SMTP_PASS  = os.getenv('SMTP_PASS', SMTP_PASS)
PORT       = int(os.getenv('PORT', str(PORT)))
ADMIN_PASS = os.getenv('ADMIN_PASS', ADMIN_PASS)

# ── Autenticação Admin ────────────────────────────────────────────────────────
def _check_auth(username, password):
    return username == ADMIN_USER and password == ADMIN_PASS

def require_admin(f):
    @functools.wraps(f)
    def decorated(*args, **kwargs):
        auth = request.authorization
        if not auth or not _check_auth(auth.username, auth.password):
            return Response(
                'Acesso restrito. Digite o usuário e senha.',
                401,
                {'WWW-Authenticate': 'Basic realm="Admin Lojinha Genomma"'}
            )
        return f(*args, **kwargs)
    return decorated

# ── Estoque ───────────────────────────────────────────────────────────────────
def load_from_excel():
    log.info('📊 Lendo planilha Excel...')
    wb = openpyxl.load_workbook(str(EXCEL), read_only=True, data_only=True)
    ws = wb['Estoque']
    products = {}
    for row in ws.iter_rows(min_row=3, values_only=True):
        if not row[0]: continue
        code = str(row[0]).strip()
        name = str(row[1]).strip() if row[1] else ''
        try: qtd = float(row[13]) if row[13] is not None else 0
        except: qtd = 0
        if qtd > 0:
            if code not in products:
                products[code] = {'code': code, 'name': name, 'stock': 0}
            products[code]['stock'] += qtd
    wb.close()
    for p in products.values(): p['stock'] = int(p['stock'])
    log.info(f'✅ {len(products)} produtos carregados.')
    return products

stock_data: dict = {}

def init_stock():
    global stock_data
    if STOCK.exists():
        try:
            stock_data = json.loads(STOCK.read_text('utf-8'))
            log.info(f'📦 Estoque carregado ({len(stock_data)} produtos).')
            return
        except: pass
    stock_data = load_from_excel()
    STOCK.write_text(json.dumps(stock_data, ensure_ascii=False, indent=2), 'utf-8')

def save_stock():
    STOCK.write_text(json.dumps(stock_data, ensure_ascii=False, indent=2), 'utf-8')

# ── Pedidos ───────────────────────────────────────────────────────────────────
def load_orders():
    if ORDERS.exists():
        try: return json.loads(ORDERS.read_text('utf-8'))
        except: pass
    return []

def write_orders(orders):
    ORDERS.write_text(json.dumps(orders, ensure_ascii=False, indent=2), 'utf-8')

# ── API: produtos ─────────────────────────────────────────────────────────────
@app.get('/api/products')
def api_products():
    with LOCK:
        lst = [p for p in stock_data.values() if p['stock'] > 0]
    lst.sort(key=lambda x: x['name'].lower())
    return jsonify(lst)

# ── API: pedido ───────────────────────────────────────────────────────────────
@app.post('/api/order')
def api_order():
    nome  = request.form.get('nome','').strip()
    email = request.form.get('email','').strip()
    pcode = request.form.get('produto_code','').strip()
    qstr  = request.form.get('quantidade','0')
    tipo  = request.form.get('tipo_comprador','').strip()
    file  = request.files.get('comprovante')

    if not all([nome, email, pcode, qstr, tipo]):
        return jsonify({'error':'Preencha todos os campos obrigatórios.'}), 400
    try:
        qty = int(float(qstr))
        if qty < 1: raise ValueError
    except:
        return jsonify({'error':'Quantidade inválida.'}), 400

    with LOCK:
        product = stock_data.get(pcode)
        if not product: return jsonify({'error':'Produto não encontrado.'}), 404
        if qty > product['stock']:
            return jsonify({'error':f'Estoque insuficiente. Disponível: {product["stock"]} un.'}), 400
        stock_data[pcode]['stock'] -= qty
        save_stock()
        snap = dict(product)

    attach_path, attach_name = None, None
    if file and file.filename:
        ts   = datetime.now().strftime('%Y%m%d_%H%M%S')
        safe = ''.join(c if c.isalnum() or c in '._-' else '_' for c in file.filename)
        attach_name = f'{ts}_{safe}'
        attach_path = UPLOAD / attach_name
        file.save(str(attach_path))

    order = {
        'id':           datetime.now().strftime('%Y%m%d%H%M%S%f')[:17],
        'data_hora':    datetime.now().strftime('%d/%m/%Y %H:%M:%S'),
        'nome':         nome,
        'email':        email,
        'tipo':         tipo,
        'produto_code': pcode,
        'produto_name': snap['name'],
        'quantidade':   qty,
        'comprovante':  attach_name or '',
        'status':       'pendente'
    }
    with LOCK:
        orders = load_orders()
        orders.insert(0, order)
        write_orders(orders)
    log.info(f'📝 Pedido {order["id"]} — {nome} / {snap["name"]} x{qty}')

    def try_email():
        try: _send_email(nome, email, tipo, snap, qty, attach_path, attach_name)
        except Exception as e: log.warning(f'⚠️ Email não enviado: {e}')
    threading.Thread(target=try_email, daemon=True).start()

    return jsonify({'success':True,'message':'Pedido finalizado!','produto':snap['name'],'quantidade':qty})

# ── API: listar pedidos ────────────────────────────────────────────────────────
@app.get('/api/orders')
@require_admin
def api_orders():
    return jsonify(load_orders())

# ── API: atualizar status ──────────────────────────────────────────────────────
@app.post('/api/orders/<order_id>/status')
@require_admin
def api_status(order_id):
    data       = request.get_json(force=True)
    new_status = data.get('status','pendente')
    with LOCK:
        orders = load_orders()
        found  = False
        for o in orders:
            if o.get('id') == order_id:
                o['status'] = new_status
                if new_status == 'entregue':
                    o['entregue_em'] = datetime.now().strftime('%d/%m/%Y %H:%M')
                else:
                    o.pop('entregue_em', None)
                found = True
                break
        if not found:
            return jsonify({'error':'Pedido não encontrado'}), 404
        write_orders(orders)
    return jsonify({'ok':True,'status':new_status})

# ── API: excluir pedido ───────────────────────────────────────────────────────
@app.delete('/api/orders/<order_id>')
@require_admin
def api_delete_order(order_id):
    with LOCK:
        orders = load_orders()
        found = None
        for o in orders:
            if o.get('id') == order_id:
                found = o
                break
        if not found:
            return jsonify({'error': 'Pedido não encontrado'}), 404
        orders.remove(found)
        write_orders(orders)
        pcode = found.get('produto_code')
        qty   = int(found.get('quantidade', 0))
        if pcode and pcode in stock_data and qty > 0:
            stock_data[pcode]['stock'] += qty
            save_stock()
            log.info(f'🔄 Estoque restaurado: +{qty} de {pcode}')
    return jsonify({'ok': True})

# ── Página Admin ──────────────────────────────────────────────────────────────
@app.get('/admin')
@require_admin
def admin():
    total = len(load_orders())
    html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Admin — Lojinha Genomma</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:'Segoe UI',system-ui,sans-serif;background:#F3EDF9;min-height:100vh}}
header{{background:linear-gradient(135deg,#4A1B7A,#1A5FB4);padding:22px 32px;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:12px}}
header h1{{color:white;font-size:1.4rem;font-weight:800}}
header p{{color:rgba(255,255,255,.75);font-size:.88rem}}
.badge{{background:rgba(255,255,255,.2);color:white;padding:5px 14px;border-radius:20px;font-size:.83rem;font-weight:700}}
.container{{max-width:1250px;margin:0 auto;padding:24px 18px}}
/* stats */
.stats{{display:flex;gap:13px;flex-wrap:wrap;margin-bottom:20px}}
.stat{{background:white;border-radius:13px;padding:15px 19px;flex:1;min-width:120px;box-shadow:0 2px 10px rgba(74,27,122,.09);border-left:4px solid #4A1B7A}}
.stat.green{{border-color:#27AE60}} .stat.orange{{border-color:#E67E22}}
.stat .n{{font-size:1.7rem;font-weight:800;color:#4A1B7A}}
.stat.green .n{{color:#27AE60}} .stat.orange .n{{color:#E67E22}}
.stat .l{{font-size:.75rem;color:#999;margin-top:1px}}
/* filtros */
.filters{{background:white;border-radius:13px;padding:16px 20px;margin-bottom:18px;box-shadow:0 2px 10px rgba(74,27,122,.09);display:flex;gap:12px;flex-wrap:wrap;align-items:flex-end}}
.fgroup label{{font-size:.72rem;font-weight:700;color:#7B3FAD;text-transform:uppercase;letter-spacing:.4px;display:block;margin-bottom:4px}}
.fgroup input,.fgroup select{{border:2px solid rgba(74,27,122,.15);border-radius:8px;padding:7px 11px;font-size:.85rem;color:#333;outline:none;background:white;min-width:130px}}
.fgroup input:focus,.fgroup select:focus{{border-color:#7B3FAD}}
.btn{{border:none;cursor:pointer;border-radius:8px;padding:8px 18px;font-size:.85rem;font-weight:700;transition:.18s}}
.btn-primary{{background:linear-gradient(135deg,#4A1B7A,#1A5FB4);color:white}}
.btn-primary:hover{{opacity:.88}}
.btn-ghost{{background:white;color:#7B3FAD;border:2px solid rgba(74,27,122,.2)}}
.btn-ghost:hover{{background:#F3EDF9}}
/* card */
.card{{background:white;border-radius:15px;box-shadow:0 2px 14px rgba(74,27,122,.09);overflow:hidden}}
.card-head{{padding:15px 20px;border-bottom:1px solid #F0E8FB;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px}}
.card-head h2{{font-size:.92rem;font-weight:700;color:#4A1B7A}}
.count-lbl{{font-size:.78rem;color:#aaa}}
/* tabela */
table{{width:100%;border-collapse:collapse}}
thead th{{background:#F8F3FF;padding:9px 12px;text-align:left;font-size:.7rem;font-weight:700;color:#7B3FAD;text-transform:uppercase;letter-spacing:.5px;white-space:nowrap}}
tbody tr{{border-bottom:1px solid #FAF5FF;transition:background .13s}}
tbody tr:hover{{background:#FAF5FF}}
tbody tr.entregue{{background:#F0FFF4}}
tbody tr.entregue:hover{{background:#E8FFF0}}
td{{padding:10px 12px;font-size:.86rem;vertical-align:middle}}
/* badges */
.tg{{display:inline-block;padding:3px 10px;border-radius:20px;font-size:.72rem;font-weight:700;white-space:nowrap}}
.tg-g{{background:#E8F5E9;color:#2E7D32}}
.tg-t{{background:#E3F2FD;color:#1565C0}}
.tg-p{{background:#FFF8E1;color:#E65100}}
.tg-e{{background:#E8F5E9;color:#1B5E20}}
/* botão entrega */
.btn-del{{border:none;cursor:pointer;border-radius:7px;padding:5px 12px;font-size:.76rem;font-weight:700;transition:.18s;white-space:nowrap}}
.btn-del.pend{{background:#4A1B7A;color:white}} .btn-del.pend:hover{{background:#6B2FA0}}
.btn-del.done{{background:#E8F5E9;color:#2E7D32;border:1.5px solid #A5D6A7}}
.btn-del.done:hover{{background:#FFEBEE;color:#c62828;border-color:#EF9A9A}}
.btn-exc{{border:none;cursor:pointer;border-radius:7px;padding:5px 10px;font-size:.76rem;font-weight:700;transition:.18s;background:#FFF0F0;color:#c62828;border:1.5px solid #FFCDD2;white-space:nowrap}}
.btn-exc:hover{{background:#FFEBEE;border-color:#EF9A9A}}
/* empty */
.empty{{text-align:center;padding:50px;color:#ccc}}
.empty span{{font-size:2.5rem;display:block;margin-bottom:10px}}
.rbtn{{background:#4A1B7A;color:white;border:none;padding:7px 15px;border-radius:8px;cursor:pointer;font-size:.82rem;font-weight:600}}
.rbtn:hover{{background:#6B2FA0}}
@media(max-width:750px){{
  thead{{display:none}} .stat .n{{font-size:1.3rem}}
  tbody tr{{display:block;padding:12px;border-bottom:2px solid #F0E8FB}}
  td{{display:block;padding:2px 0}}
}}
</style>
</head>
<body>
<header>
  <div><h1>🛍️ Painel de Pedidos</h1><p>Lojinha Interna — Genomma Lab</p></div>
  <div style="display:flex;gap:10px;align-items:center;flex-wrap:wrap">
    <span class="badge" id="hdr-count">📦 {total} pedido(s)</span>
    <button class="rbtn" onclick="loadOrders()">↻ Atualizar</button>
    <a href="/" class="rbtn" style="text-decoration:none">🛍️ Lojinha</a>
  </div>
</header>

<div class="container">
  <div class="stats">
    <div class="stat">       <div class="n" id="st-total">—</div>   <div class="l">Total de pedidos</div></div>
    <div class="stat">       <div class="n" id="st-genomma">—</div> <div class="l">Genomma</div></div>
    <div class="stat">       <div class="n" id="st-terc">—</div>    <div class="l">Terceirizados</div></div>
    <div class="stat">       <div class="n" id="st-units">—</div>   <div class="l">Unidades pedidas</div></div>
    <div class="stat green"> <div class="n" id="st-done">—</div>    <div class="l">✅ Entregues</div></div>
    <div class="stat orange"><div class="n" id="st-pend">—</div>    <div class="l">⏳ Pendentes</div></div>
  </div>

  <div class="filters">
    <div class="fgroup"><label>De</label><input type="date" id="f-de"></div>
    <div class="fgroup"><label>Até</label><input type="date" id="f-ate"></div>
    <div class="fgroup">
      <label>Status</label>
      <select id="f-status">
        <option value="">Todos</option>
        <option value="pendente">⏳ Pendente</option>
        <option value="entregue">✅ Entregue</option>
      </select>
    </div>
    <div class="fgroup">
      <label>Tipo</label>
      <select id="f-tipo">
        <option value="">Todos</option>
        <option value="genomma">🏢 Genomma</option>
        <option value="terceirizado">🤝 Terceirizado</option>
      </select>
    </div>
    <div class="fgroup"><label>&nbsp;</label><button class="btn btn-primary" onclick="applyFilter()">🔍 Filtrar</button></div>
    <div class="fgroup"><label>&nbsp;</label><button class="btn btn-ghost"   onclick="clearFilter()">✕ Limpar</button></div>
  </div>

  <div class="card">
    <div class="card-head">
      <h2>📋 Pedidos (mais recente primeiro)</h2>
      <span class="count-lbl" id="tbl-count"></span>
    </div>
    <div id="tbl-wrap" style="overflow-x:auto"></div>
  </div>
</div>

<script>
let allOrders = [];

async function loadOrders() {{
  try {{
    const r = await fetch('/api/orders');
    allOrders = await r.json();
    document.getElementById('hdr-count').textContent = '📦 ' + allOrders.length + ' pedido(s)';
    applyFilter();
  }} catch(e) {{ console.error(e); }}
}}

function stats(orders) {{
  document.getElementById('st-total').textContent   = orders.length;
  document.getElementById('st-genomma').textContent = orders.filter(o=>o.tipo==='genomma').length;
  document.getElementById('st-terc').textContent    = orders.filter(o=>o.tipo==='terceirizado').length;
  document.getElementById('st-units').textContent   = orders.reduce((s,o)=>s+(parseInt(o.quantidade)||0),0);
  document.getElementById('st-done').textContent    = orders.filter(o=>o.status==='entregue').length;
  document.getElementById('st-pend').textContent    = orders.filter(o=>o.status!=='entregue').length;
}}

function applyFilter() {{
  const de  = document.getElementById('f-de').value;
  const ate = document.getElementById('f-ate').value;
  const st  = document.getElementById('f-status').value;
  const tp  = document.getElementById('f-tipo').value;
  const fil = allOrders.filter(o => {{
    const p = o.data_hora ? o.data_hora.split(' ')[0].split('/') : null;
    const d = p ? p[2]+'-'+p[1]+'-'+p[0] : '';
    if (de  && d < de)  return false;
    if (ate && d > ate) return false;
    if (st  && o.status !== st) return false;
    if (tp  && o.tipo   !== tp) return false;
    return true;
  }});
  stats(fil);
  render(fil);
}}

function clearFilter() {{
  ['f-de','f-ate'].forEach(id=>document.getElementById(id).value='');
  ['f-status','f-tipo'].forEach(id=>document.getElementById(id).value='');
  stats(allOrders); render(allOrders);
}}

function render(orders) {{
  const wrap = document.getElementById('tbl-wrap');
  document.getElementById('tbl-count').textContent = orders.length + ' registro(s)';
  if (!orders.length) {{
    wrap.innerHTML = "<div class='empty'><span>📭</span>Nenhum pedido encontrado.</div>";
    return;
  }}
  let rows = '';
  orders.forEach(o => {{
    const done   = o.status === 'entregue';
    const tBadge = o.tipo==='genomma' ? "<span class='tg tg-g'>🏢 Genomma</span>" : "<span class='tg tg-t'>🤝 Terceirizado</span>";
    const sBadge = done
      ? "<span class='tg tg-e'>✅ Entregue" + (o.entregue_em ? '<br><small style=\\"font-weight:400;opacity:.75;font-size:.68rem\\">' + o.entregue_em + '</small>' : '') + "</span>"
      : "<span class='tg tg-p'>⏳ Pendente</span>";
    const cLink  = o.comprovante ? `<a href="/uploads/${{o.comprovante}}" target="_blank" style="color:#4A1B7A;font-weight:600;text-decoration:none;">📎 Ver</a>` : '—';
    const newSt  = done ? 'pendente' : 'entregue';
    const btnLbl = done ? '↩ Desfazer' : '✅ Marcar entregue';
    const btnCls = done ? 'done' : 'pend';
    rows += `<tr class="${{done?'entregue':''}}" id="row-${{o.id}}">
      <td style="white-space:nowrap;color:#666;font-size:.78rem;">${{o.data_hora||''}}</td>
      <td style="font-weight:600">${{o.nome||''}}</td>
      <td style="color:#4A1B7A;font-size:.82rem;">${{o.email||''}}</td>
      <td>${{tBadge}}</td>
      <td style="max-width:200px">${{o.produto_name||''}}</td>
      <td style="text-align:center;font-weight:700;color:#27AE60;font-size:1rem;">${{o.quantidade||''}}</td>
      <td style="text-align:center">${{cLink}}</td>
      <td style="text-align:center">${{sBadge}}</td>
      <td style="text-align:center"><button class="btn-del ${{btnCls}}" onclick="toggle('${{o.id}}','${{newSt}}')">${{btnLbl}}</button></td>
      <td style="text-align:center"><button class="btn-exc" onclick="excluir('${{o.id}}','${{o.produto_name||''}}','${{o.quantidade||0}}')">🗑️ Excluir</button></td>
    </tr>`;
  }});
  wrap.innerHTML = `<table><thead><tr>
    <th>Data/Hora</th><th>Nome</th><th>E-mail</th><th>Tipo</th><th>Produto</th>
    <th style="text-align:center">Qtd</th><th style="text-align:center">Comprovante</th>
    <th style="text-align:center">Status</th><th style="text-align:center">Entrega</th>
    <th style="text-align:center">Excluir</th>
  </tr></thead><tbody>${{rows}}</tbody></table>`;
}}

async function toggle(id, newStatus) {{
  const btn = document.querySelector(`#row-${{id}} .btn-del`);
  if (btn) {{ btn.disabled=true; btn.textContent='...'; }}
  try {{
    const r = await fetch(`/api/orders/${{id}}/status`,{{
      method:'POST', headers:{{'Content-Type':'application/json'}},
      body: JSON.stringify({{status:newStatus}})
    }});
    if (r.ok) {{ await loadOrders(); }}
    else {{ alert('Erro ao atualizar. Tente novamente.'); if(btn) btn.disabled=false; }}
  }} catch(e) {{ alert('Erro de conexão.'); if(btn) btn.disabled=false; }}
}}

async function excluir(id, nome, qty) {{
  if (!confirm(`Excluir este pedido?\n\n"${{nome}}" (x${{qty}} un.)\n\nO estoque será restaurado automaticamente.`)) return;
  try {{
    const r = await fetch(`/api/orders/${{id}}`, {{method: 'DELETE'}});
    if (r.ok) {{ await loadOrders(); }}
    else {{ alert('Erro ao excluir. Tente novamente.'); }}
  }} catch(e) {{ alert('Erro de conexão.'); }}
}}

loadOrders();
setInterval(loadOrders, 30000);
</script>
</body></html>"""
    return Response(html, mimetype='text/html')

# ── Uploads ───────────────────────────────────────────────────────────────────
@app.get('/uploads/<filename>')
def serve_upload(filename):
    return send_from_directory(str(UPLOAD), filename)

# ── Frontend ──────────────────────────────────────────────────────────────────
@app.get('/')
def index():
    return send_from_directory(app.static_folder, 'index.html')

# ── Email ─────────────────────────────────────────────────────────────────────
def _send_email(nome, email, tipo, product, qty, attach_path, attach_name):
    if not SMTP_USER or not SMTP_PASS: return
    tipo_label = '🏢 Genomma' if tipo == 'genomma' else '🤝 Terceirizado(a)'
    dh = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
    msg = MIMEMultipart('alternative')
    msg['Subject'] = f'🛍️ Novo Pedido — {nome}'
    msg['From']    = f'"Lojinha Genomma" <{SMTP_USER}>'
    msg['To']      = DEST_EMAIL
    msg['Reply-To']= email
    html = f"""<!DOCTYPE html><html><head><meta charset="UTF-8"></head>
<body style="font-family:Arial,sans-serif;background:#f5f5f5;padding:20px;">
<div style="max-width:600px;margin:auto;background:white;border-radius:12px;overflow:hidden;box-shadow:0 4px 20px rgba(0,0,0,.1);">
  <div style="background:linear-gradient(135deg,#4A1B7A,#1A5FB4);padding:28px;text-align:center;">
    <h1 style="color:white;margin:0;font-size:22px;">🛍️ Novo Pedido Recebido</h1>
    <p style="color:rgba(255,255,255,.8);margin:6px 0 0;">{dh}</p>
  </div>
  <div style="padding:28px;">
    <h2 style="color:#4A1B7A;border-bottom:2px solid #f0e6f6;padding-bottom:10px;">👤 Comprador</h2>
    <table style="width:100%;border-collapse:collapse;">
      <tr><td style="padding:7px 0;color:#666;width:140px;"><b>Nome:</b></td><td>{nome}</td></tr>
      <tr><td style="padding:7px 0;color:#666;"><b>E-mail:</b></td><td>{email}</td></tr>
      <tr><td style="padding:7px 0;color:#666;"><b>Vínculo:</b></td><td>{tipo_label}</td></tr>
    </table>
    <h2 style="color:#4A1B7A;border-bottom:2px solid #f0e6f6;padding-bottom:10px;margin-top:22px;">📦 Produto</h2>
    <table style="width:100%;border-collapse:collapse;">
      <tr><td style="padding:7px 0;color:#666;width:140px;"><b>Produto:</b></td><td>{product["name"]}</td></tr>
      <tr><td style="padding:7px 0;color:#666;"><b>Código:</b></td><td>{product["code"]}</td></tr>
      <tr><td style="padding:7px 0;color:#666;"><b>Quantidade:</b></td>
          <td style="font-size:18px;font-weight:bold;color:#27AE60;">{qty} un.</td></tr>
    </table>
    <div style="background:#f8f3ff;border-left:4px solid #4A1B7A;padding:13px;border-radius:4px;margin-top:22px;">
      <p style="margin:0;color:#4A1B7A;"><b>💡 Ação:</b> Separar {qty} un. de <em>{product["name"]}</em> → enviar para {email}</p>
    </div>
  </div>
  <div style="background:#f5f5f5;padding:14px;text-align:center;">
    <p style="color:#aaa;font-size:11px;margin:0;">Lojinha Interna Genomma Lab · {dh}</p>
  </div>
</div></body></html>"""
    msg.attach(MIMEText(html, 'html', 'utf-8'))
    if attach_path and Path(attach_path).exists():
        with open(str(attach_path),'rb') as f:
            part = MIMEBase('application','octet-stream')
            part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition','attachment',filename=attach_name or 'comprovante')
        full = MIMEMultipart('mixed')
        for k in ('Subject','From','To','Reply-To'): full[k] = msg[k]
        full.attach(msg); full.attach(part)
        send_msg = full
    else:
        send_msg = msg
    with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as s:
        s.ehlo(); s.starttls(); s.login(SMTP_USER, SMTP_PASS)
        s.sendmail(SMTP_USER, [DEST_EMAIL], send_msg.as_string())

# ── Inicialização do estoque (compatível com gunicorn) ────────────────────────
init_stock()

# ── Start ─────────────────────────────────────────────────────────────────────
if __name__ == '__main__':
    log.info(f'\n🚀 Lojinha            → http://localhost:{PORT}')
    log.info(f'🔧 Painel de pedidos  → http://localhost:{PORT}/admin')
    log.info(f'📧 Destino do email   → {DEST_EMAIL}')
    log.info(f'📬 SMTP               → {"configurado ("+SMTP_USER+")" if SMTP_USER else "não configurado — pedidos salvos em data/orders.json"}\n')
    app.run(host='0.0.0.0', port=PORT, debug=False)
