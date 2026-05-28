"""
Porra Mundial 2026 - Versión Render
  / y /clasificacion -> público
  /admin/*           -> protegido con login
"""
import os, json, unicodedata, secrets
from flask import Flask, request, redirect, session, render_template_string, url_for
import openpyxl

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", secrets.token_hex(32))

ADMIN_USER = os.environ.get("ADMIN_USER", "admin")
ADMIN_PASS = os.environ.get("ADMIN_PASS", "porristas26")

# ── Rutas de datos persistentes ───────────────────────────────────────────────
BASE_DIR    = os.path.dirname(__file__)
DATA_DIR    = os.path.join(BASE_DIR, "data")
UPLOAD_DIR  = os.path.join(BASE_DIR, "uploads")
STATE_FILE  = os.path.join(DATA_DIR, "state.json")
os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(UPLOAD_DIR, exist_ok=True)

# ── Estructura del Excel ──────────────────────────────────────────────────────
PARTIDOS_GRUPOS   = list(range(6, 78))
POSICIONES_GRUPOS = list(range(80, 128))
DIECISEISAVOS_C   = list(range(130, 162))
OCTAVOS_C         = list(range(182, 198))
CUARTOS_C         = list(range(210, 218))
SEMIS_C           = list(range(226, 230))
FINALISTAS_C      = [240, 241]
CAMPEON_C         = 250

BAREMO_DEFAULT = {
    "resultado_exacto": 6, "resultado_un_gol": 4, "resultado_signo": 3,
    "un_gol_falla": 1, "posicion_grupo": 2, "dieciseisavos": 3,
    "octavos": 6, "cuartos": 10, "semifinales": 15, "finalistas": 20, "campeon": 40,
}

# ── Persistencia ──────────────────────────────────────────────────────────────
def load_state():
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            s = json.load(f)
        # Asegurar claves por si el archivo es de versión anterior
        s.setdefault("jugadores", {})
        s.setdefault("oficiales", {})
        s.setdefault("baremo", BAREMO_DEFAULT.copy())
        s.setdefault("nombres_partidos", {})
        s.setdefault("nombres_posiciones", {})
        return s
    return {
        "jugadores": {}, "oficiales": {}, "baremo": BAREMO_DEFAULT.copy(),
        "nombres_partidos": {}, "nombres_posiciones": {}
    }

def save_state(s):
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(s, f, ensure_ascii=False, indent=2)

# ── Lógica Excel ──────────────────────────────────────────────────────────────
def leer_celda(ws, fila):
    val = ws.cell(row=fila, column=3).value
    return str(val).strip() if val is not None else ""

def signo_resultado(g1, g2):
    if g1 > g2: return "1"
    if g1 < g2: return "2"
    return "X"

def parse_resultado(txt):
    if not txt or txt in ("-", "0|-", ""): return None
    t = txt.strip()
    if "·" in t: t = t.split("·")[1]
    if "|" in t: t = t.split("|")[1].strip()
    if "-" in t:
        g = t.split("-")
        try:
            g1, g2 = int(g[0]), int(g[1])
            return (signo_resultado(g1, g2), g1, g2)
        except: return None
    return None

def normalizar(s):
    if not s: return ""
    s = str(s).strip().upper()
    s = unicodedata.normalize("NFD", s)
    return "".join(c for c in s if unicodedata.category(c) != "Mn")

def extraer_jugador(ruta):
    try:
        wb = openpyxl.load_workbook(ruta, data_only=True)
    except Exception as e:
        return None, str(e)
    if "Pool" not in wb.sheetnames:
        return None, "No tiene hoja Pool"
    ws = wb["Pool"]
    nombre = leer_celda(ws, 5)
    if not nombre or nombre in ("Nombre", ""):
        nombre = os.path.splitext(os.path.basename(ruta))[0]
    datos = {
        "nombre": nombre,
        "partidos":      {str(f): leer_celda(ws, f) for f in PARTIDOS_GRUPOS},
        "posiciones":    {str(f): leer_celda(ws, f) for f in POSICIONES_GRUPOS},
        "dieciseisavos": {str(f): leer_celda(ws, f) for f in DIECISEISAVOS_C},
        "octavos":       {str(f): leer_celda(ws, f) for f in OCTAVOS_C},
        "cuartos":       {str(f): leer_celda(ws, f) for f in CUARTOS_C},
        "semis":         {str(f): leer_celda(ws, f) for f in SEMIS_C},
        "finalistas":    {str(f): leer_celda(ws, f) for f in FINALISTAS_C},
        "campeon":       leer_celda(ws, CAMPEON_C),
    }
    nombres_p = {str(f): (ws.cell(row=f, column=2).value or f"Partido {f}") for f in PARTIDOS_GRUPOS}
    nombres_pos = {str(f): (ws.cell(row=f, column=2).value or f"Posición {f}") for f in POSICIONES_GRUPOS}
    return datos, None, nombres_p, nombres_pos

def calcular_puntos(jugador, oficiales, baremo):
    pts = 0; det = {}
    # Partidos grupo
    p = 0; exactos = 0
    for f, pred in jugador["partidos"].items():
        of = oficiales.get(f"p{f}", "")
        if not of: continue
        r_of = parse_resultado(of); r_p = parse_resultado(pred)
        if not r_of or not r_p: continue
        if r_of[1]==r_p[1] and r_of[2]==r_p[2]:
            p += baremo["resultado_exacto"]; exactos += 1
        elif r_of[0]==r_p[0] and (r_of[1]==r_p[1] or r_of[2]==r_p[2]): p += baremo["resultado_un_gol"]
        elif r_of[0]==r_p[0]: p += baremo["resultado_signo"]
        elif r_of[1]==r_p[1] or r_of[2]==r_p[2]: p += baremo["un_gol_falla"]
    det["partidos"] = p; det["exactos"] = exactos; pts += p
    # Posiciones grupo
    pp = 0
    for f, pred in jugador["posiciones"].items():
        of = oficiales.get(f"pos{f}", "")
        if of and pred and normalizar(of)==normalizar(pred): pp += baremo["posicion_grupo"]
    det["posiciones"] = pp; pts += pp
    # Fases eliminatorias
    def fase_pts(key_jugador, prefix, baremo_key):
        p2 = 0
        of_set = set(normalizar(v) for k,v in oficiales.items() if k.startswith(prefix) and v)
        for pred in jugador[key_jugador].values():
            if normalizar(pred) in of_set: p2 += baremo[baremo_key]
        return p2
    det["dieciseisavos"] = fase_pts("dieciseisavos","d16_","dieciseisavos"); pts += det["dieciseisavos"]
    det["octavos"]       = fase_pts("octavos","oct_","octavos");             pts += det["octavos"]
    det["cuartos"]       = fase_pts("cuartos","cua_","cuartos");             pts += det["cuartos"]
    det["semis"]         = fase_pts("semis","sem_","semifinales");           pts += det["semis"]
    det["finalistas"]    = fase_pts("finalistas","fin_","finalistas");       pts += det["finalistas"]
    # Campeón
    of_cam = normalizar(oficiales.get("campeon",""))
    cam = baremo["campeon"] if of_cam and normalizar(jugador["campeon"])==of_cam else 0
    det["campeon"] = cam; pts += cam
    det["total"] = pts
    return det

# ── CSS y HTML base ───────────────────────────────────────────────────────────
CSS = """
* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: 'Segoe UI', Arial, sans-serif; background: #0a1628; color: #e0e0e0; min-height: 100vh; }
header { background: linear-gradient(135deg, #1a3a6e, #c8102e); padding: 14px 20px; display: flex; align-items: center; gap: 12px; }
header h1 { font-size: 1.3em; color: white; }
nav { background: #0d1f3c; display: flex; gap: 2px; padding: 0 10px; border-bottom: 2px solid #c8102e; flex-wrap: wrap; }
nav a { padding: 10px 14px; color: #aaa; text-decoration: none; font-weight: 600; font-size: 0.83em; white-space: nowrap; }
nav a:hover, nav a.active { color: white; background: #1a3a6e; }
main { max-width: 1100px; margin: 20px auto; padding: 0 12px; }
.card { background: #0d1f3c; border: 1px solid #1a3a6e; border-radius: 10px; padding: 18px; margin-bottom: 18px; }
.card h2 { color: #4a9eff; margin-bottom: 14px; font-size: 1.02em; }
.btn { padding: 9px 18px; border: none; border-radius: 6px; cursor: pointer; font-weight: 600; font-size: 0.88em; text-decoration: none; display: inline-block; }
.btn-red { background: #c8102e; color: white; } .btn-red:hover { background: #a00d24; }
.btn-green { background: #1a7a3e; color: white; } .btn-green:hover { background: #145f30; }
.btn-blue { background: #1a3a6e; color: white; }
input[type=text], input[type=password], input[type=number] { background: #0a1628; border: 1px solid #1a3a6e; color: #e0e0e0; padding: 7px 10px; border-radius: 5px; }
input[type=file] { color: #aaa; }
label { display: block; margin-bottom: 4px; color: #aaa; font-size: 0.85em; }
.form-row { margin-bottom: 10px; }
.grid-2 { display: grid; grid-template-columns: 1fr 180px; gap: 10px; align-items: center; }
.tabla-wrap { overflow-x: auto; -webkit-overflow-scrolling: touch; }
.tabla { width: 100%; border-collapse: collapse; font-size: 0.82em; min-width: 540px; }
.tabla th { background: #1a3a6e; padding: 9px 5px; text-align: center; color: #4a9eff; font-size: 0.8em; white-space: nowrap; }
.tabla td { padding: 8px 5px; border-bottom: 1px solid #1a3a6e; text-align: center; white-space: nowrap; }
.tabla tr:hover td { background: #111f3a; }
.nombre { text-align: left !important; font-weight: 600; color: white; white-space: normal !important; min-width: 90px; }
.total { font-weight: 700; color: #ffd700; font-size: 1.05em; }
.exactos { color: #4adf7a; font-weight: 600; }
.pos-medal { font-size: 1.05em; }
.msg-ok { padding: 10px 15px; border-radius: 6px; margin-bottom: 14px; background: #1a3a1a; border: 1px solid #1a7a3e; color: #4adf7a; }
.msg-err { padding: 10px 15px; border-radius: 6px; margin-bottom: 14px; background: #3a1a1a; border: 1px solid #7a1a1a; color: #df4a4a; }
.sep { border: none; border-top: 1px solid #1a3a6e; margin: 18px 0; }
.baremo-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(200px, 1fr)); gap: 12px; }
.baremo-item label { font-size: 0.8em; }
.baremo-item input { width: 70px; }
.tag { background: #1a3a6e; padding: 4px 11px; border-radius: 20px; font-size: 0.82em; display: inline-block; margin: 3px; }
details summary { cursor: pointer; color: #4a9eff; font-weight: 600; padding: 8px 0; user-select: none; }
.login-box { max-width: 360px; margin: 60px auto; }
@media (max-width: 500px) {
  header h1 { font-size: 1.05em; }
  nav a { padding: 8px 9px; font-size: 0.76em; }
  .card { padding: 12px; }
  .grid-2 { grid-template-columns: 1fr; }
}
"""

def base(content, tab="pub", admin=False):
    nav_pub = f'<a href="/" class="{"active" if tab=="pub" else ""}">🏆 Clasificación</a>'
    nav_admin = ""
    if admin:
        nav_admin = f'''
        <a href="/admin" class="{"active" if tab=="home" else ""}">🏠 Admin</a>
        <a href="/admin/cargar" class="{"active" if tab=="cargar" else ""}">📂 Cargar Porras</a>
        <a href="/admin/resultados" class="{"active" if tab=="resultados" else ""}">📋 Resultados</a>
        <a href="/admin/baremo" class="{"active" if tab=="baremo" else ""}">⚙️ Baremo</a>
        <a href="/admin/logout" style="margin-left:auto;color:#c8102e">Salir</a>'''
    else:
        nav_admin = '<a href="/admin" style="margin-left:auto;color:#4a9eff;font-size:0.8em">Admin</a>'
    return f"""<!DOCTYPE html>
<html lang="es"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>🏆 Porra Mundial 2026</title>
<style>{CSS}</style></head>
<body>
<header><span style="font-size:2em">⚽</span><h1>Porra Mundial 2026</h1></header>
<nav>{nav_pub}{nav_admin}</nav>
<main>{content}</main>
</body></html>"""

# ── Auth ──────────────────────────────────────────────────────────────────────
def is_admin():
    return session.get("admin") is True

def require_admin():
    if not is_admin():
        return redirect("/admin/login")
    return None

# ── Rutas públicas ────────────────────────────────────────────────────────────
@app.route("/")
def clasificacion():
    s = load_state()
    if not s["jugadores"]:
        content = '<div class="card"><h2>🏆 Clasificación</h2><p style="color:#aaa">Todavía no hay participantes cargados.</p></div>'
        return base(content, "pub")

    resultados = []
    for nombre, datos in s["jugadores"].items():
        det = calcular_puntos(datos, s["oficiales"], s["baremo"])
        resultados.append((nombre, det))
    resultados.sort(key=lambda x: -x[1]["total"])

    medallas = ["🥇","🥈","🥉"]
    rows = ""
    for i, (nombre, det) in enumerate(resultados):
        medal = medallas[i] if i < 3 else f"{i+1}º"
        rows += f"""<tr>
            <td class="pos-medal">{medal}</td>
            <td class="nombre">{nombre}</td>
            <td>{det['partidos']}</td>
            <td class="exactos" title="Resultados exactos">🎯{det['exactos']}</td>
            <td>{det['posiciones']}</td>
            <td>{det['dieciseisavos']}</td><td>{det['octavos']}</td>
            <td>{det['cuartos']}</td><td>{det['semis']}</td>
            <td>{det['finalistas']}</td><td>{det['campeon']}</td>
            <td class="total">{det['total']}</td>
        </tr>"""

    content = f"""<div class="card">
        <h2>🏆 Clasificación General — {len(resultados)} participantes</h2>
        <div class="tabla-wrap">
        <table class="tabla">
            <thead><tr>
                <th>#</th><th style="text-align:left">Jugador</th>
                <th>Pts<br>Partidos</th><th>🎯<br>Exactos</th><th>Grupos</th>
                <th>1/16</th><th>1/8</th><th>1/4</th><th>1/2</th>
                <th>Final</th><th>Campeón</th><th>TOTAL</th>
            </tr></thead>
            <tbody>{rows}</tbody>
        </table>
        </div>
    </div>"""
    return base(content, "pub")

# ── Rutas admin ───────────────────────────────────────────────────────────────
@app.route("/admin/login", methods=["GET","POST"])
def login():
    error = ""
    if request.method == "POST":
        if request.form.get("user")==ADMIN_USER and request.form.get("pass")==ADMIN_PASS:
            session["admin"] = True
            return redirect("/admin")
        error = "Usuario o contraseña incorrectos"
    content = f"""<div class="card login-box">
        <h2>🔐 Acceso Admin</h2>
        {"<div class='msg-err'>"+error+"</div>" if error else ""}
        <form method="post">
            <div class="form-row"><label>Usuario</label><input type="text" name="user" style="width:100%"></div>
            <div class="form-row"><label>Contraseña</label><input type="password" name="pass" style="width:100%"></div>
            <br><button class="btn btn-red" type="submit">Entrar</button>
        </form>
    </div>"""
    return base(content, "login")

@app.route("/admin/logout")
def logout():
    session.clear()
    return redirect("/")

@app.route("/admin")
def admin_home():
    r = require_admin()
    if r: return r
    s = load_state()
    n_j = len(s["jugadores"])
    n_of = sum(1 for v in s["oficiales"].values() if v)
    content = f"""<div class="card">
        <h2>🏠 Panel de Administración</h2>
        <p style="color:#aaa;margin-bottom:16px">Bienvenido, <b>{ADMIN_USER}</b></p>
        <p>👥 Participantes cargados: <b style="color:#ffd700">{n_j}</b></p>
        <p style="margin-top:8px">📋 Resultados introducidos: <b style="color:#4adf7a">{n_of}</b></p>
        <hr class="sep">
        <a href="/admin/cargar" class="btn btn-blue" style="margin-right:8px">📂 Cargar Porras</a>
        <a href="/admin/resultados" class="btn btn-blue" style="margin-right:8px">📋 Resultados</a>
        <a href="/admin/baremo" class="btn btn-blue">⚙️ Baremo</a>
    </div>"""
    return base(content, "home", admin=True)

@app.route("/admin/cargar", methods=["GET","POST"])
def admin_cargar():
    r = require_admin()
    if r: return r
    s = load_state()
    msg = ""; msg_class = "msg-ok"

    if request.method == "POST":
        if "limpiar" in request.form:
            s["jugadores"].clear()
            s["nombres_partidos"] = {}
            s["nombres_posiciones"] = {}
            save_state(s)
            msg = "✅ Todos los jugadores eliminados"
        else:
            archivos = request.files.getlist("archivos")
            cargados = []; errores = []
            for f in archivos:
                if not f.filename.endswith(".xlsx"): continue
                ruta = os.path.join(UPLOAD_DIR, f.filename)
                f.save(ruta)
                resultado = extraer_jugador(ruta)
                if resultado[1]:
                    errores.append(f"{f.filename}: {resultado[1]}")
                else:
                    datos, _, np, npos = resultado
                    s["jugadores"][datos["nombre"]] = datos
                    if not s["nombres_partidos"]:
                        s["nombres_partidos"] = np
                        s["nombres_posiciones"] = npos
                    cargados.append(datos["nombre"])
            save_state(s)
            if cargados:
                msg = f"✅ Cargados: {', '.join(cargados)}"
                if errores: msg += f" | ⚠️ {', '.join(errores)}"
            else:
                msg = f"⚠️ {'; '.join(errores)}"; msg_class = "msg-err"

    tags = "".join(f'<span class="tag">✔ {n}</span>' for n in s["jugadores"])
    lista = tags or "<p style='color:#666'>Ninguno todavía</p>"
    content = f"""<div class="card">
        <h2>📂 Cargar archivos de porra</h2>
        {"<div class='"+msg_class+"'>"+msg+"</div>" if msg else ""}
        <form method="post" enctype="multipart/form-data">
            <div class="form-row">
                <label>Selecciona uno o varios archivos Excel (.xlsx)</label>
                <input type="file" name="archivos" multiple accept=".xlsx">
            </div>
            <button class="btn btn-red" type="submit">⬆ Cargar</button>
        </form>
        <hr class="sep">
        <h2>Participantes ({len(s["jugadores"])})</h2>
        <div style="margin-bottom:16px">{lista}</div>
        <form method="post">
            <button class="btn btn-blue" name="limpiar" value="1"
                onclick="return confirm('¿Borrar todos los jugadores?')">🗑 Limpiar todo</button>
        </form>
    </div>"""
    return base(content, "cargar", admin=True)

@app.route("/admin/resultados", methods=["GET","POST"])
def admin_resultados():
    r = require_admin()
    if r: return r
    s = load_state()
    msg = ""
    if request.method == "POST":
        for k, v in request.form.items():
            s["oficiales"][k] = v.strip()
        save_state(s)
        msg = "✅ Resultados guardados"

    np = s.get("nombres_partidos", {})
    npos = s.get("nombres_posiciones", {})

    partidos_html = ""
    for f in PARTIDOS_GRUPOS:
        nombre = np.get(str(f), f"Partido {f}")
        val = s["oficiales"].get(f"p{f}", "")
        partidos_html += f"""<div class="form-row grid-2">
            <label style="color:#ddd;font-size:0.88em">{nombre}</label>
            <input type="text" name="p{f}" value="{val}" placeholder="2-1" style="width:120px">
        </div>"""

    pos_html = ""
    for f in POSICIONES_GRUPOS:
        nombre = npos.get(str(f), f"Posición {f}")
        val = s["oficiales"].get(f"pos{f}", "")
        pos_html += f"""<div class="form-row grid-2">
            <label style="color:#ddd;font-size:0.88em">{nombre}</label>
            <input type="text" name="pos{f}" value="{val}" placeholder="Equipo" style="width:160px">
        </div>"""

    def fase_inputs(filas, prefix, label):
        h = f"<h3 style='color:#4a9eff;margin:14px 0 8px'>{label}</h3>"
        for f in filas:
            val = s["oficiales"].get(f"{prefix}{f}", "")
            h += f"""<div class="form-row grid-2">
                <label style="color:#aaa;font-size:0.83em">Clasificado</label>
                <input type="text" name="{prefix}{f}" value="{val}" placeholder="Equipo" style="width:160px">
            </div>"""
        return h

    campeon_val = s["oficiales"].get("campeon", "")
    content = f"""<div class="card">
        <h2>📋 Resultados Oficiales</h2>
        {"<div class='msg-ok'>"+msg+"</div>" if msg else ""}
        <form method="post">
        <details open><summary>▶ Partidos Fase de Grupos</summary>
            <div style="margin-top:8px">{partidos_html}</div>
        </details>
        <hr class="sep">
        <details><summary>▶ Clasificaciones de Grupos</summary>
            <div style="margin-top:8px">{pos_html}</div>
        </details>
        <hr class="sep">
        <details><summary>▶ Eliminatorias</summary>
            <div style="margin-top:8px">
            {fase_inputs(DIECISEISAVOS_C,"d16_","Clasificados a Dieciseisavos")}
            {fase_inputs(OCTAVOS_C,"oct_","Clasificados a Octavos")}
            {fase_inputs(CUARTOS_C,"cua_","Clasificados a Cuartos")}
            {fase_inputs(SEMIS_C,"sem_","Semifinalistas")}
            {fase_inputs(FINALISTAS_C,"fin_","Finalistas")}
            <h3 style='color:#4a9eff;margin:14px 0 8px'>🥇 Campeón</h3>
            <div class="form-row grid-2">
                <label style="color:#ddd">Campeón del Mundial</label>
                <input type="text" name="campeon" value="{campeon_val}" placeholder="Equipo" style="width:160px">
            </div>
            </div>
        </details>
        <hr class="sep">
        <button class="btn btn-green" type="submit">💾 Guardar resultados</button>
        </form>
    </div>"""
    return base(content, "resultados", admin=True)

@app.route("/admin/baremo", methods=["GET","POST"])
def admin_baremo():
    r = require_admin()
    if r: return r
    s = load_state()
    msg = ""
    if request.method == "POST":
        for k in BAREMO_DEFAULT:
            try: s["baremo"][k] = int(request.form.get(k, s["baremo"][k]))
            except: pass
        save_state(s)
        msg = "✅ Baremo guardado"

    items = [
        ("resultado_exacto","Resultado exacto (2-1 ✓ 2-1)"),
        ("resultado_un_gol","Resultado + un gol (2-1 ✓ 2-0)"),
        ("resultado_signo","Signo acertado (2-1 ✓ 1-0)"),
        ("un_gol_falla","Falla resultado, acierta un gol"),
        ("posicion_grupo","Posición en grupo acertada"),
        ("dieciseisavos","Clasificado a dieciseisavos"),
        ("octavos","Clasificado a octavos"),
        ("cuartos","Clasificado a cuartos"),
        ("semifinales","Semifinalista"),
        ("finalistas","Finalista"),
        ("campeon","Campeón"),
    ]
    inputs = "".join(f"""<div class="baremo-item">
        <label>{desc}</label>
        <input type="number" name="{key}" value="{s['baremo'][key]}" min="0" max="999">
    </div>""" for key, desc in items)
    content = f"""<div class="card">
        <h2>⚙️ Baremo de puntuación</h2>
        {"<div class='msg-ok'>"+msg+"</div>" if msg else ""}
        <form method="post">
        <div class="baremo-grid">{inputs}</div>
        <hr class="sep">
        <button class="btn btn-green" type="submit">💾 Guardar baremo</button>
        </form>
    </div>"""
    return base(content, "baremo", admin=True)

if __name__ == "__main__":
    import webbrowser
    webbrowser.open("http://localhost:5000")
    app.run(debug=True, port=5000)