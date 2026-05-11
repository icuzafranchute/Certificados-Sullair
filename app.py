import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.drawing.image import Image as XLImage
import os, re, io, hashlib
from datetime import datetime, date

st.set_page_config(page_title="Sullair — Certificados", page_icon="📄", layout="centered")

def _hash(pw): return hashlib.sha256(pw.encode()).hexdigest()
USUARIOS = {
    "AMASTRACCI": _hash("sullair2026"),
    "JICUZA":     _hash("sullair2026"),
    "ETOLOSA":    _hash("sullair2026"),
    "FCENDRA":    _hash("sullair2026"),
}

SUCURSALES = [
    "Neuquén","Buenos Aires","Mendoza","Bahía Blanca",
    "Comodoro Rivadavia","Salta","Tucumán","San Juan",
    "Córdoba","Mar del Plata","Olavarría",
]

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@700;800&family=DM+Sans:wght@300;400;500;600&display=swap');
html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }
.titulo { font-family:'Syne',sans-serif; font-size:1.5rem; font-weight:800; color:#00783C; }
.sub    { font-size:.82rem; color:#888; margin-bottom:20px; }
.card   { background:#fff; border:1.5px solid #E0E0E0; border-radius:14px;
          padding:22px 26px; margin-bottom:16px; box-shadow:0 2px 10px rgba(0,0,0,.05); }
.card h3{ font-family:'Syne',sans-serif; color:#00783C; font-size:.95rem; margin-bottom:12px; }
.stat   { background:#E8F5EE; border:1.5px solid #b2dfc5; border-radius:10px;
          padding:12px; text-align:center; }
.stat .n{ font-family:'Syne',sans-serif; font-size:1.8rem; font-weight:800; color:#00783C; display:block; }
.stat .l{ font-size:.72rem; color:#555; }
.stButton>button { background:#00783C!important; color:#fff!important;
    font-family:'Syne',sans-serif!important; font-weight:700!important;
    border-radius:10px!important; border:none!important; width:100%; }
div[data-testid="stDownloadButton"]>button {
    background:#fff!important; color:#00783C!important;
    border:2px solid #00783C!important; font-family:'Syne',sans-serif!important;
    font-weight:700!important; border-radius:10px!important; width:100%; }
div[data-testid="stDownloadButton"]>button:hover { background:#00783C!important; color:#fff!important; }
</style>
""", unsafe_allow_html=True)

# ── LOGIN ─────────────────────────────────────────────────────────────────────
def mostrar_login():
    st.markdown("""
    <style>
    .login-box { max-width:360px; margin:60px auto 0; background:#fff;
        border-radius:16px; padding:40px 36px; box-shadow:0 4px 24px rgba(0,0,0,.10);
        border:1.5px solid #E0E0E0; text-align:center; }
    .login-titulo { font-family:'Syne',sans-serif; font-size:1.4rem; font-weight:800; color:#00783C; }
    .login-sub { font-size:.82rem; color:#888; margin-bottom:24px; }
    </style>""", unsafe_allow_html=True)
    _, col, _ = st.columns([1, 2, 1])
    with col:
        st.markdown('<div class="login-box">', unsafe_allow_html=True)
        st.markdown('<p class="login-titulo">📄 Sullair Argentina</p>', unsafe_allow_html=True)
        st.markdown('<p class="login-sub">Generador de Certificados</p>', unsafe_allow_html=True)
        usuario  = st.text_input("Usuario", placeholder="Ej: JICUZA").strip().upper()
        password = st.text_input("Contraseña", type="password")
        if st.button("Ingresar", use_container_width=True):
            if usuario in USUARIOS and USUARIOS[usuario] == _hash(password):
                st.session_state.update(logged_in=True, usuario=usuario)
                st.rerun()
            else:
                st.error("Usuario o contraseña incorrectos.")
        st.markdown('</div>', unsafe_allow_html=True)

if not st.session_state.get('logged_in'):
    mostrar_login(); st.stop()

# ── HELPERS ───────────────────────────────────────────────────────────────────
BASE_DIR  = os.path.dirname(os.path.abspath(__file__))
LOGO_PATH = os.path.join(BASE_DIR, "assets", "logo_sullair.jpg")
VERDE     = "FF71AF47"
BLANCO    = "FFFFFFFF"
NEGRO     = "FF000000"

def med():  return Side(border_style='medium', color=NEGRO)
def thin(): return Side(border_style='thin',   color=NEGRO)
def nob():  return Side(border_style=None)

def fmt_fecha(dt):
    if dt is None: return ''
    if isinstance(dt, (datetime, date)): return dt.strftime('%d/%m/%Y')
    return str(dt)

def limpiar(s): return re.sub(r'[\\/*?:"<>|]', '', str(s).strip()).replace('  ', ' ')

FAMILIA_MAP = {
    '70R':'Generador','120R':'Generador','45KVA':'Generador',
    '375Q':'Compresor','375QH':'Compresor','185Q':'Compresor','3007':'Compresor',
    '860SJ':'Plataforma','1200SJP':'Plataforma',
    'STH1256':'Manipulador','AMT':'Manipulador','540.170':'Manipulador','4017':'Manipulador',
    'ATABDIS':'Luminaria',
}
def clasificar(art, interno):
    for k, v in FAMILIA_MAP.items():
        if k.upper() in str(art).upper(): return v
    i = str(interno).strip().upper()
    if i.startswith(('E01','A01')): return 'Compresor'
    if i.startswith(('E02','A02')): return 'Generador'
    if i.startswith('E03'):         return 'Plataforma'
    if i.startswith(('E04','A04')): return 'Manipulador'
    if i.startswith('E05'):         return 'Luminaria'
    return 'Equipo'

# ── LEER VOUCHERS ─────────────────────────────────────────────────────────────
def leer_vouchers(archivo_bytes):
    df = pd.read_excel(io.BytesIO(archivo_bytes), header=None)
    periodo_raw = str(df.iloc[0, 0]).strip() if pd.notna(df.iloc[0, 0]) else ''
    headers = [str(v).strip() if pd.notna(v) else f'col_{i}' for i, v in enumerate(df.iloc[1])]
    datos = df.iloc[2:].copy()
    datos.columns = headers
    datos = datos.dropna(subset=['Artículo', 'Interno'], how='all')
    return datos, periodo_raw

# ── CREAR HOJA — REPLICA EXACTA DEL MODELO ───────────────────────────────────
def crear_hoja(ws, cert_num, sucursal, direccion, periodo,
               cliente, articulo, interno, equipo, inicio, fin,
               cnt_dias, precio, total, moneda):

    ws.sheet_view.showGridLines = False

    # Columnas exactas del modelo
    for col, w in {'A':7.14,'B':4.14,'C':15.43,'D':11.57,'E':8.29,
                   'F':15.0,'G':14.57,'H':12.14,'I':20.57,'J':23.14,
                   'K':6.71,'L':8.43}.items():
        ws.column_dimensions[col].width = w

    # Alturas exactas del modelo
    for r, h in {3:15.75,4:14.25,5:18.75,6:15.75,7:15.0,
                 8:15.75,9:15.75,10:15.75,11:15.75,
                 12:16.5,13:15.75,14:15.75,
                 17:27.75,18:24.0,19:14.25,20:15.75,21:28.5,22:28.5}.items():
        ws.row_dimensions[r].height = h

    # Borde exterior: top en fila 4, left en col B, right en col K, bottom en fila 20
    for col in ['B','C','D','E','F','G','H','I','J','K']:
        c = ws[f'{col}4']
        c.border = Border(top=med(),
                          left=med()  if col == 'B' else nob(),
                          right=med() if col == 'K' else nob())
    for row in range(5, 20):
        ws[f'B{row}'].border = Border(left=med())
        ws[f'K{row}'].border = Border(right=med())
    for col in ['B','C','D','E','F','G','H','I','J','K']:
        c = ws[f'{col}20']
        c.border = Border(bottom=med(),
                          left=med()  if col == 'B' else nob(),
                          right=med() if col == 'K' else nob())

    # Logo — posición exacta del modelo: merge H4:K7, anchor H4, 253x53px
    ws.merge_cells('H4:K7')
    if os.path.exists(LOGO_PATH):
        try:
            img = XLImage(LOGO_PATH)
            img.width = 253; img.height = 53; img.anchor = 'H4'
            ws.add_image(img)
        except Exception:
            pass

    # Título
    ws.merge_cells('C5:E5')
    c = ws['C5']; c.value = 'CERTIFICACIÓN DE SERVICIOS'
    c.font = Font(name='Calibri', bold=True, size=14)
    c.alignment = Alignment(horizontal='left')

    ws.merge_cells('C6:D6')
    c = ws['C6']; c.value = 'Alquiler de equipos'
    c.font = Font(name='Calibri', size=12)
    c.alignment = Alignment(horizontal='left')

    # Encabezado
    def lbl(r, col, txt):
        c = ws.cell(r, col, txt)
        c.font = Font(name='Calibri', bold=True, size=12)

    def val(r, col, txt, h='left'):
        c = ws.cell(r, col, txt)
        c.font = Font(name='Calibri', size=11)
        c.alignment = Alignment(horizontal=h)

    lbl(8, 3, 'Certificado N°');  val(8, 4, cert_num);    val(8, 9, sucursal)
    lbl(9, 3, 'Cliente ')
    ws.merge_cells('D9:G9');      val(9, 4, cliente);     val(9, 9, direccion)
    lbl(10, 3, 'Período');        val(10, 4, periodo)

    # Tabla — encabezado verde
    fv  = PatternFill('solid', fgColor=VERDE)
    fhb = Font(name='Calibri', bold=True, size=12, color=BLANCO)
    ac  = Alignment(horizontal='center', vertical='center')

    hdr_data = [
        (3, 'Equipo',             Border(top=med(),bottom=med(),left=med(), right=thin())),
        (4, 'Modelo',             Border(top=med(),bottom=med(),left=thin(),right=thin())),
        (5, 'Interno',            Border(top=med(),bottom=med(),left=thin(),right=thin())),
        (6, 'Inicio Periodo',     Border(top=med(),bottom=med(),left=thin(),right=thin())),
        (7, 'Fin Periodo',        Border(top=med(),bottom=med(),left=thin(),right=thin())),
        (8, 'Dias',               Border(top=med(),bottom=med(),left=thin(),right=thin())),
        (9, 'Valor diario (USD)', Border(top=med(),bottom=med(),left=thin(),right=thin())),
        (10,'Total (U$S)',        Border(top=med(),bottom=med(),left=thin(),right=med())),
    ]
    for col_n, txt, brd in hdr_data:
        c = ws.cell(12, col_n, txt)
        c.fill = fv; c.font = fhb; c.alignment = ac; c.border = brd

    # Tabla — fila de datos
    fd      = Font(name='Calibri', size=11)
    num_fmt = f'"{moneda}" #,##0.00'
    dat_borders = [
        Border(bottom=med(),left=med(), right=thin()),
        Border(bottom=med(),left=thin(),right=thin()),
        Border(bottom=med(),left=thin(),right=thin()),
        Border(bottom=med(),left=thin(),right=thin()),
        Border(bottom=med(),left=thin(),right=thin()),
        Border(bottom=med(),left=thin(),right=thin()),
        Border(bottom=med(),left=thin(),right=thin()),
        Border(bottom=med(),left=thin(),right=med()),
    ]
    dat_vals = [equipo, articulo.strip(), interno.strip(),
                fmt_fecha(inicio), fmt_fecha(fin),
                cnt_dias, precio, total]

    for col_n, v, brd in zip(range(3, 11), dat_vals, dat_borders):
        c = ws.cell(13, col_n, v)
        c.font = fd; c.alignment = ac; c.border = brd
        if col_n == 9:  c.number_format = num_fmt
        if col_n == 10: c.number_format = num_fmt

    # Fila 14 — celdas vacías C-H + SUBTOTAL en I + total en J
    for col_n in range(3, 9):
        c = ws.cell(14, col_n)
        c.font = Font(name='Calibri', size=11)
        c.alignment = Alignment(vertical='center')
        c.border = Border(bottom=med(), left=med() if col_n == 3 else nob())

    c = ws.cell(14, 9, 'SUBTOTAL')
    c.font = Font(name='Calibri', bold=True, size=11)
    c.alignment = Alignment(horizontal='center', vertical='center')
    c.border = Border(bottom=med())

    c = ws.cell(14, 10, total)
    c.font = Font(name='Calibri', bold=True, size=11)
    c.number_format = num_fmt
    c.alignment = Alignment(horizontal='center', vertical='center')
    c.border = Border(bottom=med(), right=med())

    # Observaciones
    c = ws.cell(17, 3, 'OBSERVACIONES:')
    c.font = Font(name='Calibri', bold=True, size=12)

    ws.merge_cells('C18:G18')
    c = ws['C18']
    c.value = 'Los valores del presente certificado no incluyen IVA ni otros impuestos'
    c.font  = Font(name='Calibri', italic=True, size=11)
    c.alignment = Alignment(horizontal='left', vertical='center')

    ws.merge_cells('I18:J19')
    c = ws['I18']
    c.value = 'Firma cliente'
    c.font  = Font(name='Calibri', size=10)
    c.alignment = Alignment(horizontal='center', vertical='center')

# ── PROCESO PRINCIPAL ─────────────────────────────────────────────────────────
def procesar(archivo_bytes, cert_inicio, periodo_manual,
             sucursal, direccion, prog_bar, status_txt):

    datos, periodo_detectado = leer_vouchers(archivo_bytes)
    periodo = periodo_manual.strip() if periodo_manual.strip() else periodo_detectado

    wb = Workbook(); wb.remove(wb.active)
    cert_num = cert_inicio; generados = 0; log_lines = []
    total = len(datos)

    for idx, (_, row) in enumerate(datos.iterrows()):
        art  = str(row.get('Artículo', '')).strip()
        int_ = str(row.get('Interno',  '')).strip()
        if not art and not int_: continue

        cli  = str(row.get('Nombre', '')).strip()
        dias = int(row.get('Cnt Días', 0) or 0)
        ini  = row.get('Inicio Período')
        fin  = row.get('Fin Período')
        prec = float(row.get('Pcio Unitario', 0) or 0)
        tot  = float(row.get('Total', 0) or 0)
        mon  = str(row.get('Moneda', 'U$S')).strip().replace('(','').replace(')','').strip()
        eq   = clasificar(art, int_)

        nombre_hoja = limpiar(f"{int_[:12].strip()} - {cert_num}")[:31]
        ws = wb.create_sheet(title=nombre_hoja)
        crear_hoja(ws, cert_num, sucursal, direccion, periodo,
                   cli, art, int_, eq, ini, fin, dias, prec, tot, mon)

        log_lines.append(f"✅ Cert {cert_num:04d} | {int_:<12} | {art}")
        cert_num += 1; generados += 1

        prog_bar.progress(int((idx + 1) / total * 100))
        status_txt.text(f"Generando {idx + 1} de {total}...")

    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return {'xlsx_bytes': buf.read(), 'generados': generados,
            'cert_inicio': cert_inicio, 'cert_fin': cert_num - 1,
            'periodo': periodo, 'log': log_lines}

# ── UI ────────────────────────────────────────────────────────────────────────
col_t, col_logout = st.columns([4, 1])
with col_t:
    st.markdown('<p class="titulo">📄 Generador de Certificados</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="sub">Sullair Argentina S.A. &nbsp;|&nbsp; 👤 {st.session_state.get("usuario","")}</p>', unsafe_allow_html=True)
with col_logout:
    st.write("")
    if st.button("🚪 Salir"):
        st.session_state.update(logged_in=False, usuario=""); st.rerun()

# Paso 1
st.markdown('<div class="card"><h3>⚙️ Configuración</h3>', unsafe_allow_html=True)
c1, c2 = st.columns(2)
with c1: cert_inicio = st.number_input("N° de certificado inicial", min_value=1, value=100, step=1)
with c2: periodo_manual = st.text_input("Período", placeholder="Se detecta automáticamente del Excel")
sucursal_sel = st.selectbox("Sucursal", SUCURSALES)
direccion    = st.text_input("Dirección de la sucursal", placeholder="Ej: Av. San Martín 1234, Neuquén")
st.markdown('</div>', unsafe_allow_html=True)

# Paso 2
st.markdown('<div class="card"><h3>📊 Cargar listado de vouchers</h3>', unsafe_allow_html=True)
archivo = st.file_uploader("Arrastrá o seleccioná el Excel de vouchers", type=["xlsx","xls"], label_visibility="collapsed")
if archivo:
    st.success(f"✅ **{archivo.name}** cargado")
    datos_prev, periodo_det = leer_vouchers(archivo.read()); archivo.seek(0)
    st.info(f"📅 Período detectado: **{periodo_det}** &nbsp;|&nbsp; 🔢 **{len(datos_prev)}** vouchers")
    with st.expander("👁️ Vista previa — primeros 5 registros"):
        cols_ok = [c for c in ['Artículo','Interno','Cnt Días','Inicio Período','Fin Período','Pcio Unitario','Total'] if c in datos_prev.columns]
        st.dataframe(datos_prev[cols_ok].head(5), use_container_width=True)
st.markdown('</div>', unsafe_allow_html=True)

# Paso 3
st.markdown('<div class="card"><h3>⚡ Generar certificados</h3>', unsafe_allow_html=True)
btn_disabled = (archivo is None) or (not direccion.strip())
if archivo is not None and not direccion.strip():
    st.warning("⚠️ Completá la dirección de la sucursal para continuar.")
if st.button("⚡ Generar Excel con certificados", disabled=btn_disabled):
    prog = st.progress(0); status = st.empty()
    with st.spinner("Procesando..."):
        resultado = procesar(archivo.read(), int(cert_inicio), periodo_manual,
                             sucursal_sel, direccion.strip(), prog, status)
    prog.progress(100); status.text("✅ Completado")
    st.session_state["resultado"] = resultado
    with st.expander("📋 Detalle del proceso"):
        for line in resultado["log"]: st.text(line)
st.markdown('</div>', unsafe_allow_html=True)

# Paso 4
if "resultado" in st.session_state:
    r = st.session_state["resultado"]
    st.markdown('<div class="card"><h3>📥 Descargar resultado</h3>', unsafe_allow_html=True)
    c1,c2,c3,c4 = st.columns(4)
    with c1: st.markdown(f'<div class="stat"><span class="n">{r["generados"]}</span><span class="l">Certificados</span></div>', unsafe_allow_html=True)
    with c2: st.markdown(f'<div class="stat"><span class="n">{r["cert_inicio"]}</span><span class="l">N° inicial</span></div>', unsafe_allow_html=True)
    with c3: st.markdown(f'<div class="stat"><span class="n">{r["cert_fin"]}</span><span class="l">N° final</span></div>', unsafe_allow_html=True)
    with c4: st.markdown(f'<div class="stat"><span class="n" style="font-size:1rem">{r["periodo"]}</span><span class="l">Período</span></div>', unsafe_allow_html=True)
    st.write("")
    st.download_button(
        label="📥 Descargar Excel (un certificado por hoja)",
        data=r["xlsx_bytes"],
        file_name=f"Certificados_{limpiar(r['periodo'])}_{r['cert_inicio']}_{r['cert_fin']}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.markdown('</div>', unsafe_allow_html=True)

st.markdown("---")
st.markdown('<p style="text-align:center;color:#aaa;font-size:.72rem">Sullair Argentina S.A. · Generador Universal de Certificados · v3.0</p>', unsafe_allow_html=True)
