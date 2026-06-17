"""
=========================================================
  BITÁCORA DE BIOLÓGICOS  V6
  Jurisdicción 06 Xonacatlán · CESAEM
  NUEVOS MÓDULOS:
  - Anexo A: llenado automático desde Base entradas
  - Dashboard con segmentador de tiempo y gráficas
  - Reporte: salidas por biológico × elemento (destino/
    responsable) con segmentador de período
=========================================================
"""

import streamlit as st
import pandas as pd
import io
from datetime import date

try:
    from weasyprint import HTML as WpHTML
    WEASYPRINT_OK = True
except Exception:
    WEASYPRINT_OK = False

try:
    import xlsxwriter  # noqa
    XL_ENGINE = "xlsxwriter"
except ImportError:
    XL_ENGINE = "openpyxl"

# ═══════════════════════════════════════════════════════
st.set_page_config(page_title="Bitácora Biológicos – Dist. 06",
                   page_icon="💉", layout="wide")
st.markdown("""
<style>
.hdr{background:linear-gradient(135deg,#1a5276,#2e86c1);
  padding:.9rem 1.6rem;border-radius:10px;color:#fff;margin-bottom:1.2rem}
.hdr h1{margin:0;font-size:1.4rem}
.hdr p{margin:.2rem 0 0;font-size:.82rem;opacity:.85}
.kpi{background:white;border-left:5px solid #2e86c1;padding:.8rem 1rem;
  border-radius:8px;box-shadow:0 2px 6px rgba(0,0,0,.07)}
.kpi.rojo{border-left-color:#e74c3c}
.kpi.verde{border-left-color:#27ae60}
.kpi.naranja{border-left-color:#f39c12}
.kpi h2{margin:0;font-size:1.9rem;color:#1a5276}
.kpi p{margin:.1rem 0 0;font-size:.8rem;color:#666}
.anexo-hdr{background:#1a5276;color:white;text-align:center;
  padding:6px;font-weight:700;font-size:13px;margin-bottom:4px;border-radius:4px}
</style>""", unsafe_allow_html=True)

MESES_ES = {1:"Ene",2:"Feb",3:"Mar",4:"Abr",5:"May",6:"Jun",
            7:"Jul",8:"Ago",9:"Sep",10:"Oct",11:"Nov",12:"Dic"}
MESES_LARGO = {1:"ENERO",2:"FEBRERO",3:"MARZO",4:"ABRIL",5:"MAYO",6:"JUNIO",
               7:"JULIO",8:"AGOSTO",9:"SEPTIEMBRE",10:"OCTUBRE",
               11:"NOVIEMBRE",12:"DICIEMBRE"}

# Mapeo nombre entrada → nombre salida (donde difieren)
MAPA_BIO_E_S = {
    "TDPA":  "TDPa",
    "HEPATITIS A": "HEPATITIS a",
    "NEUMOCOCO 20": "NEUMOCOCO 13",
    "COVID 19 Moderma": "COVID ABDALA",
    "COVID 19 Pzfizer": "COVID ABDALA",
    "VSR": None,
}

# ═══════════════════════════════════════════════════════
#  CARGA
# ═══════════════════════════════════════════════════════
@st.cache_data(show_spinner=False)
def cargar(bts: bytes) -> dict:
    xl  = pd.ExcelFile(io.BytesIO(bts))
    nom = xl.sheet_names
    pe  = next((s for s in nom if "entrada" in s.lower()), None)
    ps  = next((s for s in nom if "salida"  in s.lower()), None)
    if not pe or not ps:
        st.error("❌ Faltan pestañas 'Base entradas' / 'Base salidas'."); st.stop()

# ENTRADAS
    e = pd.read_excel(io.BytesIO(bts), sheet_name=pe, header=4)
    e.columns = [str(c).strip() for c in e.columns]
    
    # 🛡️ FILTRO DE SEGURIDAD: Normaliza nombres por si varían los acentos en el Excel
    mapeo_columnas = {}
    for col in e.columns:
        norm = col.strip().upper().replace("Á","A").replace("É","E").replace("Í","I").replace("Ó","O").replace("Ú","U")
        if norm == "BIOLOGICO": mapeo_columnas[col] = "BIOLOGICO"
        elif norm == "FECHA DE RECEPCION": mapeo_columnas[col] = "FECHA DE RECEPCIÓN"
        elif norm == "FECHA DE CADUCIDAD": mapeo_columnas[col] = "FECHA DE CADUCIDAD"
    if mapeo_columnas:
        e = e.rename(columns=mapeo_columnas)

    e = e.dropna(subset=["BIOLOGICO"]).copy()
    e["BIOLOGICO"]          = e["BIOLOGICO"].astype(str).str.strip()
    e["FECHA DE RECEPCIÓN"] = pd.to_datetime(e["FECHA DE RECEPCIÓN"], errors="coerce")
    e["FECHA DE CADUCIDAD"] = pd.to_datetime(e["FECHA DE CADUCIDAD"],  errors="coerce")
    e["NO. DE LOTE."]       = e["NO. DE LOTE."].astype(str).str.strip()
    e["NUMERO DE DOSIS"]    = pd.to_numeric(e.get("NUMERO DE DOSIS",  0), errors="coerce").fillna(0)
    e["NUMERO DE DOSIS2"]   = pd.to_numeric(e.get("NUMERO DE DOSIS2", 0), errors="coerce").fillna(0)
    e["NUMERO DE FRASCOS"]  = pd.to_numeric(e.get("NUMERO DE FRASCOS",0), errors="coerce").fillna(0)
    e["TEMP. °C"]           = e.get("TEMP. °C", pd.Series(dtype=str)).fillna("").astype(str)
    e["PROCEDENCIA"]        = e.get("PROCEDENCIA", pd.Series(dtype=str)).fillna("").astype(str).str.strip()
    e["PRESENTAQCION"]      = e.get("PRESENTAQCION", e.get("PRESENTACION",
                              pd.Series(dtype=str))).fillna("").astype(str).str.strip()

    # SALIDAS (formato largo)
    s = pd.read_excel(io.BytesIO(bts), sheet_name=ps, header=0)
    s.columns = [str(c).strip() for c in s.columns]
    s = s.dropna(subset=["VACUNA"]).copy()
    s["FECHA DE SALIDA"] = pd.to_datetime(s["FECHA DE SALIDA"], errors="coerce")
    s["CANTIDAD"]        = pd.to_numeric(s["CANTIDAD"], errors="coerce").fillna(0)
    s["VACUNA"]          = s["VACUNA"].astype(str).str.strip()
    s["DESTINO"]         = s.get("DESTINO",     pd.Series(dtype=str)).fillna("").astype(str).str.strip()
    s["RESPONSABLE"]     = s.get("RESPONSABLE", pd.Series(dtype=str)).fillna("").astype(str).str.strip()

    años = sorted(set(e["FECHA DE RECEPCIÓN"].dt.year.dropna().astype(int)) |
                  set(s["FECHA DE SALIDA"].dt.year.dropna().astype(int)))
    biologicos = sorted(e["BIOLOGICO"].dropna().unique().tolist())

    return {"e": e, "s": s, "años": años, "biologicos": biologicos}

# ═══════════════════════════════════════════════════════
#  ARRASTRE
# ═══════════════════════════════════════════════════════
def _bio_en_salidas(bio: str, vacunas_s) -> str:
    if bio in vacunas_s: return bio
    mapped = MAPA_BIO_E_S.get(bio)
    if mapped and mapped in vacunas_s: return mapped
    candidatos = [v for v in vacunas_s if bio.upper() in v.upper()]
    return candidatos[0] if candidatos else bio

def bitacora_bio(datos, bio, f_ini, f_fin):
    e = datos["e"]; s = datos["s"]
    vacunas_s = s["VACUNA"].unique()
    bio_s = _bio_en_salidas(bio, vacunas_s)
    ts_ini = pd.Timestamp(f_ini); ts_fin = pd.Timestamp(f_fin)

    e_ant = e[(e["BIOLOGICO"]==bio) & (e["FECHA DE RECEPCIÓN"]<ts_ini)]
    s_ant = s[(s["VACUNA"]==bio_s) & (s["FECHA DE SALIDA"]<ts_ini)]
    saldo_hist = max(0, int(e_ant["NUMERO DE DOSIS2"].sum()) - int(s_ant["CANTIDAD"].sum()))

    ent = e[(e["BIOLOGICO"]==bio) &
            (e["FECHA DE RECEPCIÓN"]>=ts_ini) &
            (e["FECHA DE RECEPCIÓN"]<=ts_fin)].sort_values("FECHA DE RECEPCIÓN")
    sal = s[(s["VACUNA"]==bio_s) &
            (s["FECHA DE SALIDA"]>=ts_ini) &
            (s["FECHA DE SALIDA"]<=ts_fin)].sort_values("FECHA DE SALIDA")

    ev = ([("E", r["FECHA DE RECEPCIÓN"], r) for _, r in ent.iterrows()] +
          [("S", r["FECHA DE SALIDA"],    r) for _, r in sal.iterrows()])
    ev.sort(key=lambda x: x[1] if pd.notna(x[1]) else pd.Timestamp("2099-01-01"))

    filas = [{"tipo":"A","saldo":saldo_hist,
              "fecha_e":None,"procedencia":"","presentacion":"",
              "dosis_u":0,"temp_e":"","lote_e":"","cad_e":None,
              "frascos_e":0,"dosis_e":0,
              "fecha_s":None,"destino":"◀ Saldo inicial del período",
              "cantidad":0,"resp":""}]
    saldo = saldo_hist
    for tipo, fecha, r in ev:
        if tipo == "E":
            d = int(r["NUMERO DE DOSIS2"]); saldo += d
            filas.append({"tipo":"E","fecha_e":r["FECHA DE RECEPCIÓN"],
                "procedencia":r["PROCEDENCIA"],"presentacion":r["PRESENTAQCION"],
                "dosis_u":int(r["NUMERO DE DOSIS"]),"temp_e":r["TEMP. °C"],
                "lote_e":r["NO. DE LOTE."],"cad_e":r["FECHA DE CADUCIDAD"],
                "frascos_e":int(r["NUMERO DE FRASCOS"]),"dosis_e":d,"saldo":saldo,
                "fecha_s":None,"destino":"","cantidad":0,"resp":""})
        else:
            d = int(r["CANTIDAD"]); saldo = max(0, saldo-d)
            filas.append({"tipo":"S","saldo":saldo,
                "fecha_e":None,"procedencia":"","presentacion":"",
                "dosis_u":0,"temp_e":"","lote_e":"","cad_e":None,
                "frascos_e":0,"dosis_e":0,
                "fecha_s":r["FECHA DE SALIDA"],"destino":r["DESTINO"],
                "cantidad":d,"resp":r["RESPONSABLE"]})
    return filas

def _fmt(v):
    if v is None: return ""
    try:
        ts = pd.Timestamp(v)
        return ts.strftime("%d/%m/%Y") if pd.notna(ts) else ""
    except: return str(v)

# ═══════════════════════════════════════════════════════
#  FILTRO DE TIEMPO REUTILIZABLE
# ═══════════════════════════════════════════════════════
def filtro_tiempo(años_disp, key_prefix=""):
    c1, c2, c3 = st.columns([1, 1, 1])
    with c1:
        año_sel = st.selectbox("📅 Año",
            ["Todos"] + [str(a) for a in reversed(años_disp)],
            key=f"{key_prefix}_año")
    with c2:
        meses = ["Todos"] + [f"{m:02d} - {MESES_LARGO[m]}" for m in range(1,13)]
        mes_ini = st.selectbox("Mes inicio", meses, key=f"{key_prefix}_mesi")
    with c3:
        mes_fin = st.selectbox("Mes fin",    meses, key=f"{key_prefix}_mesf",
                               index=len(meses)-1)

    if año_sel == "Todos":
        return date(2000,1,1), date(2099,12,31), "Todos los años"
    año = int(año_sel)
    mi = int(mes_ini[:2]) if mes_ini != "Todos" else 1
    mf = int(mes_fin[:2]) if mes_fin != "Todos" else 12
    import calendar
    ultimo_dia = calendar.monthrange(año, mf)[1]
    etiqueta = (f"{año}" if mes_ini=="Todos" else
                f"{MESES_LARGO[mi]}–{MESES_LARGO[mf]} {año}")
    return date(año, mi, 1), date(año, mf, ultimo_dia), etiqueta

# ═══════════════════════════════════════════════════════
#  EXPORTAR EXCEL GENÉRICO
# ═══════════════════════════════════════════════════════
def to_excel(df: pd.DataFrame, hoja="Reporte") -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine=XL_ENGINE) as writer:
        df.to_excel(writer, index=False, sheet_name=hoja[:31])
        if XL_ENGINE == "xlsxwriter":
            wb = writer.book; ws = writer.sheets[hoja[:31]]
            hf = wb.add_format({"bold":True,"bg_color":"#1a5276",
                                 "font_color":"white","border":1,"align":"center"})
            for i, col in enumerate(df.columns):
                ws.write(0, i, col, hf)
                ws.set_column(i, i, max(14, len(str(col))+2))
    return buf.getvalue()

def to_excel_multi(sheets: dict) -> bytes:
    """Genera Excel con múltiples hojas {nombre: dataframe}"""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine=XL_ENGINE) as writer:
        for nombre, df in sheets.items():
            if df is None or df.empty: continue
            df.to_excel(writer, index=True, sheet_name=str(nombre)[:31])
            if XL_ENGINE == "xlsxwriter":
                wb = writer.book; ws = writer.sheets[str(nombre)[:31]]
                hf = wb.add_format({"bold":True,"bg_color":"#1a5276",
                                    "font_color":"white","border":1,"align":"center"})
                ws.write(0, 0, "Biológico / Destino", hf)
                for i, col in enumerate(df.columns):
                    ws.write(0, i+1, str(col), hf)
                    ws.set_column(i+1, i+1, 14)
                ws.set_column(0, 0, 22)
    return buf.getvalue()

# ═══════════════════════════════════════════════════════
#  PDF GENÉRICO
# ═══════════════════════════════════════════════════════
def df_to_pdf(df: pd.DataFrame, titulo: str, subtitulo: str = "") -> bytes:
    css = """
@page{size:A4 landscape;margin:8mm 6mm 8mm 6mm;
  @bottom-right{content:"Pág. " counter(page) "/" counter(pages);font-size:6pt}}
*{box-sizing:border-box;font-family:Arial,sans-serif;color:#1a252f}
h1{font-size:10pt;font-weight:700;color:#1a5276;text-align:center;margin:0 0 2px}
h2{font-size:7.5pt;color:#555;text-align:center;margin:0 0 5px;font-weight:400}
table{width:100%;border-collapse:collapse;font-size:6.5pt}
th{background:#1a5276;color:#fff;padding:3px 4px;border:0.4px solid #888;text-align:center}
td{padding:3px 4px;border:0.4px solid #ccc;vertical-align:middle}
.zb{background:#f4f6f7}.r{text-align:right}.c{text-align:center}
"""
    max_cols = 20
    df_p = df.reset_index() if df.index.name else df.copy()
    cols = list(df_p.columns)[:max_cols]
    ths = "".join(f"<th>{c}</th>" for c in cols)
    rows = ""
    for i, (_, row) in enumerate(df_p[cols].iterrows()):
        cls = ' class="zb"' if i%2 else ""
        tds = "".join(
            f'<td class="r">{str(v)}</td>' if isinstance(v,(int,float)) else f"<td>{v}</td>"
            for v in row
        )
        rows += f"<tr{cls}>{tds}</tr>"

    html = (f"<!DOCTYPE html><html><head><meta charset='UTF-8'>"
            f"<style>{css}</style></head><body>"
            f"<h1>{titulo}</h1>"
            f"<h2>{subtitulo}</h2>"
            f"<table><thead><tr>{ths}</tr></thead><tbody>{rows}</tbody></table>"
            f"</body></html>")

    if WEASYPRINT_OK:
        buf = io.BytesIO(); WpHTML(string=html).write_pdf(buf); return buf.getvalue()
    return html.encode("utf-8")

# ═══════════════════════════════════════════════════════
#  SIDEBAR
# ═══════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("## 💉 Bitácora Biológicos")
    st.markdown("**Jurisdicción 06 – Xonacatlán**")
    st.divider()
    archivo = st.file_uploader("📂 Cargar libro Excel (.xlsx)", type=["xlsx"])
    st.divider()
    modulo = st.radio("Módulo:", [
        "📋 Bitácora General",
        "📄 Anexo A",
        "📊 Dashboard",
        "📦 Reporte por Biológico",
        "🏭 Inventario por Lote",
        "🖨️ Reporte PDF",
    ])

st.markdown("""
<div class="hdr">
  <h1>💉 Bitácora de Control de Biológicos V6</h1>
  <p>Jurisdicción / Distrito 06 Xonacatlán · CESAEM · Programa de Vacunación Universal</p>
</div>""", unsafe_allow_html=True)

if archivo is None:
    st.info("👈 Carga el archivo **Bitacora_de_Biologuicos.xlsx** en el panel izquierdo.")
    st.stop()

with st.spinner("Procesando..."):
    datos = cargar(archivo.read())

df_e      = datos["e"]
df_s      = datos["s"]
años_disp = datos["años"]
biologicos= datos["biologicos"]

# ══════════════════════════════════════════════════════
#  MÓDULO: BITÁCORA GENERAL
# ══════════════════════════════════════════════════════
if "Bitácora" in modulo:
    st.subheader("📋 Bitácora General – Doble entrada con arrastre")

    ca, cb = st.columns([3,2])
    with ca:
        bios_sel = st.multiselect("🔬 Biológico",biologicos,
                                   default=biologicos[:1] if biologicos else [])
    with cb:
        f_ini, f_fin, etiq = filtro_tiempo(años_disp, "bit")

    if not bios_sel:
        st.warning("Selecciona al menos un biológico."); st.stop()

    filas_por_bio = {bio: bitacora_bio(datos,bio,f_ini,f_fin) for bio in bios_sel}

    st.caption("🟣 Violeta=saldo anterior │ 🔵 Azul=entrada │ ⬜ Blanco=salida │ 🟣 SALDO=remanente")

    tabs = st.tabs([f"💉 {b}" for b in bios_sel])
    for tab, bio in zip(tabs, bios_sel):
        with tab:
            filas = filas_por_bio[bio]
            n_e = sum(1 for f in filas if f["tipo"]=="E")
            n_s = sum(1 for f in filas if f["tipo"]=="S")
            s_ini = filas[0]["saldo"]; s_fin = filas[-1]["saldo"]
            d_ent = sum(f["dosis_e"]  for f in filas if f["tipo"]=="E")
            d_sal = sum(f["cantidad"] for f in filas if f["tipo"]=="S")
            m1,m2,m3,m4,m5 = st.columns(5)
            m1.metric("🟣 Saldo inicial", f"{s_ini:,}")
            m2.metric("📥 Entradas", n_e, f"+{d_ent:,}")
            m3.metric("📤 Salidas",  n_s, f"-{d_sal:,}")
            m4.metric("⚖️ Saldo final", f"{s_fin:,}")
            m5.metric("📅 Eventos", n_e+n_s)

            COLS=[("F.Recepción","c"),("Procedencia","l"),("Present.","l"),
                  ("Dos/U","r"),("°C","c"),("Lote","c"),("Caducidad","c"),
                  ("Frascos","r"),("Dosis Rec.","r"),("⚖ SALDO","r"),
                  ("F.Salida","c"),("Destino","l"),("Cantidad","r"),("Responsable","l")]
            TH="padding:5px 6px;border:1px solid #aaa;font-size:11px;white-space:nowrap;color:white"
            ths="".join(f'<th style="{TH};background:{"#7d3c98" if i==9 else "#1a5276"}">{h}</th>'
                        for i,(h,_) in enumerate(COLS))
            rows=""; mes_a=None
            for f in filas:
                fe=f["fecha_e"]; fs=f["fecha_s"]
                ref=fe if fe and pd.notna(fe) else fs
                if ref and pd.notna(ref) and f["tipo"]!="A":
                    ts=pd.Timestamp(ref); mk=(ts.year,ts.month)
                    if mk!=mes_a:
                        mes_a=mk
                        rows+=(f'<tr><td colspan="14" style="background:#ecf0f1;font-size:11px;'
                               f'font-weight:600;color:#444;padding:4px 8px;border-top:2px solid #bdc3c7">'
                               f'▸ {MESES_LARGO.get(mk[1],"")} {mk[0]}</td></tr>')
                bg=("#f0e6ff" if f["tipo"]=="A" else "#eaf4fb" if f["tipo"]=="E" else "#fff")
                cells=[_fmt(f["fecha_e"]),f["procedencia"],f["presentacion"],
                       str(int(f["dosis_u"])) if f["dosis_u"] else "",
                       f["temp_e"],f["lote_e"],_fmt(f["cad_e"]),
                       str(int(f["frascos_e"])) if f["frascos_e"] else "",
                       str(int(f["dosis_e"])) if f["dosis_e"] else "",
                       f'{int(f["saldo"]):,}',
                       _fmt(f["fecha_s"]),f["destino"],
                       str(int(f["cantidad"])) if f["cantidad"] else "",f["resp"]]
                tds=""
                for j,(cell,(_,al)) in enumerate(zip(cells,COLS)):
                    al2={"r":"right","c":"center","l":"left"}.get(al,"left")
                    ex=("background:#7d3c98;color:white;font-weight:700;" if j==9
                        else f"background:{bg};")
                    tds+=(f'<td style="{ex}text-align:{al2};padding:4px 6px;'
                          f'border:1px solid #ddd;font-size:11px;white-space:nowrap">{cell}</td>')
                rows+=f"<tr>{tds}</tr>"
            st.markdown(
                f'<div style="overflow-x:auto"><table style="border-collapse:collapse;'
                f'min-width:900px;width:100%"><thead><tr>{ths}</tr></thead>'
                f'<tbody>{rows}</tbody></table></div>', unsafe_allow_html=True)

    st.divider()
    dc1,dc2=st.columns(2)
    with dc1:
        recs=[]
        for bio,filas in filas_por_bio.items():
            for f in filas:
                recs.append({"Biológico":bio,
                    "Tipo":"SALDO ANT." if f["tipo"]=="A" else "ENTRADA" if f["tipo"]=="E" else "SALIDA",
                    "F.Recepción":_fmt(f["fecha_e"]),"Lote":f["lote_e"],
                    "Dosis Rec.":f["dosis_e"] or "","SALDO":int(f["saldo"]),
                    "F.Salida":_fmt(f["fecha_s"]),"Destino":f["destino"],
                    "Cantidad":f["cantidad"] or "","Responsable":f["resp"]})
        st.download_button("⬇️ Descargar Excel", to_excel(pd.DataFrame(recs),"Bitácora"),
            f"bitacora_{etiq}.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ══════════════════════════════════════════════════════
#  MÓDULO: ANEXO A
# ══════════════════════════════════════════════════════
elif "Anexo" in modulo:
    st.subheader("📄 Anexo A – Tarjeta de Control de Biológicos")

    # Encabezado institucional
    st.markdown("""
    <div style="border:2px solid #1a5276;border-radius:8px;padding:12px 16px;
                background:#f8fbff;margin-bottom:16px;font-size:12px">
      <div style="text-align:center;font-weight:700;font-size:14px;color:#1a5276;margin-bottom:6px">
        CENTRO NACIONAL PARA LA SALUD DE LA INFANCIA Y LA ADOLESCENCIA<br>
        PROGRAMA DE VACUNACIÓN UNIVERSAL · TARJETA DE CONTROL DE BIOLÓGICOS
      </div>
      <table style="width:100%;font-size:12px;border-collapse:collapse">
        <tr>
          <td><b>Entidad Federativa:</b> Estado de México</td>
          <td><b>Jurisdicción / Distrito:</b> 06 Xonacatlán</td>
          <td><b>Localidad / Municipio:</b> Huixquilucan</td>
        </tr>
        <tr>
          <td><b>Unidad de Salud:</b> San Fernando</td>
          <td><b>Responsable:</b> MTO. N. Alberto Reyes Jiménez</td>
          <td><b>Unidad Refrigerante:</b> 1</td>
        </tr>
      </table>
    </div>""", unsafe_allow_html=True)

    ca, cb = st.columns([2,2])
    with ca:
        bio_a = st.selectbox("🔬 Biológico", biologicos)
    with cb:
        f_ini_a, f_fin_a, etiq_a = filtro_tiempo(años_disp, "anx")

    # Construir tabla Anexo A desde Base entradas
    e_filtrado = df_e[
        (df_e["BIOLOGICO"] == bio_a) &
        (df_e["FECHA DE RECEPCIÓN"] >= pd.Timestamp(f_ini_a)) &
        (df_e["FECHA DE RECEPCIÓN"] <= pd.Timestamp(f_fin_a))
    ].copy().reset_index(drop=True)

    # Calcular saldo acumulado por fila (arrastre desde origen)
    vacunas_s = df_s["VACUNA"].unique()
    bio_s = _bio_en_salidas(bio_a, vacunas_s)

    e_antes_total = df_e[(df_e["BIOLOGICO"]==bio_a) &
                          (df_e["FECHA DE RECEPCIÓN"]<pd.Timestamp(f_ini_a))]
    s_antes_total = df_s[(df_s["VACUNA"]==bio_s) &
                          (df_s["FECHA DE SALIDA"]<pd.Timestamp(f_ini_a))]
    saldo_inicio  = max(0, int(e_antes_total["NUMERO DE DOSIS2"].sum()) -
                           int(s_antes_total["CANTIDAD"].sum()))

    # Total salidas en el período (para columna "Salidas")
    s_periodo = df_s[
        (df_s["VACUNA"]==bio_s) &
        (df_s["FECHA DE SALIDA"]>=pd.Timestamp(f_ini_a)) &
        (df_s["FECHA DE SALIDA"]<=pd.Timestamp(f_fin_a))
    ]
    total_sal_periodo = int(s_periodo["CANTIDAD"].sum())

    # Tabla fila por fila
    saldo = saldo_inicio
    filas_anexo = []
    for n, (_, r) in enumerate(e_filtrado.iterrows(), 1):
        d = int(r["NUMERO DE DOSIS2"])
        saldo += d
        filas_anexo.append({
            "N.":            n,
            "Procedencia":   r["PROCEDENCIA"],
            "Destino":       "Coord. San Fernando",
            "F. Recepción":  _fmt(r["FECHA DE RECEPCIÓN"]),
            "Temp. °C":      r["TEMP. °C"],
            "Biológico":     r["BIOLOGICO"],
            "Lote":          r["NO. DE LOTE."],
            "Caducidad":     _fmt(r["FECHA DE CADUCIDAD"]),
            "Entradas":      d,
            "Salidas":       "",
            "Saldo":         saldo,
            "Observaciones": "",
        })

    if not filas_anexo:
        st.info(f"Sin entradas de {bio_a} en el período seleccionado.")
    else:
        # KPIs rápidos
        total_ent = sum(f["Entradas"] for f in filas_anexo)
        k1,k2,k3,k4 = st.columns(4)
        k1.metric("📥 Total entradas",     f"{total_ent:,} dosis")
        k2.metric("📤 Total salidas per.", f"{total_sal_periodo:,} dosis")
        k3.metric("🟣 Saldo al inicio",    f"{saldo_inicio:,}")
        k4.metric("⚖️ Saldo al cierre",    f"{saldo:,}")

        df_anexo = pd.DataFrame(filas_anexo)

        # Render tabla Anexo A
        cols_ax = list(df_anexo.columns)
        TH="padding:5px 7px;border:1px solid #7f8c8d;font-size:11px;color:white;background:#1a5276"
        ths="".join(f'<th style="{TH}">{c}</th>' for c in cols_ax)
        rows=""
        for i, row in df_anexo.iterrows():
            bg="#eaf4fb" if i%2==0 else "#fff"
            tds=""
            for j,col in enumerate(cols_ax):
                v=str(row[col])
                al="right" if col in ["Entradas","Salidas","Saldo","N."] else "center" if col in ["F. Recepción","Caducidad","Temp. °C","Lote"] else "left"
                saldo_style="background:#7d3c98;color:white;font-weight:700;" if col=="Saldo" else f"background:{bg};"
                tds+=(f'<td style="{saldo_style}text-align:{al};padding:4px 7px;'
                      f'border:1px solid #ddd;font-size:11px;white-space:nowrap">{v}</td>')
            rows+=f"<tr>{tds}</tr>"

        st.markdown(
            f'<div class="anexo-hdr">TARJETA DE CONTROL · {bio_a} · {etiq_a}</div>'
            f'<div style="overflow-x:auto"><table style="border-collapse:collapse;'
            f'width:100%;min-width:900px"><thead><tr>{ths}</tr></thead>'
            f'<tbody>{rows}</tbody></table></div>', unsafe_allow_html=True)

        st.divider()
        d1,d2=st.columns(2)
        with d1:
            st.download_button("⬇️ Excel Anexo A",
                to_excel(df_anexo, f"Anexo A {bio_a}"),
                f"anexo_a_{bio_a}_{etiq_a}.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with d2:
            pdf=df_to_pdf(df_anexo,
                f"ANEXO A – {bio_a}",
                f"Jurisdicción 06 Xonacatlán · {etiq_a} · "
                f"Responsable: MTO. N. Alberto Reyes Jiménez")
            ext="pdf" if WEASYPRINT_OK else "html"
            st.download_button(f"⬇️ PDF Anexo A",pdf,
                f"anexo_a_{bio_a}_{etiq_a}.{ext}",
                "application/pdf" if WEASYPRINT_OK else "text/html")

# ══════════════════════════════════════════════════════
#  MÓDULO: DASHBOARD
# ══════════════════════════════════════════════════════
elif "Dashboard" in modulo:
    st.subheader("📊 Dashboard – Segmentador de Tiempo")

    f_ini_d, f_fin_d, etiq_d = filtro_tiempo(años_disp, "dash")
    ts_i=pd.Timestamp(f_ini_d); ts_f=pd.Timestamp(f_fin_d)

    e_d = df_e[(df_e["FECHA DE RECEPCIÓN"]>=ts_i)&(df_e["FECHA DE RECEPCIÓN"]<=ts_f)]
    s_d = df_s[(df_s["FECHA DE SALIDA"]>=ts_i)   &(df_s["FECHA DE SALIDA"]<=ts_f)]

    total_rec  = int(e_d["NUMERO DE DOSIS2"].sum())
    total_dist = int(s_d["CANTIDAD"].sum())
    n_bio_rec  = e_d["BIOLOGICO"].nunique()
    n_dest     = s_d["DESTINO"].nunique()
    efic = round(total_dist/total_rec*100,1) if total_rec else 0

    st.markdown(f"**Período: {etiq_d}**")
    k1,k2,k3,k4,k5 = st.columns(5)
    k1.metric("💉 Dosis Recibidas",    f"{total_rec:,}")
    k2.metric("✅ Dosis Distribuidas", f"{total_dist:,}")
    k3.metric("📈 Eficiencia",          f"{efic}%")
    k4.metric("🔬 Biológicos",         n_bio_rec)
    k5.metric("🏥 Destinos activos",   n_dest)

    st.divider()
    tab1, tab2, tab3, tab4 = st.tabs(["📥 Entradas","📤 Salidas por biológico",
                                       "🏥 Por destino","📅 Tendencia mensual"])

    with tab1:
        por_bio_e = (e_d.groupby("BIOLOGICO")["NUMERO DE DOSIS2"]
                       .sum().sort_values(ascending=False).reset_index())
        por_bio_e.columns = ["Biológico","Dosis Recibidas"]
        st.dataframe(por_bio_e, use_container_width=True, hide_index=True)
        st.bar_chart(por_bio_e.set_index("Biológico"))

    with tab2:
        por_bio_s = (s_d.groupby("VACUNA")["CANTIDAD"]
                       .sum().sort_values(ascending=False).reset_index())
        por_bio_s.columns = ["Vacuna","Dosis Distribuidas"]
        c1,c2=st.columns(2)
        c1.dataframe(por_bio_s, use_container_width=True, hide_index=True)
        c2.bar_chart(por_bio_s.set_index("Vacuna"))

    with tab3:
        por_dest = (s_d.groupby("DESTINO")["CANTIDAD"]
                      .sum().sort_values(ascending=False).reset_index())
        por_dest.columns=["Destino","Dosis"]
        st.bar_chart(por_dest.set_index("Destino"))
        st.dataframe(por_dest, use_container_width=True, hide_index=True)

    with tab4:
        s_d2 = s_d.copy()
        s_d2["Mes"] = s_d2["FECHA DE SALIDA"].dt.to_period("M").astype(str)
        tend = s_d2.groupby("Mes")["CANTIDAD"].sum().reset_index()
        tend.columns=["Mes","Dosis"]
        st.line_chart(tend.set_index("Mes"))
        e_d2=e_d.copy()
        e_d2["Mes"]=e_d2["FECHA DE RECEPCIÓN"].dt.to_period("M").astype(str)
        tend_e=e_d2.groupby("Mes")["NUMERO DE DOSIS2"].sum().reset_index()
        tend_e.columns=["Mes","Dosis Recibidas"]
        st.caption("Recibidas por mes")
        st.bar_chart(tend_e.set_index("Mes"))

    st.divider()
    st.download_button("⬇️ Descargar resumen Excel",
        to_excel(por_bio_s,"Dist x Biológico"),
        f"dashboard_{etiq_d}.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ══════════════════════════════════════════════════════
#  MÓDULO: REPORTE POR BIOLÓGICO × ELEMENTO
# ══════════════════════════════════════════════════════
elif "Reporte" in modulo:
    st.subheader("📦 Reporte de Salidas – Por Biológico × Elemento")
    st.caption("Tabla dinámica: filas=Biológico, columnas=Destino o Responsable, valores=Dosis")

    ca, cb, cc = st.columns([2,2,1])
    with ca:
        f_ini_r, f_fin_r, etiq_r = filtro_tiempo(años_disp, "rep")
    with cb:
        bios_r = st.multiselect("🔬 Filtrar biológico",
            ["Todos"]+sorted(df_s["VACUNA"].unique().tolist()),
            default=["Todos"], key="rep_bio")
    with cc:
        eje = st.radio("Columnas por:", ["Destino","Responsable"], key="rep_eje")

    ts_ir=pd.Timestamp(f_ini_r); ts_fr=pd.Timestamp(f_fin_r)
    s_r = df_s[(df_s["FECHA DE SALIDA"]>=ts_ir)&(df_s["FECHA DE SALIDA"]<=ts_fr)].copy()
    if "Todos" not in bios_r and bios_r:
        s_r = s_r[s_r["VACUNA"].isin(bios_r)]

    if s_r.empty:
        st.info("Sin datos para el período y filtros seleccionados.")
    else:
        col_eje = "DESTINO" if eje=="Destino" else "RESPONSABLE"
        pivot = (s_r.groupby(["VACUNA", col_eje])["CANTIDAD"]
                    .sum()
                    .unstack(fill_value=0))
        pivot["TOTAL"] = pivot.sum(axis=1)
        pivot = pivot.sort_values("TOTAL", ascending=False)

        # KPIs
        k1,k2,k3 = st.columns(3)
        k1.metric("💉 Total dosis distribuidas", f"{int(pivot['TOTAL'].sum()):,}")
        k2.metric(f"🏥 {eje}s únicos",           pivot.shape[1]-1)
        k3.metric("🔬 Biológicos",                len(pivot))

        st.markdown(f"**Período: {etiq_r}  |  Columnas por: {eje}**")

        # Render tabla pivot con colores de calor
        def render_pivot(df_piv):
            max_val = df_piv.drop(columns=["TOTAL"],errors="ignore").max().max()
            if max_val == 0: max_val = 1
            cols_p = list(df_piv.columns)
            TH="padding:5px 6px;border:1px solid #888;font-size:10.5px;color:white"
            ths=(f'<th style="{TH};background:#1a5276;min-width:120px">Biológico</th>' +
                 "".join(
                    f'<th style="{TH};background:{"#d35400" if c=="TOTAL" else "#2c3e50"};'
                    f'white-space:nowrap">{c}</th>' for c in cols_p))
            rows=""
            for i,(idx,row) in enumerate(df_piv.iterrows()):
                bg="#f4f6f7" if i%2 else "#fff"
                tds=f'<td style="background:{bg};padding:4px 6px;border:1px solid #ddd;font-size:11px;font-weight:600">{idx}</td>'
                for col in cols_p:
                    val=int(row[col])
                    if col=="TOTAL":
                        cell_bg="#d35400"; txt_c="white"; fw="700"
                    elif val==0:
                        cell_bg="#f9f9f9"; txt_c="#ccc"; fw="400"
                    else:
                        intensity=min(val/max_val,1)
                        r_v=int(234-intensity*150); g_v=int(244-intensity*100); b_v=255
                        cell_bg=f"rgb({r_v},{g_v},{b_v})"
                        txt_c="#1a252f" if intensity<0.7 else "white"; fw="500"
                    tds+=(f'<td style="background:{cell_bg};color:{txt_c};font-weight:{fw};'
                          f'text-align:right;padding:4px 6px;border:1px solid #ddd;'
                          f'font-size:11px;white-space:nowrap">'
                          f'{"" if val==0 else f"{val:,}"}</td>')
                rows+=f"<tr>{tds}</tr>"
            return (f'<div style="overflow-x:auto">'
                    f'<table style="border-collapse:collapse;width:100%">'
                    f'<thead><tr>{ths}</tr></thead><tbody>{rows}</tbody></table></div>')

        st.markdown(render_pivot(pivot), unsafe_allow_html=True)

        st.divider()
        # Gráfica de barras agrupadas por top destinos
        top_cols = [c for c in pivot.columns if c!="TOTAL"][:10]
        if top_cols:
            st.markdown(f"**Top {len(top_cols)} {eje}s – distribución por biológico**")
            st.bar_chart(pivot[top_cols])

        # Tendencia mensual por biológico
        st.markdown("**Tendencia mensual de distribución**")
        s_r2=s_r.copy()
        s_r2["Mes"]=s_r2["FECHA DE SALIDA"].dt.to_period("M").astype(str)
        tend_bio=s_r2.groupby(["Mes","VACUNA"])["CANTIDAD"].sum().unstack(fill_value=0)
        st.line_chart(tend_bio)

        # Descargas
        st.divider()
        d1,d2=st.columns(2)
        with d1:
            pivot_dl=pivot.reset_index().rename(columns={"VACUNA":"Biológico"})
            st.download_button("⬇️ Descargar Excel (tabla dinámica)",
                to_excel(pivot_dl,"Reporte"),
                f"reporte_{eje}_{etiq_r}.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with d2:
            pdf=df_to_pdf(pivot.reset_index().rename(columns={"VACUNA":"Biológico"}),
                f"Reporte de Distribución – Por {eje}",
                f"Jurisdicción 06 Xonacatlán · {etiq_r}")
            ext="pdf" if WEASYPRINT_OK else "html"
            st.download_button(f"⬇️ {ext.upper()}", pdf,
                f"reporte_{eje}_{etiq_r}.{ext}",
                "application/pdf" if WEASYPRINT_OK else "text/html")

# ══════════════════════════════════════════════════════
#  MÓDULO: INVENTARIO POR LOTE
# ══════════════════════════════════════════════════════
elif "Inventario" in modulo:
    st.subheader("🏭 Inventario en Tiempo Real – Saldo por Lote")
    hoy=pd.Timestamp(date.today())
    filas_inv=[]
    for bio in biologicos:
        fb=bitacora_bio(datos,bio,date(2000,1,1),date(2099,12,31))
        s_fin=fb[-1]["saldo"] if fb else 0
        for f in [x for x in fb if x["tipo"]=="E"]:
            cad=f["cad_e"]
            diff=(pd.Timestamp(cad)-hoy).days if cad and pd.notna(cad) else None
            alerta=("VENCIDO" if diff is not None and diff<0 else
                    "RIESGO"  if diff is not None and diff<=90 else
                    "VIGENTE" if diff is not None else "SIN FECHA")
            filas_inv.append({"Biológico":bio,"Lote":f["lote_e"],
                "Caducidad":_fmt(f["cad_e"]),"Dosis Entr.":f["dosis_e"],
                "Saldo Global":s_fin,"Estado":alerta})
    inv_df=pd.DataFrame(filas_inv)
    def _c(v):
        if v=="VENCIDO": return "background-color:#fadbd8;color:#c0392b;font-weight:700"
        if v=="RIESGO":  return "background-color:#fef9e7;color:#d35400;font-weight:700"
        if v=="VIGENTE": return "background-color:#eafaf1;color:#1e8449;font-weight:700"
        return ""
    a1,a2,a3=st.columns(3)
    a1.metric("🔴 Vencidos", int((inv_df["Estado"]=="VENCIDO").sum()))
    a2.metric("🟡 Riesgo",   int((inv_df["Estado"]=="RIESGO").sum()))
    a3.metric("🟢 Vigentes", int((inv_df["Estado"]=="VIGENTE").sum()))
    bf=st.multiselect("Filtrar",["Todos"]+biologicos,default=["Todos"],key="inv_f")
    df_f=inv_df if "Todos" in bf else inv_df[inv_df["Biológico"].isin(bf)]
    st.dataframe(df_f.style.applymap(_c,subset=["Estado"]),
                 use_container_width=True,hide_index=True)
    st.download_button("⬇️ Excel",to_excel(df_f.reset_index(drop=True),"Inventario"),
        "inventario.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ══════════════════════════════════════════════════════
#  MÓDULO: PDF GENERAL
# ══════════════════════════════════════════════════════
elif "PDF" in modulo:
    st.subheader("🖨️ Reporte PDF Institucional – Bitácora por Biológico")
    if not WEASYPRINT_OK:
        st.error("WeasyPrint no disponible. Instala GTK3 Runtime + pip install weasyprint")
    ca,cb=st.columns([3,2])
    with ca:
        bp=st.multiselect("Biológicos",biologicos,default=biologicos[:3])
    with cb:
        f_ini_p,f_fin_p,etiq_p=filtro_tiempo(años_disp,"pdf")
    if st.button("🖨️ Generar",type="primary"):
        s_p=df_s[(df_s["FECHA DE SALIDA"]>=pd.Timestamp(f_ini_p))&
                  (df_s["FECHA DE SALIDA"]<=pd.Timestamp(f_fin_p))]
        col_e="DESTINO"; pivot_p=(s_p.groupby(["VACUNA","DESTINO"])["CANTIDAD"]
                                     .sum().unstack(fill_value=0))
        pivot_p["TOTAL"]=pivot_p.sum(axis=1)
        pdf=df_to_pdf(pivot_p.reset_index().rename(columns={"VACUNA":"Biológico"}),
            "Reporte de Distribución de Biológicos",
            f"Jurisdicción 06 Xonacatlán · {etiq_p} · "
            f"Responsable: MTO. N. Alberto Reyes Jiménez")
        ext="pdf" if WEASYPRINT_OK else "html"
        st.download_button(f"⬇️ {ext.upper()}",pdf,
            f"reporte_{etiq_p}.{ext}",
            "application/pdf" if WEASYPRINT_OK else "text/html")

st.divider()
st.markdown(f"<small>Bitácora V6 · Jurisdicción 06 Xonacatlán · "
            f"{date.today().strftime('%d/%m/%Y')}</small>",unsafe_allow_html=True)
