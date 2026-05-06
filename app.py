"""
PCC Integrity — Procesamiento CIPS
Corriente Interrumpida: limpieza, LRS y sincronización con SharePoint
"""
import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import msal
import requests
import io
import os
import sys
import tempfile
import datetime
import json

st.set_page_config(
    page_title="PCC – Procesamiento CIPS",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
  html, body, [class*="css"] { font-family: 'Inter', sans-serif !important; }
  .stApp { background: #F8FAFC; }
  .block-container { padding: 2rem 2.5rem 1rem 2.5rem !important; max-width: 1200px; }
  [data-testid="stSidebar"] > div:first-child {
    background: linear-gradient(180deg, #b8233e 0%, #7a1429 100%) !important;
  }
  /* Texto general del sidebar en blanco */
  [data-testid="stSidebar"] * { color: rgba(255,255,255,0.92) !important; }

  /* Selectboxes: fondo blanco, texto oscuro para que se lea */
  [data-testid="stSidebar"] [data-baseweb="select"] > div {
    background: rgba(255,255,255,0.95) !important;
    border: 1px solid rgba(255,255,255,0.3) !important;
    border-radius: 8px !important;
  }
  [data-testid="stSidebar"] [data-baseweb="select"] * {
    color: #1E293B !important;
  }
  [data-testid="stSidebar"] [data-baseweb="select"] svg {
    fill: #1E293B !important;
  }
  /* Labels encima del selectbox: blancos */
  [data-testid="stSidebar"] p,
  [data-testid="stSidebar"] strong {
    color: rgba(255,255,255,0.92) !important;
  }
  .stButton > button {
    background: linear-gradient(135deg, #b8233e, #d42848) !important;
    color: white !important; border: none !important;
    border-radius: 10px !important; font-weight: 600 !important;
    transition: transform 0.2s ease, box-shadow 0.2s ease !important;
  }
  .stButton > button:hover { transform: translateY(-2px) !important; }
  [data-testid="stDownloadButton"] > button {
    background: linear-gradient(135deg, #1A7A4A, #22A06B) !important;
    color: white !important; border: none !important;
    border-radius: 10px !important; font-weight: 600 !important;
  }
  .bloque { background:white; border-radius:14px; padding:1.5rem;
    box-shadow:0 2px 12px rgba(0,0,0,0.06); margin-bottom:1rem; }
  .bloque-titulo { font-weight:700; font-size:0.85rem; color:#8B0000;
    text-transform:uppercase; letter-spacing:0.8px; margin-bottom:1rem;
    padding-bottom:0.5rem; border-bottom:2px solid #F0F2F6; }
  #MainMenu, footer { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# ── Rutas ─────────────────────────────────────────────────────────────────────
def _rp(rel):
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)

EXCEL_INFRA = _rp(os.path.join("data", "Listado de Infraestructura para Cod Informes.xlsx"))
SHAPEFILES  = _rp("shapefiles")

# ── Session state ──────────────────────────────────────────────────────────────
for _k, _v in {
    "res_df": None, "res_bytes": None, "res_name": None, "sp_url": None,
}.items():
    if _k not in st.session_state:
        st.session_state[_k] = _v


@st.cache_data(ttl=0)
def cargar_infra():
    return pd.read_excel(EXCEL_INFRA)


def get_shp(distrito, linea):
    try:
        df_lineas = cargar_infra()
        fila = df_lineas[(df_lineas["DISTRITO"] == distrito) & (df_lineas["TRAMO"] == linea)]
        if fila.empty:
            return None
        return os.path.join(SHAPEFILES, fila["ID TRAMO"].values[0] + ".shp")
    except Exception:
        return None


def get_sp_token():
    try:
        cfg = st.secrets.get("sharepoint", {})
        app_obj = msal.ConfidentialClientApplication(
            cfg["client_id"],
            authority=f"https://login.microsoftonline.com/{cfg['tenant_id']}",
            client_credential=cfg["client_secret"],
        )
        result = app_obj.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
        return result.get("access_token")
    except Exception:
        return None


@st.cache_data(ttl=600, show_spinner="Sincronizando archivos CIPS desde SharePoint...")
def fetch_cips_sharepoint_files():
    if "sharepoint" not in st.secrets:
        return []
    cfg = st.secrets["sharepoint"]
    folder = cfg.get("cips_folder_path", "")
    if not folder:
        return []
    try:
        app_obj = msal.ConfidentialClientApplication(
            cfg["client_id"],
            authority=f"https://login.microsoftonline.com/{cfg['tenant_id']}",
            client_credential=cfg["client_secret"],
        )
        token = app_obj.acquire_token_for_client(
            scopes=["https://graph.microsoft.com/.default"]
        ).get("access_token")
        if not token:
            return []
        headers  = {"Authorization": f"Bearer {token}"}
        hostname = f"{cfg['tenant_name']}.sharepoint.com"
        site_path = cfg["site_url"].replace(f"https://{hostname}", "")
        site_id   = requests.get(
            f"https://graph.microsoft.com/v1.0/sites/{hostname}:{site_path}", headers=headers
        ).json().get("id")
        items = requests.get(
            f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{folder}:/children",
            headers=headers
        ).json().get("value", [])
        out = []
        for item in items:
            name  = item.get("name", "")
            d_url = item.get("@microsoft.graph.downloadUrl")
            if not name.endswith(".xlsx") or name.startswith("~") or not d_url:
                continue
            r = requests.get(d_url)
            if not r.ok:
                continue
            f = io.BytesIO(r.content)
            f.name = name
            try:
                if "Survey Data" in pd.ExcelFile(f).sheet_names:
                    f.seek(0)
                    out.append(f)
            except Exception:
                pass
        return out
    except Exception as e:
        st.sidebar.warning(f"SharePoint: {e}")
        return []


# ── Sidebar ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div style="padding:1.2rem 0 0.5rem; text-align:center;">
      <div style="font-size:1.3rem;font-weight:800;letter-spacing:-0.5px;">PCC Integrity</div>
      <div style="font-size:0.75rem;opacity:0.7;margin-top:2px;">Procesamiento CIPS</div>
    </div>
    <hr style="border-color:rgba(255,255,255,0.2);margin:0.5rem 0;">
    """, unsafe_allow_html=True)

    try:
        df_lineas = cargar_infra()
        st.markdown("**Cliente**")
        cliente = st.selectbox("Cliente", ["TGI","OCENSA"], label_visibility="collapsed")
        if cliente == "OCENSA":
            df_c  = df_lineas[df_lineas["DISTRITO"] == "OCENSA"]
            dist  = "OCENSA"
            st.markdown("**Tramo**")
            linea = st.selectbox("Tramo", df_c["TRAMO"].tolist(), label_visibility="collapsed")
        else:
            df_c  = df_lineas[df_lineas["DISTRITO"] != "OCENSA"]
            dists = sorted(df_c["DISTRITO"].unique())
            st.markdown("**Distrito**")
            dist  = st.selectbox("Distrito", dists, label_visibility="collapsed")
            lineas = df_c[df_c["DISTRITO"] == dist]["TRAMO"].tolist()
            st.markdown("**Línea**")
            linea = st.selectbox("Línea", lineas, label_visibility="collapsed")
    except Exception:
        st.warning("No se encontró el archivo de infraestructura.")
        cliente, dist, linea = "TGI", "—", "—"

    st.markdown('<hr style="border-color:rgba(255,255,255,0.2);margin:0.8rem 0;">', unsafe_allow_html=True)
    sp_files = fetch_cips_sharepoint_files()
    if sp_files:
        st.markdown(f'<p style="font-size:0.8rem;font-weight:600;">SharePoint: {len(sp_files)} archivo(s)</p>', unsafe_allow_html=True)
        for f in sp_files:
            st.markdown(f'<div style="font-size:0.75rem;opacity:0.85;padding:2px 0;">{f.name}</div>', unsafe_allow_html=True)
        if st.button("Refrescar", use_container_width=True):
            fetch_cips_sharepoint_files.clear(); st.rerun()


# ── Main ───────────────────────────────────────────────────────────────────────
shp    = get_shp(dist, linea)
shp_ok = bool(shp and os.path.exists(shp))

st.markdown(f"""
<div style="background:linear-gradient(135deg,#b8233e,#7a1429);color:white;
            padding:1.5rem 2rem;border-radius:14px;margin-bottom:1.5rem;
            box-shadow:0 6px 24px rgba(123,30,58,0.35);">
  <h1 style="margin:0;font-size:1.6rem;font-weight:800;">Procesamiento CIPS</h1>
  <p style="margin:0.3rem 0 0;opacity:0.8;font-size:0.95rem;">
    Corriente Interrumpida · Limpieza de datos, cálculo LRS y sincronización SharePoint
  </p>
</div>
""", unsafe_allow_html=True)

col_up, col_param = st.columns([3, 2])

with col_up:
    st.markdown('<div class="bloque"><div class="bloque-titulo">Archivos de inspección</div>', unsafe_allow_html=True)
    if sp_files:
        st.markdown(f'<p style="font-size:0.82rem;color:#1B5E20;font-weight:600;">{len(sp_files)} archivo(s) de SharePoint</p>', unsafe_allow_html=True)
    archivos_subidos = st.file_uploader("Excel CIPS", type=["xlsx"], accept_multiple_files=True, label_visibility="collapsed")
    nombres_vistos = {f.name for f in sp_files}
    archivos_extra = []
    for a in (archivos_subidos or []):
        if a.name in nombres_vistos: continue
        nombres_vistos.add(a.name)
        try:
            if "Survey Data" in pd.ExcelFile(a).sheet_names:
                archivos_extra.append(a)
        except Exception:
            pass
    archivos = sp_files + archivos_extra
    if archivos_extra:
        st.success(f"{len(archivos_extra)} archivo(s) adicional(es) cargado(s)")
    st.markdown('</div>', unsafe_allow_html=True)

with col_param:
    shp_badge  = "✅ Encontrado" if shp_ok else "❌ No encontrado"
    shp_color  = "#166534" if shp_ok else "#991B1B"
    shp_bg     = "#F0FFF4" if shp_ok else "#FFF5F5"
    n_arch     = len(archivos)
    arch_color = "#166534" if n_arch > 0 else "#64748B"

    filas = []
    filas.append(("Cliente",    cliente,       "#0F172A"))
    if cliente == "TGI":
        filas.append(("Distrito", dist,         "#0F172A"))
    filas.append(("Tramo",      linea,          "#0F172A"))

    rows_html = "".join(f"""
    <div style="display:flex;justify-content:space-between;align-items:center;
                padding:6px 0;border-bottom:1px solid #F1F5F9;">
      <span style="font-size:0.8rem;color:#64748B;font-weight:600;">{k}</span>
      <span style="font-size:0.85rem;color:{c};font-weight:700;text-align:right;
                   max-width:60%;word-break:break-word;">{v}</span>
    </div>""" for k,v,c in filas)

    st.markdown(f"""
    <div style="background:white;border:1px solid #E2E8F0;border-radius:12px;
                padding:1.2rem;box-shadow:0 2px 8px rgba(0,0,0,0.04);">
      <p style="font-size:0.72rem;text-transform:uppercase;font-weight:700;
                color:#D50032;letter-spacing:0.08em;margin:0 0 0.8rem 0;">
        Parámetros seleccionados
      </p>
      {rows_html}
      <div style="display:flex;justify-content:space-between;align-items:center;
                  padding:6px 0;border-bottom:1px solid #F1F5F9;">
        <span style="font-size:0.8rem;color:#64748B;font-weight:600;">Shapefile</span>
        <span style="font-size:0.82rem;font-weight:700;color:{shp_color};
                     background:{shp_bg};padding:2px 8px;border-radius:20px;">
          {shp_badge}
        </span>
      </div>
      <div style="display:flex;justify-content:space-between;align-items:center;
                  padding:6px 0;">
        <span style="font-size:0.8rem;color:#64748B;font-weight:600;">Archivos</span>
        <span style="font-size:0.85rem;font-weight:700;color:{arch_color};">
          {n_arch} {'archivo' if n_arch == 1 else 'archivos'}
        </span>
      </div>
    </div>
    """, unsafe_allow_html=True)

_, col_btn, _ = st.columns([2, 1, 2])
with col_btn:
    procesar = st.button("Procesar inspección", use_container_width=True)

if procesar:
    if not archivos:
        st.error("Sube al menos un archivo Excel con hoja 'Survey Data'.")
    elif not shp_ok:
        st.error(f"No se encontró el shapefile para el tramo **{linea}**.")
    else:
        st.session_state.res_df = st.session_state.res_bytes = st.session_state.res_name = None
        st.session_state.sp_url = None
        prog = st.progress(0)
        estado = st.empty()

        def upd(p, msg):
            prog.progress(p, text=msg)
            estado.caption(msg)

        _ok = False
        with tempfile.TemporaryDirectory() as tmp:
            for a in archivos:
                a.seek(0)
                with open(os.path.join(tmp, a.name), "wb") as f:
                    f.write(a.read())
            try:
                upd(15, "Unificando archivos...")
                from mod_unificar import ejecutar_unificar
                unif = ejecutar_unificar(tmp)

                upd(55, "Calculando PK geométrico (LRS)...")
                from mod_cips_lrs import ejecutar_cips_lrs
                salida = ejecutar_cips_lrs(tmp, unif, shp)

                upd(85, "Cargando resultados...")
                df_res = pd.read_excel(salida, sheet_name="Survey Data")
                with open(salida, "rb") as f:
                    xbytes = f.read()

                ts     = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                nombre = f"CIPS_{linea.replace(' ','_')}_{ts}.xlsx"
                st.session_state.res_df    = df_res
                st.session_state.res_bytes = xbytes
                st.session_state.res_name  = nombre

                upd(92, "Subiendo a SharePoint...")
                try:
                    from mod_cips_sharepoint import subir_a_sharepoint
                    tok = get_sp_token()
                    if tok:
                        tmp_sp = os.path.join(tmp, nombre)
                        with open(tmp_sp, "wb") as f:
                            f.write(xbytes)
                        sub = f"ocensa/{linea.replace(' ','_')}" if cliente == "OCENSA" \
                              else datetime.datetime.now().strftime("%Y/%m")
                        url = subir_a_sharepoint(tmp_sp, tok, subcarpeta=sub)
                        st.session_state.sp_url = url
                except Exception as e_sp:
                    st.warning(f"Procesado OK, no se pudo subir a SharePoint: {e_sp}")

                upd(100, "¡Proceso completado!")
                _ok = True
            except Exception as e:
                import traceback
                prog.empty(); estado.empty()
                st.error(f"Error: {e}")
                with st.expander("Ver detalle"):
                    st.code(traceback.format_exc())

        if _ok:
            prog.empty(); estado.empty()
            st.rerun()

if st.session_state.res_df is not None:
    df = st.session_state.res_df
    st.markdown("---")
    total = len(df)
    prot  = int((df["Estado_CP"] == "PROTEGIDO").sum())   if "Estado_CP" in df.columns else 0
    desp  = int((df["Estado_CP"] == "DESPROTEGIDO").sum()) if "Estado_CP" in df.columns else 0
    sobre = int((df["Estado_CP"] == "SOBREPROTEGIDO").sum()) if "Estado_CP" in df.columns else 0
    c1,c2,c3,c4 = st.columns(4)
    c1.metric("Puntos totales", total)
    c2.metric("Protegidos",     prot)
    c3.metric("Desprotegidos",  desp)
    c4.metric("Sobreprotegidos",sobre)

    # ── Gráfica On/Off vs PK ──────────────────────────────────────────────────
    pk_col = next((c for c in ["PK_geom_m","PK_real_m"] if c in df.columns), None)
    if pk_col and ("On_mV_limpio" in df.columns or "Off_mV_limpio" in df.columns):
        sub = df.dropna(subset=[pk_col]).sort_values(pk_col)
        if len(sub) > 3000:
            sub = sub.iloc[::max(1, len(sub)//3000)]

        fig = go.Figure()
        if "On_mV_limpio" in sub.columns:
            fig.add_trace(go.Scatter(
                x=sub[pk_col], y=sub["On_mV_limpio"],
                mode="lines", name="On mV",
                line=dict(color="#9CA3AF", width=1.4),
                fill="tozeroy", fillcolor="rgba(156,163,175,0.05)"))
        if "Off_mV_limpio" in sub.columns:
            fig.add_trace(go.Scatter(
                x=sub[pk_col], y=sub["Off_mV_limpio"],
                mode="lines", name="Off mV",
                line=dict(color="#D50032", width=2.0)))

        fig.add_hrect(y0=-1200, y1=-850, fillcolor="rgba(55,65,81,0.05)", line_width=0,
                      annotation_text="Zona protegida", annotation_position="top left",
                      annotation_font=dict(size=9, color="#374151"))
        fig.add_hline(y=-850,  line=dict(color="#6B7280", dash="dash", width=1.2),
                      annotation_text="-850 mV", annotation_position="top right",
                      annotation_font=dict(size=9, color="#6B7280"))
        fig.add_hline(y=-1200, line=dict(color="#D50032", dash="dash", width=1.2),
                      annotation_text="-1.200 mV", annotation_position="bottom right",
                      annotation_font=dict(size=9, color="#D50032"))
        fig.update_layout(
            plot_bgcolor="white", paper_bgcolor="white",
            height=300, margin=dict(t=30,b=40,l=50,r=20),
            font=dict(size=12, family="Inter, sans-serif", color="#475569"),
            xaxis_title="PK (m)", yaxis_title="mV",
            legend=dict(orientation="h", y=-0.25, font_size=11),
            hovermode="x unified",
            xaxis=dict(showgrid=True, gridcolor="#F1F5F9", zeroline=False),
            yaxis=dict(showgrid=True, gridcolor="#F1F5F9", zeroline=False),
        )
        st.plotly_chart(fig, use_container_width=True)

    with st.expander("Vista previa de datos"):
        cols = [c for c in ["PK_geom_m","Lat_corr","Long_corr","On_mV_limpio","Off_mV_limpio","Estado_CP"] if c in df.columns]
        st.dataframe(df[cols].head(300), use_container_width=True, height=300)

    col_dl, col_sp = st.columns([1,2])
    with col_dl:
        st.download_button("Descargar Excel procesado",
                           data=st.session_state.res_bytes,
                           file_name=st.session_state.res_name,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)
    with col_sp:
        if st.session_state.sp_url:
            st.success(f"Subido: {st.session_state.sp_url}")
        else:
            if st.button("Subir a SharePoint", use_container_width=True):
                with st.spinner("Subiendo..."):
                    try:
                        from mod_cips_sharepoint import subir_a_sharepoint
                        tok = get_sp_token()
                        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tf:
                            tf.write(st.session_state.res_bytes); tf_path = tf.name
                        url = subir_a_sharepoint(tf_path, tok, subcarpeta=datetime.datetime.now().strftime("%Y/%m"))
                        os.unlink(tf_path)
                        st.session_state.sp_url = url
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error: {e}")
