import subprocess, sys
for pkg in ["openpyxl", "pandas", "streamlit"]:
    try: __import__(pkg)
    except ImportError: subprocess.check_call([sys.executable, "-m", "pip", "install", pkg, "-q"])

import streamlit as st
import pandas as pd
import base64

st.set_page_config(page_title="Itaú · Cupos Disponibles", page_icon="🏦", layout="wide")

# Minimal CSS — only what Streamlit reliably applies
st.markdown("""
<link href="https://fonts.googleapis.com/css2?family=Nunito:wght@400;600;700;800;900&display=swap" rel="stylesheet">
<style>
html, body, [class*="css"] { font-family: 'Nunito', sans-serif !important; }
#MainMenu, footer, header { visibility: hidden; }
.block-container { padding-top: 1rem !important; }
div.stButton > button {
    background-color: #FF6200 !important;
    color: white !important;
    font-weight: 700 !important;
    border: none !important;
    border-radius: 8px !important;
    width: 100% !important;
}
div.stButton > button:hover { background-color: #cc4f00 !important; }
</style>
""", unsafe_allow_html=True)

ORANGE = "#FF6200"

# ── Data ──────────────────────────────────────────────────────────────────────
@st.cache_data
def load_data():
    def load_sheet(sheet):
        df = pd.read_excel("Consumo_Bancos_Internacionales_02062026.xlsx", sheet_name=sheet)
        df = df.dropna(subset=["ID Solcred"]).copy()
        df["ID Solcred"] = df["ID Solcred"].astype(int)
        df["País"] = df["País"].astype(str)
        for col in ["Disponible Trade & Garantias","Disponible Gestão de Caixa",
                    "Disponible Derivativos","Disponible Bonds","Total Disponible"]:
            if col not in df.columns: df[col] = 0.0
            else: df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
        if "RATING" not in df.columns: df["RATING"] = ""
        return df

    df_col = load_sheet("Consumo Bcos Colombia")
    df_pan = load_sheet("Consumo Bcos Panamá")

    df_r = pd.read_excel("Riesgo_País_20260602.xlsx", header=None)
    df_r = df_r.iloc[2:, :8].copy()
    df_r.columns = ["País","Línea","LímiteTotal","ConsumoTotal","DisponibleTotal","L2","C2","D2"]
    df_r = df_r[df_r["País"].notna()].copy()
    df_r["País"] = df_r["País"].astype(str)
    for col in ["LímiteTotal","ConsumoTotal","DisponibleTotal"]:
        df_r[col] = pd.to_numeric(df_r[col], errors="coerce").fillna(0) * 1_000_000
    return df_col, df_pan, df_r

df_col, df_pan, df_r = load_data()

SHEETS = {"🇨🇴 Colombia": df_col, "🇵🇦 Panamá": df_pan}
PRODUCT_MAP_COL = {
    "Trade & Garantías": ("Disponible Trade & Garantias", "4. Trade & Garantias"),
    "Gestão de Caixa":   ("Disponible Gestão de Caixa",   "2. Gestão de Caixa"),
    "Derivativos":       ("Disponible Derivativos",        "3. Derivativos"),
    "Bonds":             ("Disponible Bonds",              "5. Bonds | Aplicações LP"),
}
PRODUCT_MAP_PAN = {
    "Trade & Garantías": ("Disponible Trade & Garantias", "4. Trade & Garantias"),
    "Gestão de Caixa":   ("Disponible Gestão de Caixa",   "2. Gestão de Caixa"),
    "Bonds":             ("Disponible Bonds",              "5. Bonds | Aplicações LP"),
}

def fmt(n):
    if pd.isna(n) or n is None: return "—"
    n = float(n)
    if abs(n) >= 1e6:  return f"{n/1e6:,.2f} MM USD"
    if abs(n) >= 1e3:  return f"{n/1e3:,.1f} K USD"
    return f"{n:,.2f} USD"

def alertas_for(df_rp):
    out = []
    for _, row in df_rp.iterrows():
        lim, con = row["LímiteTotal"], row["ConsumoTotal"]
        if lim and lim > 0:
            pct = (con / lim) * 100
            if pct >= 80:
                out.append({"linea": row["Línea"], "pct": pct,
                            "limite": lim, "consumo": con, "disponible": row["DisponibleTotal"]})
    return sorted(out, key=lambda x: -x["pct"])

def all_alertas():
    out = []
    for pais in df_r["País"].unique():
        for a in alertas_for(df_r[df_r["País"] == pais]):
            a["pais"] = pais; out.append(a)
    return sorted(out, key=lambda x: -x["pct"])

def get_logo_b64():
    with open("itau_logo.png", "rb") as f:
        return base64.b64encode(f.read()).decode()

def render_header(title):
    logo = get_logo_b64()
    st.markdown(f"""
<div style="background:#FF6200;padding:18px 32px;display:flex;align-items:center;justify-content:space-between;margin:-1rem -1rem 1.5rem -1rem;border-radius:0 0 12px 12px">
  <img src="data:image/png;base64,{logo}" style="height:90px;border-radius:8px">
  <div style="text-align:right">
    <div style="color:#fff;font-size:20px;font-weight:800;font-family:Nunito,sans-serif">{title}</div>
    <div style="color:rgba(255,255,255,0.85);font-size:11px;margin-top:4px;font-family:Nunito,sans-serif">Valores en USD · Corte 02 Jun 2026</div>
  </div>
</div>
""", unsafe_allow_html=True)

def render_alerta(a, show_pais=False):
    ic = a["pct"] >= 95
    bg     = "#fef2f2" if ic else "#fffbeb"
    border = "#dc2626" if ic else "#f59e0b"
    txt    = "#dc2626" if ic else "#b45309"
    badge  = "🔴 CRÍTICO" if ic else "🟡 ALERTA"
    pw     = min(100, a["pct"])
    bar_bg = "#fee2e2" if ic else "#fef3c7"
    pais_html = f'<span style="background:#e5e7eb;color:#374151;border-radius:5px;padding:1px 8px;font-size:11px;font-weight:700;margin-left:8px">País {a["pais"]}</span>' if show_pais else ""
    st.markdown(f"""
<div style="background:{bg};border-left:6px solid {border};border-radius:10px;padding:16px 20px;margin-bottom:12px">
  <div style="display:flex;justify-content:space-between;align-items:flex-start">
    <div>
      <div style="font-size:12px;font-weight:800;color:{txt};margin-bottom:3px">{badge}{pais_html}</div>
      <div style="font-size:15px;font-weight:800;color:#222">{a['linea']}</div>
    </div>
    <div style="text-align:right">
      <div style="font-size:28px;font-weight:900;color:{txt}">{a['pct']:.1f}%</div>
      <div style="font-size:10px;color:#888">del límite consumido</div>
    </div>
  </div>
  <div style="margin-top:10px;height:8px;background:{bar_bg};border-radius:4px;overflow:hidden">
    <div style="height:8px;background:{border};width:{pw:.1f}%;border-radius:4px"></div>
  </div>
  <div style="display:flex;gap:20px;margin-top:8px;font-size:12px;color:#555;flex-wrap:wrap">
    <span>Límite: <b>{fmt(a['limite'])}</b></span>
    <span>Consumo: <b style="color:{txt}">{fmt(a['consumo'])}</b></span>
    <span>Disponible: <b>{fmt(a['disponible'])}</b></span>
  </div>
</div>
""", unsafe_allow_html=True)


# ── Session ───────────────────────────────────────────────────────────────────
if "page"    not in st.session_state: st.session_state.page    = "home"
if "country" not in st.session_state: st.session_state.country = "🇨🇴 Colombia"


# ══════════════════════════════════════════════════════════════════════════════
# HOME
# ══════════════════════════════════════════════════════════════════════════════
if st.session_state.page == "home":
    render_header("Dashboard de Cupos Disponibles")

    alertas  = all_alertas()
    criticos = [a for a in alertas if a["pct"] >= 95]
    normales = [a for a in alertas if 80 <= a["pct"] < 95]
    paises_af = len(set(a["pais"] for a in alertas))

    # Country toggle
    st.markdown("**Seleccionar base de consumo**")
    country_opts = list(SHEETS.keys())
    sel_c = st.radio("", country_opts, horizontal=True,
                     index=country_opts.index(st.session_state.country),
                     label_visibility="collapsed")
    if sel_c != st.session_state.country:
        st.session_state.country = sel_c; st.rerun()

    df_active = SHEETS[st.session_state.country]
    flag  = "🇨🇴" if "Colombia" in st.session_state.country else "🇵🇦"
    label = "Colombia" if "Colombia" in st.session_state.country else "Panamá"

    st.divider()

    # Nav cards
    col1, col2 = st.columns(2, gap="large")

    with col1:
        st.markdown(f"""
<div style="background:#fff;border:2px solid #FF6200;border-radius:16px;padding:32px 24px;text-align:center;box-shadow:0 2px 12px rgba(255,98,0,0.1)">
  <div style="font-size:48px;margin-bottom:12px">🔍</div>
  <div style="font-size:20px;font-weight:800;color:#222;margin-bottom:8px">Búsqueda por ID</div>
  <div style="font-size:13px;color:#888;line-height:1.6">Consulta cupos disponibles y riesgo país por ID Solcred</div>
  <div style="margin-top:12px;display:inline-block;background:#fff3ec;border-radius:20px;padding:4px 14px;font-size:13px;font-weight:700;color:#FF6200">{flag} Consumo Bcos {label}</div>
  <div style="margin-top:8px;font-size:13px;font-weight:700;color:#FF6200">{len(df_active)} registros disponibles →</div>
</div>
""", unsafe_allow_html=True)
        st.markdown("")
        if st.button("Ir a Búsqueda →", key="go_busqueda"):
            st.session_state.page = "busqueda"; st.rerun()

    with col2:
        tri_c  = "#dc2626" if criticos else "#f59e0b"
        alert_bg     = "#fef2f2" if criticos else "#fffbeb"
        alert_border = "#dc2626" if criticos else "#f59e0b"
        st.markdown(f"""
<div style="background:{alert_bg};border:2px solid {alert_border};border-radius:16px;padding:32px 24px;text-align:center;box-shadow:0 2px 12px rgba(0,0,0,0.08)">
  <div style="display:flex;justify-content:center;margin-bottom:12px">
    <svg width="52" height="46" viewBox="0 0 52 46"><path d="M26 2L50 44H2Z" fill="{tri_c}"/><text x="26" y="34" text-anchor="middle" font-size="19" font-weight="bold" fill="white">!</text></svg>
  </div>
  <div style="font-size:20px;font-weight:800;color:#222;margin-bottom:8px">Alertas de Riesgo País</div>
  <div style="font-size:13px;color:#888;line-height:1.6">Líneas con consumo ≥ 80% del límite asignado por país</div>
  <div style="margin-top:14px;display:flex;justify-content:center;gap:8px;flex-wrap:wrap">
    <span style="background:#dc2626;color:#fff;border-radius:20px;padding:4px 12px;font-size:12px;font-weight:700">🔴 {len(criticos)} críticas</span>
    <span style="background:#f59e0b;color:#fff;border-radius:20px;padding:4px 12px;font-size:12px;font-weight:700">🟡 {len(normales)} alertas</span>
    <span style="background:#6b7280;color:#fff;border-radius:20px;padding:4px 12px;font-size:12px;font-weight:700">{paises_af} países</span>
  </div>
</div>
""", unsafe_allow_html=True)
        st.markdown("")
        if st.button("Ver Alertas →", key="go_alertas"):
            st.session_state.page = "alertas"; st.rerun()

    st.divider()

    # Resumen global
    st.markdown("**RESUMEN GLOBAL**")
    paises_n = len(df_active["País"].unique())
    g1, g2, g3, g4 = st.columns(4)
    for gcol, lbl, val, color in [
        (g1, "Entidades",       len(df_active),  "#222"),
        (g2, "Países",          paises_n,        "#222"),
        (g3, "Alertas activas", len(alertas),    "#f59e0b" if alertas else "#222"),
        (g4, "Casos críticos",  len(criticos),   "#dc2626" if criticos else "#222"),
    ]:
        gcol.markdown(f"""
<div style="background:#fff;border-radius:12px;padding:20px;text-align:center;box-shadow:0 1px 6px rgba(0,0,0,0.07)">
  <div style="font-size:34px;font-weight:900;color:{color}">{val}</div>
  <div style="font-size:12px;color:#888;margin-top:4px">{lbl}</div>
</div>
""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(f"""
<div style="background:#FF6200;padding:44px 32px;text-align:center;border-radius:12px;margin-top:16px">
  <p style="color:#fff;font-size:20px;font-weight:700;font-style:italic;letter-spacing:3px;text-transform:uppercase;margin:0">Cada resultado cuenta una historia</p>
</div>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# BÚSQUEDA
# ══════════════════════════════════════════════════════════════════════════════
elif st.session_state.page == "busqueda":
    render_header(f"Búsqueda — {st.session_state.country}")

    bc = st.columns([1, 9])
    with bc[0]:
        if st.button("‹ Inicio"):
            st.session_state.page = "home"; st.rerun()

    # Country switcher
    country_opts = list(SHEETS.keys())
    sc2 = st.radio("Base de consumo:", country_opts, horizontal=True,
                   index=country_opts.index(st.session_state.country))
    if sc2 != st.session_state.country:
        st.session_state.country = sc2; st.rerun()

    df_active = SHEETS[st.session_state.country]
    prod_map  = PRODUCT_MAP_COL if "Colombia" in st.session_state.country else PRODUCT_MAP_PAN
    all_ids   = sorted(df_active["ID Solcred"].unique().tolist())

    sel = st.selectbox("Selecciona o busca un ID Solcred", [None] + all_ids,
                       format_func=lambda x: "— Selecciona un ID —" if x is None else f"ID {x}")

    if sel:
        row    = df_active[df_active["ID Solcred"] == sel].iloc[0]
        pais   = str(row["País"])
        df_rp  = df_r[df_r["País"] == pais]
        rating = str(row["RATING"]) if pd.notna(row["RATING"]) else "N/A"
        total  = float(row["Total Disponible"]) if pd.notna(row["Total Disponible"]) else 0
        flag   = "🇨🇴" if "Colombia" in st.session_state.country else "🇵🇦"

        # ID header
        st.markdown(f"""
<div style="background:#FF6200;border-radius:12px;padding:18px 24px;margin:12px 0;display:flex;align-items:center;gap:24px;flex-wrap:wrap;box-shadow:0 4px 14px rgba(255,98,0,0.2)">
  <div><div style="color:rgba(255,255,255,0.8);font-size:10px;text-transform:uppercase;letter-spacing:1.5px">ID Solcred</div>
       <div style="color:#fff;font-size:28px;font-weight:900">{sel}</div></div>
  <div style="width:1px;height:44px;background:rgba(255,255,255,0.3)"></div>
  <div><div style="color:rgba(255,255,255,0.8);font-size:10px;text-transform:uppercase;letter-spacing:1.5px">País</div>
       <div style="color:#fff;font-size:22px;font-weight:800">{pais}</div></div>
  <div style="width:1px;height:44px;background:rgba(255,255,255,0.3)"></div>
  <div><div style="color:rgba(255,255,255,0.8);font-size:10px;text-transform:uppercase;letter-spacing:1.5px">Rating</div>
       <div style="background:#fff;color:#FF6200;border-radius:7px;padding:3px 12px;font-size:17px;font-weight:900;display:inline-block">{rating}</div></div>
  <div style="width:1px;height:44px;background:rgba(255,255,255,0.3)"></div>
  <div><div style="color:rgba(255,255,255,0.8);font-size:10px;text-transform:uppercase;letter-spacing:1.5px">Base</div>
       <div style="color:#fff;font-size:17px;font-weight:800">{flag} {st.session_state.country.split()[1]}</div></div>
  <div style="margin-left:auto;text-align:right">
    <div style="color:rgba(255,255,255,0.8);font-size:10px;text-transform:uppercase;letter-spacing:1.5px">Total Disponible</div>
    <div style="color:#fff;font-size:22px;font-weight:900">{fmt(total)}</div>
  </div>
</div>
""", unsafe_allow_html=True)

        # Cupos
        st.markdown("**CUPOS DISPONIBLES POR TIPO DE TRANSACCIÓN**")
        st.caption("Consumo Bancos Internacionales (USD) · Disponible Riesgo País por línea (USD)")

        cols = st.columns(len(prod_map))
        for i, (tipo, (col_c, linea_r)) in enumerate(prod_map.items()):
            val_c   = float(row[col_c]) if pd.notna(row[col_c]) else 0
            rp_l    = df_rp[df_rp["Línea"] == linea_r]
            lim     = float(rp_l["LímiteTotal"].values[0])     if len(rp_l) else None
            disp_r  = float(rp_l["DisponibleTotal"].values[0]) if len(rp_l) else None
            pct     = min(100, (val_c / lim) * 100) if (lim and lim > 0) else 0
            bc_col  = "#dc2626" if pct >= 95 else "#f59e0b" if pct >= 80 else "#FF6200"
            c_color = "#cc4f00" if val_c > 0 else "#999"
            r_color = "#1d4ed8" if (disp_r and disp_r > 0) else "#999"
            bar_html = ""
            if lim:
                bar_html = f"""
  <div style="height:6px;background:#f0f0f0;border-radius:4px;margin-top:10px;overflow:hidden">
    <div style="height:6px;background:{bc_col};width:{pct:.1f}%;border-radius:4px"></div>
  </div>
  <div style="font-size:10px;color:#aaa;margin-top:4px">Límite: {fmt(lim)} · <b style="color:{bc_col}">{pct:.1f}%</b> consumido</div>"""

            cols[i].markdown(f"""
<div style="background:#fff;border-radius:12px;padding:18px 16px 14px;border-top:5px solid #FF6200;box-shadow:0 2px 8px rgba(0,0,0,0.07);height:100%">
  <div style="font-size:11px;color:#777;font-weight:700;text-transform:uppercase;letter-spacing:1px;margin-bottom:10px">{tipo}</div>
  <div style="font-size:10px;color:#aaa;margin-bottom:2px">Disponible (Consumo)</div>
  <div style="font-size:19px;font-weight:800;color:{c_color}">{fmt(val_c)}</div>
  <div style="height:1px;background:#f0f0f0;margin:10px 0"></div>
  <div style="font-size:10px;color:#aaa;margin-bottom:2px">Disponible País (Riesgo)</div>
  <div style="font-size:19px;font-weight:800;color:{r_color}">{fmt(disp_r) if disp_r is not None else "—"}</div>
  {bar_html}
</div>
""", unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # Alertas del ID
        alts = alertas_for(df_rp)
        if alts:
            st.markdown(f"**⚠ {len(alts)} ALERTAS DE RIESGO PAÍS**")
            st.caption("Líneas con consumo ≥ 80% del límite asignado al país")
            for a in alts:
                render_alerta(a)
            st.markdown("<br>", unsafe_allow_html=True)

        # Tabla Riesgo País
        st.markdown(f"**RIESGO PAÍS — PAÍS {pais}**")
        st.caption("Fuente: Riesgo País · convertido a USD")

        rp_tot = df_rp[df_rp["Línea"] == "Total Transferência"]
        if not rp_tot.empty:
            rt = rp_tot.iloc[0]
            tc1, tc2, tc3 = st.columns(3)
            for tcol, lbl, val, bg, fc in [
                (tc1, "Límite Total",     rt["LímiteTotal"],    "#fff",    "#cc4f00"),
                (tc2, "Consumo Total",    rt["ConsumoTotal"],   "#fff",    "#cc4f00"),
                (tc3, "Disponible Total", rt["DisponibleTotal"],"#FF6200", "#fff"),
            ]:
                tcol.markdown(f"""
<div style="background:{bg};border-radius:12px;padding:16px 20px;box-shadow:0 1px 4px rgba(0,0,0,0.07);margin-bottom:12px">
  <div style="font-size:10px;color:{'rgba(255,255,255,0.8)' if bg=='#FF6200' else '#888'};text-transform:uppercase;letter-spacing:1.2px;margin-bottom:4px">{lbl}</div>
  <div style="font-size:22px;font-weight:800;color:{fc}">{fmt(val)}</div>
</div>
""", unsafe_allow_html=True)

        if not df_rp.empty:
            tabla = df_rp[["Línea","LímiteTotal","ConsumoTotal","DisponibleTotal"]].copy()
            tabla.columns = ["Línea","Límite (USD)","Consumo (USD)","Disponible (USD)"]
            for c in ["Límite (USD)","Consumo (USD)","Disponible (USD)"]:
                tabla[c] = tabla[c].apply(fmt)
            st.dataframe(tabla, use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════════════════════════════════════
# ALERTAS
# ══════════════════════════════════════════════════════════════════════════════
elif st.session_state.page == "alertas":
    render_header("Alertas de Riesgo País")

    bc = st.columns([1, 9])
    with bc[0]:
        if st.button("‹ Inicio"):
            st.session_state.page = "home"; st.rerun()

    alertas  = all_alertas()
    criticos = [a for a in alertas if a["pct"] >= 95]
    normales = [a for a in alertas if 80 <= a["pct"] < 95]

    filtro = st.radio("", ["Todas", "Críticas (≥95%)", "Alertas (80-94%)"], horizontal=True,
                      label_visibility="collapsed")
    lista = criticos if "Críticas" in filtro else normales if "Alertas" in filtro else alertas

    st.markdown(f"**{len(lista)} resultado(s)**")
    st.markdown("")

    if not lista:
        st.success("✅ Sin alertas en esta categoría")
    else:
        for a in lista:
            render_alerta(a, show_pais=True)
