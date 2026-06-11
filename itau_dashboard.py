import subprocess, sys
for pkg in ["openpyxl","pandas","streamlit"]:
    try: __import__(pkg)
    except ImportError: subprocess.check_call([sys.executable,"-m","pip","install",pkg,"-q"])

import streamlit as st
import pandas as pd
import base64

st.set_page_config(page_title="Itaú · Cupos Disponibles", page_icon="🏦", layout="wide")

st.markdown("""
<link href="https://fonts.googleapis.com/css2?family=Nunito:wght@400;600;700;800;900&display=swap" rel="stylesheet">
<style>
  html,body,[class*="css"],.stMarkdown,.stSelectbox,button,input,label,p,div {
    font-family:'Nunito',sans-serif !important;
  }
  #MainMenu,footer,header{visibility:hidden;}
  .block-container{padding:0!important;max-width:100%!important;}
  section[data-testid="stSidebar"]{display:none;}

  .itau-header{background:#FF6200;padding:20px 36px;display:flex;align-items:center;justify-content:space-between;}
  .header-right h2{color:#fff;margin:0;font-size:20px;font-weight:800;}
  .header-right p{color:rgba(255,255,255,0.85);margin:4px 0 0;font-size:11px;}
  .page-body{padding:28px 36px;}

  .nav-card{background:#fff;border-radius:20px;padding:38px 28px;text-align:center;box-shadow:0 2px 12px rgba(0,0,0,0.08);}
  .nav-card-alert{border-radius:20px;padding:38px 28px;text-align:center;box-shadow:0 2px 12px rgba(0,0,0,0.08);}
  .nav-title{font-size:22px;font-weight:800;color:#222;margin-bottom:10px;}
  .nav-desc{font-size:14px;color:#888;line-height:1.6;}
  .nav-count{margin-top:18px;font-size:13px;font-weight:700;color:#FF6200;}

  /* Country toggle */
  .country-toggle{display:flex;gap:0;border-radius:12px;overflow:hidden;border:2px solid #FF6200;width:fit-content;margin:0 auto 28px;}
  .ctab{padding:10px 28px;font-size:14px;font-weight:700;cursor:pointer;border:none;font-family:'Nunito',sans-serif!important;}
  .ctab-active{background:#FF6200;color:#fff;}
  .ctab-inactive{background:#fff;color:#FF6200;}

  .metric-box{background:#fff;border-radius:14px;padding:20px 22px;text-align:center;box-shadow:0 1px 6px rgba(0,0,0,0.07);}
  .metric-val{font-size:36px;font-weight:900;}
  .metric-lbl{font-size:12px;color:#888;margin-top:4px;}

  .id-header{background:#FF6200;border-radius:14px;padding:20px 28px;margin-bottom:22px;display:flex;gap:32px;align-items:center;flex-wrap:wrap;box-shadow:0 4px 14px rgba(255,98,0,0.25);}
  .id-field .lbl{color:rgba(255,255,255,0.8);font-size:10px;text-transform:uppercase;letter-spacing:1.5px;}
  .id-field .val{color:#fff;font-size:26px;font-weight:900;margin-top:2px;}
  .id-sep{width:1px;height:44px;background:rgba(255,255,255,0.3);}
  .rating-badge{background:#fff;color:#FF6200;border-radius:8px;padding:4px 14px;font-size:17px;font-weight:900;display:inline-block;}
  .total-right{margin-left:auto;text-align:right;}
  .total-right .lbl{color:rgba(255,255,255,0.8);font-size:10px;text-transform:uppercase;letter-spacing:1.5px;}
  .total-right .val{color:#fff;font-size:22px;font-weight:900;}

  .cupo-card{background:#fff;border-radius:14px;padding:20px 18px 16px;border-top:5px solid #FF6200;box-shadow:0 2px 8px rgba(0,0,0,0.07);}
  .cupo-tipo{font-size:11px;color:#777;font-weight:700;text-transform:uppercase;letter-spacing:1px;margin-bottom:10px;}
  .cupo-label{font-size:10px;color:#aaa;margin-bottom:3px;}
  .cupo-val{font-size:20px;font-weight:800;}
  .cupo-sep{height:1px;background:#f0f0f0;margin:12px 0;}
  .bar-wrap{height:7px;background:#f0f0f0;border-radius:4px;margin-top:10px;overflow:hidden;}
  .bar-fill{height:7px;border-radius:4px;}
  .bar-label{font-size:10px;color:#aaa;margin-top:4px;}

  .alerta-crit{background:#fef2f2;border-left:6px solid #dc2626;border-radius:12px;padding:16px 20px;margin-bottom:14px;}
  .alerta-warn{background:#fffbeb;border-left:6px solid #f59e0b;border-radius:12px;padding:16px 20px;margin-bottom:14px;}
  .alerta-head{display:flex;justify-content:space-between;align-items:flex-start;flex-wrap:wrap;gap:8px;}
  .alerta-badge-crit{font-size:12px;font-weight:800;color:#dc2626;margin-bottom:3px;}
  .alerta-badge-warn{font-size:12px;font-weight:800;color:#b45309;margin-bottom:3px;}
  .alerta-linea{font-size:16px;font-weight:800;color:#222;}
  .alerta-pais{font-size:11px;background:#e5e7eb;color:#374151;border-radius:6px;padding:2px 8px;font-weight:700;margin-left:8px;}
  .alerta-pct-crit{font-size:30px;font-weight:900;color:#dc2626;text-align:right;}
  .alerta-pct-warn{font-size:30px;font-weight:900;color:#b45309;text-align:right;}
  .alerta-sub{font-size:10px;color:#888;}
  .alerta-bar-wrap{margin-top:10px;height:9px;background:rgba(0,0,0,0.08);border-radius:4px;overflow:hidden;}
  .alerta-bar{height:9px;border-radius:4px;}
  .alerta-meta{display:flex;gap:24px;margin-top:10px;font-size:12px;color:#555;flex-wrap:wrap;}

  .sec-title{font-size:13px;font-weight:800;color:#222;text-transform:uppercase;letter-spacing:1.2px;margin-bottom:4px;}
  .sec-sub{font-size:11px;color:#aaa;margin-bottom:16px;}

  .total-card{border-radius:12px;padding:18px 22px;box-shadow:0 1px 4px rgba(0,0,0,0.07);margin-bottom:16px;}
  .total-card .lbl{font-size:10px;text-transform:uppercase;letter-spacing:1.2px;margin-bottom:5px;}
  .total-card .val{font-size:22px;font-weight:800;}

  .itau-footer{background:#FF6200;padding:48px 32px;text-align:center;margin-top:48px;}
  .itau-footer p{color:#fff;font-size:20px;font-weight:700;font-style:italic;letter-spacing:3px;text-transform:uppercase;margin:0;}

  div.stButton>button{background:#FF6200!important;color:#fff!important;font-weight:700!important;border:none!important;border-radius:10px!important;padding:10px 20px!important;font-size:14px!important;font-family:'Nunito',sans-serif!important;width:100%;}
  div.stButton>button:hover{background:#cc4f00!important;}
  .back-btn>button{background:transparent!important;color:#FF6200!important;border:2px solid #FF6200!important;width:auto!important;}

  .badge-crit{background:#dc2626;color:#fff;border-radius:20px;padding:4px 12px;font-size:12px;font-weight:700;margin-right:6px;}
  .badge-warn{background:#f59e0b;color:#fff;border-radius:20px;padding:4px 12px;font-size:12px;font-weight:700;margin-right:6px;}
  .badge-grey{background:#6b7280;color:#fff;border-radius:20px;padding:4px 12px;font-size:12px;font-weight:700;margin-right:6px;}

  /* Active tab highlight using radio */
  div[data-testid="stRadio"] label{font-weight:700;font-size:14px;}
  div[data-testid="stRadio"] > div{gap:0!important;}
  div[data-testid="stRadio"] > div > label{
    background:#fff;color:#FF6200;border:2px solid #FF6200;
    padding:8px 24px;font-family:'Nunito',sans-serif!important;
    font-weight:700;font-size:14px;border-radius:0;margin:0;
  }
  div[data-testid="stRadio"] > div > label:first-child{border-radius:10px 0 0 10px;}
  div[data-testid="stRadio"] > div > label:last-child{border-radius:0 10px 10px 0;}
  div[data-testid="stRadio"] > div > label[data-checked="true"],
  div[data-testid="stRadio"] > div > label:has(input:checked){
    background:#FF6200!important;color:#fff!important;
  }
</style>
""", unsafe_allow_html=True)


# ── Data ──────────────────────────────────────────────────────────────────────
COLS = ["ID Solcred","Entidad","País",
        "Disponible Trade & Garantias",
        "Disponible Gestão de Caixa",
        "Disponible Derivativos",
        "Disponible Bonds",
        "Total Disponible","RATING"]

@st.cache_data
def load_data():
    def load_sheet(sheet):
        df = pd.read_excel("Consumo_Bancos_Internacionales_02062026.xlsx", sheet_name=sheet)
        df = df.dropna(subset=["ID Solcred"]).copy()
        df["ID Solcred"] = df["ID Solcred"].astype(int)
        df["País"] = df["País"].astype(str)
        # Ensure all product columns exist (Panamá lacks Derivativos)
        for col in ["Disponible Trade & Garantias","Disponible Gestão de Caixa",
                    "Disponible Derivativos","Disponible Bonds","Total Disponible"]:
            if col not in df.columns:
                df[col] = 0.0
            else:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
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
    if abs(n)>=1e6:  return f"{n/1e6:,.2f} MM USD"
    if abs(n)>=1e3:  return f"{n/1e3:,.1f} K USD"
    return f"{n:,.2f} USD"

def get_logo():
    with open("itau_logo.png","rb") as f:
        return base64.b64encode(f.read()).decode()

def alertas_for(df_rp):
    out = []
    for _,row in df_rp.iterrows():
        lim,con = row["LímiteTotal"],row["ConsumoTotal"]
        if lim and lim>0:
            pct = (con/lim)*100
            if pct>=80:
                out.append({"linea":row["Línea"],"pct":pct,
                            "limite":lim,"consumo":con,"disponible":row["DisponibleTotal"]})
    return sorted(out, key=lambda x:-x["pct"])

def all_alertas():
    out=[]
    for pais in df_r["País"].unique():
        for a in alertas_for(df_r[df_r["País"]==pais]):
            a["pais"]=pais; out.append(a)
    return sorted(out, key=lambda x:-x["pct"])


# ── Session ───────────────────────────────────────────────────────────────────
if "page"    not in st.session_state: st.session_state.page    = "home"
if "country" not in st.session_state: st.session_state.country = "🇨🇴 Colombia"


# ── Header ────────────────────────────────────────────────────────────────────
titles = {"home":"Dashboard de Cupos Disponibles",
          "busqueda":f"Búsqueda — {st.session_state.country}",
          "alertas":"Alertas de Riesgo País"}
logo = get_logo()
st.markdown(f"""
<div class="itau-header">
  <img src="data:image/png;base64,{logo}" style="height:100px;border-radius:10px">
  <div class="header-right">
    <h2>{titles[st.session_state.page]}</h2>
    <p>Valores en USD &nbsp;·&nbsp; Corte 02 Jun 2026</p>
  </div>
</div>
""", unsafe_allow_html=True)

if st.session_state.page != "home":
    st.markdown('<div style="padding:16px 36px 0">', unsafe_allow_html=True)
    bc = st.columns([1,8])
    with bc[0]:
        st.markdown('<div class="back-btn">', unsafe_allow_html=True)
        if st.button("‹ Inicio"):
            st.session_state.page = "home"; st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

st.markdown('<div class="page-body">', unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# HOME
# ══════════════════════════════════════════════════════════════════════════════
if st.session_state.page == "home":

    alertas = all_alertas()
    criticos = [a for a in alertas if a["pct"]>=95]
    normales = [a for a in alertas if 80<=a["pct"]<95]
    paises_af = len(set(a["pais"] for a in alertas))

    # ── Country toggle ────────────────────────────────────────────────────────
    st.markdown('<div class="sec-title" style="text-align:center;margin-bottom:8px">Seleccionar base de consumo</div>', unsafe_allow_html=True)
    country_opts = list(SHEETS.keys())
    sel_country = st.radio("", country_opts, horizontal=True,
                           index=country_opts.index(st.session_state.country),
                           label_visibility="collapsed", key="country_radio")
    if sel_country != st.session_state.country:
        st.session_state.country = sel_country
        st.rerun()

    df_active = SHEETS[st.session_state.country]
    n_ids = len(df_active)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Nav cards ─────────────────────────────────────────────────────────────
    col1, col2 = st.columns(2, gap="large")

    with col1:
        flag = "🇨🇴" if "Colombia" in st.session_state.country else "🇵🇦"
        label = "Colombia" if "Colombia" in st.session_state.country else "Panamá"
        st.markdown(f"""
        <div class="nav-card" style="border:2px solid #FF6200">
          <div style="font-size:50px;margin-bottom:16px">🔍</div>
          <div class="nav-title">Búsqueda por ID</div>
          <div class="nav-desc">Consulta cupos disponibles y riesgo país por ID Solcred</div>
          <div style="margin-top:12px;display:inline-block;background:#fff3ec;border-radius:20px;padding:4px 14px;font-size:13px;font-weight:700;color:#FF6200">
            {flag} Consumo Bcos {label}
          </div>
          <div class="nav-count">{n_ids} registros disponibles →</div>
        </div><br>
        """, unsafe_allow_html=True)
        if st.button("Ir a Búsqueda →", key="go_busqueda"):
            st.session_state.page = "busqueda"; st.rerun()

    with col2:
        tri_c = "#dc2626" if criticos else "#f59e0b"
        tri = f'<svg width="54" height="48" viewBox="0 0 52 46"><path d="M26 2L50 44H2Z" fill="{tri_c}" stroke="{tri_c}" stroke-width="2"/><text x="26" y="34" text-anchor="middle" font-size="19" font-weight="bold" fill="white">!</text></svg>'
        alert_bg     = "#fef2f2" if criticos else "#fffbeb"
        alert_border = "#dc2626" if criticos else "#f59e0b"
        badges = f'<span class="badge-crit">🔴 {len(criticos)} críticas</span><span class="badge-warn">🟡 {len(normales)} alertas</span><span class="badge-grey">{paises_af} países</span>'
        st.markdown(f"""
        <div class="nav-card-alert" style="background:{alert_bg};border:2px solid {alert_border}">
          <div style="display:flex;justify-content:center;margin-bottom:16px">{tri}</div>
          <div class="nav-title">Alertas de Riesgo País</div>
          <div class="nav-desc">Líneas con consumo ≥ 80% del límite asignado por país</div>
          <div style="margin-top:18px">{badges}</div>
        </div><br>
        """, unsafe_allow_html=True)
        if st.button("Ver Alertas →", key="go_alertas"):
            st.session_state.page = "alertas"; st.rerun()

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Resumen global ────────────────────────────────────────────────────────
    st.markdown('<div class="sec-title">Resumen Global</div>', unsafe_allow_html=True)
    g1,g2,g3,g4 = st.columns(4)
    paises_n = len(df_active["País"].unique())
    for gcol,lbl,val,color in [
        (g1,"Entidades",       n_ids,        "#222"),
        (g2,"Países",          paises_n,     "#222"),
        (g3,"Alertas activas", len(alertas), "#f59e0b" if alertas else "#222"),
        (g4,"Casos críticos",  len(criticos),"#dc2626" if criticos else "#222"),
    ]:
        gcol.markdown(f"""
        <div class="metric-box">
          <div class="metric-val" style="color:{color}">{val}</div>
          <div class="metric-lbl">{lbl}</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("""
    </div>
    <div class="itau-footer"><p>Cada resultado cuenta una historia</p></div>
    """, unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# BÚSQUEDA
# ══════════════════════════════════════════════════════════════════════════════
elif st.session_state.page == "busqueda":

    df_active  = SHEETS[st.session_state.country]
    prod_map   = PRODUCT_MAP_COL if "Colombia" in st.session_state.country else PRODUCT_MAP_PAN
    all_ids    = sorted(df_active["ID Solcred"].unique().tolist())

    # Country switcher inside search
    country_opts = list(SHEETS.keys())
    sc2 = st.radio("Base de consumo:", country_opts, horizontal=True,
                   index=country_opts.index(st.session_state.country),
                   label_visibility="visible", key="country_busqueda")
    if sc2 != st.session_state.country:
        st.session_state.country = sc2; st.rerun()

    sel = st.selectbox("Selecciona o busca un ID Solcred",
                       [None]+all_ids,
                       format_func=lambda x:"— Selecciona un ID —" if x is None else f"ID {x}")

    if sel:
        row   = df_active[df_active["ID Solcred"]==sel].iloc[0]
        pais  = str(row["País"])
        df_rp = df_r[df_r["País"]==pais]
        rating= str(row["RATING"]) if pd.notna(row["RATING"]) else "N/A"
        total = float(row["Total Disponible"]) if pd.notna(row["Total Disponible"]) else 0

        flag = "🇨🇴" if "Colombia" in st.session_state.country else "🇵🇦"

        st.markdown(f"""
        <div class="id-header">
          <div class="id-field"><div class="lbl">ID Solcred</div><div class="val">{sel}</div></div>
          <div class="id-sep"></div>
          <div class="id-field"><div class="lbl">País</div><div class="val">{pais}</div></div>
          <div class="id-sep"></div>
          <div class="id-field"><div class="lbl">Rating</div><div class="rating-badge">{rating}</div></div>
          <div class="id-sep"></div>
          <div class="id-field"><div class="lbl">Base</div><div class="val" style="font-size:18px">{flag} {st.session_state.country.split()[1]}</div></div>
          <div class="total-right">
            <div class="lbl">Total Disponible</div>
            <div class="val">{fmt(total)}</div>
          </div>
        </div>
        """, unsafe_allow_html=True)

        # Cupos
        st.markdown('<div class="sec-title">Cupos Disponibles por Tipo de Transacción</div>', unsafe_allow_html=True)
        st.markdown('<div class="sec-sub">Consumo Bancos Internacionales (USD) · Disponible Riesgo País por línea (USD)</div>', unsafe_allow_html=True)

        cols = st.columns(len(prod_map))
        for i,(tipo,(col_c,linea_r)) in enumerate(prod_map.items()):
            val_c   = float(row[col_c]) if pd.notna(row[col_c]) else 0
            rp_l    = df_rp[df_rp["Línea"]==linea_r]
            lim     = float(rp_l["LímiteTotal"].values[0])     if len(rp_l) else None
            disp_r  = float(rp_l["DisponibleTotal"].values[0]) if len(rp_l) else None
            pct     = min(100,(val_c/lim)*100) if (lim and lim>0) else 0
            bc      = "#dc2626" if pct>=95 else "#f59e0b" if pct>=80 else "#FF6200"
            c_color = "#cc4f00" if val_c>0 else "#bbb"
            r_color = "#1d4ed8" if (disp_r and disp_r>0) else "#bbb"
            bar_html = ""
            if lim:
                bar_html = f"""
                <div class="bar-wrap"><div class="bar-fill" style="background:{bc};width:{pct:.1f}%"></div></div>
                <div class="bar-label">Límite: {fmt(lim)} · <span style="color:{bc};font-weight:700">{pct:.1f}% consumido</span></div>
                """
            cols[i].markdown(f"""
            <div class="cupo-card">
              <div class="cupo-tipo">{tipo}</div>
              <div class="cupo-label">Disponible (Consumo)</div>
              <div class="cupo-val" style="color:{c_color}">{fmt(val_c)}</div>
              <div class="cupo-sep"></div>
              <div class="cupo-label">Disponible País (Riesgo)</div>
              <div class="cupo-val" style="color:{r_color}">{fmt(disp_r) if disp_r is not None else "—"}</div>
              {bar_html}
            </div>
            """, unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # Alertas
        alts = alertas_for(df_rp)
        if alts:
            st.markdown(f'<div class="sec-title">⚠ {len(alts)} Alertas de Riesgo País</div>', unsafe_allow_html=True)
            st.markdown('<div class="sec-sub">Líneas con consumo ≥ 80% del límite asignado al país</div>', unsafe_allow_html=True)
            for a in alts:
                ic   = a["pct"]>=95
                css  = "alerta-crit" if ic else "alerta-warn"
                badge= "🔴 CRÍTICO"  if ic else "🟡 ALERTA"
                bcss = "alerta-badge-crit" if ic else "alerta-badge-warn"
                pcss = "alerta-pct-crit"   if ic else "alerta-pct-warn"
                bc2  = "#dc2626" if ic else "#f59e0b"
                txt  = "#dc2626" if ic else "#b45309"
                pw   = min(100,a["pct"])
                st.markdown(f"""
                <div class="{css}">
                  <div class="alerta-head">
                    <div><div class="{bcss}">{badge}</div><div class="alerta-linea">{a['linea']}</div></div>
                    <div><div class="{pcss}">{a['pct']:.1f}%</div><div class="alerta-sub">del límite consumido</div></div>
                  </div>
                  <div class="alerta-bar-wrap"><div class="alerta-bar" style="background:{bc2};width:{pw:.1f}%"></div></div>
                  <div class="alerta-meta">
                    <span>Límite: <strong>{fmt(a['limite'])}</strong></span>
                    <span>Consumo: <strong style="color:{txt}">{fmt(a['consumo'])}</strong></span>
                    <span>Disponible: <strong>{fmt(a['disponible'])}</strong></span>
                  </div>
                </div>
                """, unsafe_allow_html=True)
            st.markdown("<br>", unsafe_allow_html=True)

        # Tabla Riesgo País
        st.markdown(f'<div class="sec-title">Riesgo País — País {pais}</div>', unsafe_allow_html=True)
        st.markdown('<div class="sec-sub">Fuente: Riesgo País · convertido a USD</div>', unsafe_allow_html=True)
        rp_tot = df_rp[df_rp["Línea"]=="Total Transferência"]
        if not rp_tot.empty:
            rt = rp_tot.iloc[0]
            tc1,tc2,tc3 = st.columns(3)
            for tcol,lbl,val,bg,fc in [
                (tc1,"Límite Total",    rt["LímiteTotal"],    "#fff",    "#cc4f00"),
                (tc2,"Consumo Total",   rt["ConsumoTotal"],   "#fff",    "#cc4f00"),
                (tc3,"Disponible Total",rt["DisponibleTotal"],"#FF6200","#fff"),
            ]:
                tcol.markdown(f"""
                <div class="total-card" style="background:{bg}">
                  <div class="lbl" style="color:{'rgba(255,255,255,0.8)' if bg=='#FF6200' else '#888'}">{lbl}</div>
                  <div class="val" style="color:{fc}">{fmt(val)}</div>
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
    alertas  = all_alertas()
    criticos = [a for a in alertas if a["pct"]>=95]
    normales = [a for a in alertas if 80<=a["pct"]<95]

    filtro = st.radio("", ["Todas","Críticas (≥95%)","Alertas (80-94%)"], horizontal=True)
    lista  = criticos if "Críticas" in filtro else normales if "Alertas" in filtro else alertas

    st.markdown(f"<br><b>{len(lista)} resultado(s)</b><br><br>", unsafe_allow_html=True)

    if not lista:
        st.success("✅ Sin alertas en esta categoría")
    else:
        for a in lista:
            ic   = a["pct"]>=95
            css  = "alerta-crit" if ic else "alerta-warn"
            badge= "🔴 CRÍTICO"  if ic else "🟡 ALERTA"
            bcss = "alerta-badge-crit" if ic else "alerta-badge-warn"
            pcss = "alerta-pct-crit"   if ic else "alerta-pct-warn"
            bc   = "#dc2626" if ic else "#f59e0b"
            txt  = "#dc2626" if ic else "#b45309"
            pw   = min(100,a["pct"])
            st.markdown(f"""
            <div class="{css}">
              <div class="alerta-head">
                <div>
                  <div style="display:flex;align-items:center;margin-bottom:5px">
                    <span class="{bcss}">{badge}</span>
                    <span class="alerta-pais">País {a['pais']}</span>
                  </div>
                  <div class="alerta-linea">{a['linea']}</div>
                </div>
                <div>
                  <div class="{pcss}">{a['pct']:.1f}%</div>
                  <div class="alerta-sub">del límite consumido</div>
                </div>
              </div>
              <div class="alerta-bar-wrap"><div class="alerta-bar" style="background:{bc};width:{pw:.1f}%"></div></div>
              <div class="alerta-meta">
                <span>Límite: <strong>{fmt(a['limite'])}</strong></span>
                <span>Consumo: <strong style="color:{txt}">{fmt(a['consumo'])}</strong></span>
                <span>Disponible: <strong>{fmt(a['disponible'])}</strong></span>
              </div>
            </div>
            """, unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True)
