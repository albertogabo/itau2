import subprocess, sys

# Auto-install missing dependencies
for pkg in ["openpyxl", "pandas", "streamlit"]:
    try:
        __import__(pkg)
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg, "-q"])

import streamlit as st
import pandas as pd
import base64
from pathlib import Path

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Itaú · Cupos Disponibles",
    page_icon="🏦",
    layout="wide",
)

# ── Brand colors ──────────────────────────────────────────────────────────────
ORANGE      = "#FF6200"
ORANGE_DARK = "#cc4f00"
ORANGE_BG   = "#fff3ec"

# ── Custom CSS (Nunito font + Itaú styling) ───────────────────────────────────
st.markdown("""
<link href="https://fonts.googleapis.com/css2?family=Nunito:wght@400;600;700;800;900&display=swap" rel="stylesheet">
<style>
  html, body, [class*="css"] { font-family: 'Nunito', sans-serif !important; }

  /* Header strip */
  .itau-header {
    background: #FF6200;
    padding: 20px 32px;
    border-radius: 0 0 12px 12px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    margin-bottom: 28px;
  }
  .itau-header h1 { color: #fff; margin: 0; font-size: 22px; font-weight: 800; }
  .itau-header p  { color: rgba(255,255,255,0.85); margin: 4px 0 0; font-size: 12px; }

  /* Footer strip */
  .itau-footer {
    background: #FF6200;
    padding: 48px 32px;
    text-align: center;
    border-radius: 12px 12px 0 0;
    margin-top: 40px;
  }
  .itau-footer p {
    color: #fff;
    font-size: 20px;
    font-weight: 700;
    font-style: italic;
    letter-spacing: 3px;
    text-transform: uppercase;
    margin: 0;
  }

  /* ID header card */
  .id-header {
    background: #FF6200;
    border-radius: 14px;
    padding: 20px 28px;
    margin-bottom: 22px;
    display: flex;
    gap: 28px;
    align-items: center;
    flex-wrap: wrap;
    box-shadow: 0 4px 12px rgba(255,98,0,0.25);
  }
  .id-header .field { display: flex; flex-direction: column; }
  .id-header .label { color: rgba(255,255,255,0.8); font-size: 10px; text-transform: uppercase; letter-spacing: 1.5px; margin-bottom: 3px; }
  .id-header .value { color: #fff; font-size: 24px; font-weight: 800; }
  .id-header .rating-badge { background: #fff; color: #FF6200; border-radius: 8px; padding: 3px 14px; font-size: 18px; font-weight: 800; display: inline-block; }
  .id-header .sep { width: 1px; height: 44px; background: rgba(255,255,255,0.3); }

  /* Cupo card */
  .cupo-card {
    background: #fff;
    border-radius: 14px;
    padding: 22px 20px 18px;
    border-top: 5px solid #FF6200;
    box-shadow: 0 2px 8px rgba(0,0,0,0.08);
    height: 100%;
  }
  .cupo-card .tipo { font-size: 11px; color: #777; font-weight: 700; text-transform: uppercase; letter-spacing: 0.8px; margin-bottom: 10px; }
  .cupo-card .val-main { font-size: 20px; font-weight: 800; }
  .cupo-card .val-label { font-size: 10px; color: #aaa; margin-bottom: 3px; }
  .cupo-card .sep-line { height: 1px; background: #f0f0f0; margin: 12px 0; }
  .cupo-card .val-blue { color: #1d4ed8; }
  .cupo-card .val-muted { color: #bbb; }

  /* Alerta card */
  .alerta-critica { background: #fef2f2; border-left: 6px solid #dc2626; border-radius: 12px; padding: 16px 20px; margin-bottom: 12px; }
  .alerta-media   { background: #fffbeb; border-left: 6px solid #f59e0b; border-radius: 12px; padding: 16px 20px; margin-bottom: 12px; }

  /* Metric override */
  [data-testid="metric-container"] { background: #fff; border-radius: 12px; padding: 16px 20px; box-shadow: 0 1px 4px rgba(0,0,0,0.07); }

  /* Section titles */
  .section-title { font-size: 13px; font-weight: 800; color: #222; text-transform: uppercase; letter-spacing: 1.2px; margin-bottom: 4px; }
  .section-sub   { font-size: 11px; color: #aaa; margin-bottom: 14px; }

  /* Nav card */
  .nav-card {
    border-radius: 20px;
    padding: 40px 28px;
    text-align: center;
    cursor: pointer;
    transition: transform 0.15s;
    box-shadow: 0 2px 12px rgba(0,0,0,0.08);
  }
  .nav-card:hover { transform: translateY(-4px); }

  /* Streamlit cleanup */
  #MainMenu, footer, header { visibility: hidden; }
  .block-container { padding-top: 0 !important; }
  div[data-testid="stVerticalBlock"] { gap: 0; }
</style>
""", unsafe_allow_html=True)


# ── Data loading ──────────────────────────────────────────────────────────────
@st.cache_data
def load_data():
    # Consumo (USD)
    df_c = pd.read_excel("Consumo_Bancos_Internacionales_02062026.xlsx")
    df_c = df_c[["ID Solcred", "Entidad", "País",
                 "Disponible Trade & Garantias",
                 "Disponible Gestão de Caixa",
                 "Disponible Derivativos",
                 "Disponible Bonds",
                 "Total Disponible", "RATING"]].dropna(subset=["ID Solcred"])
    df_c["ID Solcred"] = df_c["ID Solcred"].astype(int)
    df_c["País"] = df_c["País"].astype(str)

    # Riesgo País (MM USD → USD)
    df_r = pd.read_excel("Riesgo_País_20260602.xlsx", header=None)
    df_r = df_r.iloc[2:, :8].copy()
    df_r.columns = ["País","Línea","LímiteTotal","ConsumoTotal","DisponibleTotal","LímiteLíq","ConsumoLíq","DisponibleLíq"]
    df_r = df_r[df_r["País"].notna()].copy()
    df_r["País"] = df_r["País"].astype(str)
    for col in ["LímiteTotal","ConsumoTotal","DisponibleTotal"]:
        df_r[col] = pd.to_numeric(df_r[col], errors="coerce").fillna(0) * 1_000_000

    return df_c, df_r

df_consumo, df_riesgo = load_data()

PRODUCT_MAP = {
    "Trade & Garantías":  ("Disponible Trade & Garantias",  "4. Trade & Garantias"),
    "Gestão de Caixa":    ("Disponible Gestão de Caixa",    "2. Gestão de Caixa"),
    "Derivativos":        ("Disponible Derivativos",         "3. Derivativos"),
    "Bonds":              ("Disponible Bonds",               "5. Bonds | Aplicações LP"),
}


# ── Helpers ───────────────────────────────────────────────────────────────────
def fmt_usd(n):
    if pd.isna(n) or n is None:
        return "—"
    n = float(n)
    if abs(n) >= 1_000_000:
        return f"{n/1_000_000:,.2f} MM USD"
    if abs(n) >= 1_000:
        return f"{n/1_000:,.1f} K USD"
    return f"{n:,.2f} USD"

def logo_b64():
    with open("itau_logo.png", "rb") as f:
        return base64.b64encode(f.read()).decode()

def get_alertas():
    alertas = []
    for pais in df_riesgo["País"].unique():
        df_p = df_riesgo[df_riesgo["País"] == pais]
        for _, row in df_p.iterrows():
            lim = row["LímiteTotal"]
            con = row["ConsumoTotal"]
            if lim and lim > 0:
                pct = (con / lim) * 100
                if pct >= 80:
                    alertas.append({
                        "pais": pais, "linea": row["Línea"],
                        "pct": pct, "limite": lim,
                        "consumo": con, "disponible": row["DisponibleTotal"]
                    })
    return sorted(alertas, key=lambda x: -x["pct"])


# ── Session state ─────────────────────────────────────────────────────────────
if "page" not in st.session_state:
    st.session_state.page = "home"
if "sel_id" not in st.session_state:
    st.session_state.sel_id = None


# ── Header (all pages) ────────────────────────────────────────────────────────
logo = logo_b64()
back_btn = ""
if st.session_state.page != "home":
    back_html = ""
else:
    back_html = ""

title_map = {
    "home":     "Dashboard de Cupos Disponibles",
    "busqueda": "Búsqueda por ID Solcred",
    "alertas":  "Alertas de Riesgo País",
}

st.markdown(f"""
<div class="itau-header">
  <div style="display:flex;align-items:center;gap:20px">
    <img src="data:image/png;base64,{logo}" style="height:110px;border-radius:10px">
  </div>
  <div style="text-align:right">
    <h1>{title_map[st.session_state.page]}</h1>
    <p>Valores en USD &nbsp;·&nbsp; Corte 02 Jun 2026</p>
  </div>
</div>
""", unsafe_allow_html=True)

if st.session_state.page != "home":
    if st.button("‹ Volver al inicio"):
        st.session_state.page = "home"
        st.rerun()
    st.markdown("<br>", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# HOME
# ══════════════════════════════════════════════════════════════════════════════
if st.session_state.page == "home":

    all_alertas = get_alertas()
    criticos    = [a for a in all_alertas if a["pct"] >= 95]
    normales    = [a for a in all_alertas if 80 <= a["pct"] < 95]
    paises_af   = len(set(a["pais"] for a in all_alertas))

    col1, col2 = st.columns(2, gap="large")

    # ── Búsqueda card
    with col1:
        st.markdown(f"""
        <div class="nav-card" style="background:#fff;border:2px solid {ORANGE}">
          <div style="font-size:52px;margin-bottom:16px">🔍</div>
          <div style="font-size:22px;font-weight:800;color:#222;margin-bottom:10px">Búsqueda por ID</div>
          <div style="font-size:14px;color:#888;line-height:1.6">Consulta cupos disponibles y riesgo país por ID Solcred</div>
          <div style="margin-top:20px;font-size:13px;color:{ORANGE};font-weight:700">{len(df_consumo)} registros disponibles →</div>
        </div>
        """, unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("Ir a Búsqueda →", key="btn_busqueda", use_container_width=True):
            st.session_state.page = "busqueda"
            st.rerun()

    # ── Alertas card
    with col2:
        alert_bg    = "#fef2f2" if criticos else "#fffbeb"
        alert_border = "#dc2626" if criticos else "#f59e0b"
        triangle_color = "#dc2626" if criticos else "#f59e0b"
        triangle_svg = f"""<svg width="56" height="50" viewBox="0 0 52 46" fill="none" xmlns="http://www.w3.org/2000/svg">
          <path d="M26 2L50 44H2L26 2Z" fill="{triangle_color}" stroke="{triangle_color}" stroke-width="2" stroke-linejoin="round"/>
          <text x="26" y="34" text-anchor="middle" font-size="20" font-weight="bold" fill="white">!</text>
        </svg>"""

        badges = f"""
          <span style="background:#dc2626;color:#fff;border-radius:20px;padding:4px 12px;font-size:12px;font-weight:700">🔴 {len(criticos)} críticas</span>
          <span style="background:#f59e0b;color:#fff;border-radius:20px;padding:4px 12px;font-size:12px;font-weight:700;margin-left:8px">🟡 {len(normales)} alertas</span>
          <span style="background:#6b7280;color:#fff;border-radius:20px;padding:4px 12px;font-size:12px;font-weight:700;margin-left:8px">{paises_af} países</span>
        """
        st.markdown(f"""
        <div class="nav-card" style="background:{alert_bg};border:2px solid {alert_border}">
          <div style="display:flex;justify-content:center;margin-bottom:16px">{triangle_svg}</div>
          <div style="font-size:22px;font-weight:800;color:#222;margin-bottom:10px">Alertas de Riesgo País</div>
          <div style="font-size:14px;color:#888;line-height:1.6">Líneas con consumo ≥ 80% del límite asignado por país</div>
          <div style="margin-top:20px">{badges}</div>
        </div>
        """, unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("Ver Alertas →", key="btn_alertas", use_container_width=True):
            st.session_state.page = "alertas"
            st.rerun()

    # Resumen global
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<div class="section-title">Resumen Global</div>', unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns(4)
    paises_total = len(df_consumo["País"].unique())
    for col, label, val, color in [
        (c1, "Entidades",      len(df_consumo),     "#222"),
        (c2, "Países",         paises_total,         "#222"),
        (c3, "Alertas activas",len(all_alertas),     "#f59e0b" if all_alertas else "#222"),
        (c4, "Casos críticos", len(criticos),        "#dc2626" if criticos else "#222"),
    ]:
        col.markdown(f"""
        <div style="background:#fff;border-radius:12px;padding:20px 22px;box-shadow:0 1px 4px rgba(0,0,0,0.07);text-align:center">
          <div style="font-size:32px;font-weight:900;color:{color}">{val}</div>
          <div style="font-size:12px;color:#888;margin-top:4px">{label}</div>
        </div>
        """, unsafe_allow_html=True)

    # Footer
    st.markdown("""
    <div class="itau-footer">
      <p>Cada resultado cuenta una historia</p>
    </div>
    """, unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# BÚSQUEDA
# ══════════════════════════════════════════════════════════════════════════════
elif st.session_state.page == "busqueda":

    all_ids = sorted(df_consumo["ID Solcred"].unique().tolist())

    sel_id = st.selectbox(
        "Selecciona o busca un ID Solcred",
        options=[None] + all_ids,
        format_func=lambda x: "— Selecciona un ID —" if x is None else f"ID {x}",
        index=0,
    )

    if sel_id:
        row = df_consumo[df_consumo["ID Solcred"] == sel_id].iloc[0]
        pais = str(row["País"])
        df_rp = df_riesgo[df_riesgo["País"] == pais]
        rp_total = df_rp[df_rp["Línea"] == "Total Transferência"]

        # ── ID Header
        rating = str(row["RATING"]) if pd.notna(row["RATING"]) else "N/A"
        total_disp = row["Total Disponible"] if pd.notna(row["Total Disponible"]) else 0
        st.markdown(f"""
        <div class="id-header">
          <div class="field"><span class="label">ID Solcred</span><span class="value" style="font-size:30px">{sel_id}</span></div>
          <div class="sep"></div>
          <div class="field"><span class="label">País</span><span class="value">{pais}</span></div>
          <div class="sep"></div>
          <div class="field"><span class="label">Rating</span><span class="rating-badge">{rating}</span></div>
          <div style="margin-left:auto;text-align:right">
            <div class="label" style="color:rgba(255,255,255,0.8)">Total Disponible</div>
            <div class="value">{fmt_usd(total_disp)}</div>
          </div>
        </div>
        """, unsafe_allow_html=True)

        # ── Cupos Disponibles
        st.markdown('<div class="section-title">Cupos Disponibles por Tipo de Transacción</div>', unsafe_allow_html=True)
        st.markdown('<div class="section-sub">Consumo Bancos Internacionales (USD) · Disponible Riesgo País por línea (USD)</div>', unsafe_allow_html=True)

        cols = st.columns(4)
        for i, (tipo, (col_consumo, linea_riesgo)) in enumerate(PRODUCT_MAP.items()):
            val_consumo = float(row[col_consumo]) if pd.notna(row[col_consumo]) else 0
            rp_linea    = df_rp[df_rp["Línea"] == linea_riesgo]
            lim_usd     = float(rp_linea["LímiteTotal"].values[0])     if len(rp_linea) else None
            disp_riesgo = float(rp_linea["DisponibleTotal"].values[0]) if len(rp_linea) else None
            pct         = min(100, (val_consumo / lim_usd) * 100) if (lim_usd and lim_usd > 0) else 0
            bar_color   = "#dc2626" if pct >= 95 else "#f59e0b" if pct >= 80 else ORANGE

            consumo_color = ORANGE_DARK if val_consumo > 0 else "#bbb"
            riesgo_color  = "#1d4ed8" if (disp_riesgo is not None and disp_riesgo > 0) else "#bbb"

            bar_html = ""
            if lim_usd is not None:
                bar_html = f"""
                <div style="height:7px;background:#f0f0f0;border-radius:4px;margin-top:10px">
                  <div style="height:7px;border-radius:4px;background:{bar_color};width:{pct:.1f}%"></div>
                </div>
                <div style="font-size:10px;color:#aaa;margin-top:4px">Límite: {fmt_usd(lim_usd)} · <span style="color:{bar_color};font-weight:700">{pct:.1f}% consumido</span></div>
                """

            cols[i].markdown(f"""
            <div class="cupo-card">
              <div class="tipo">{tipo}</div>
              <div class="val-label">Disponible (Consumo)</div>
              <div class="val-main" style="color:{consumo_color}">{fmt_usd(val_consumo)}</div>
              <div class="sep-line"></div>
              <div class="val-label">Disponible País (Riesgo)</div>
              <div class="val-main" style="color:{riesgo_color}">{fmt_usd(disp_riesgo) if disp_riesgo is not None else "—"}</div>
              {bar_html}
            </div>
            """, unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # ── Alertas del ID
        alertas_id = []
        for _, lrow in df_rp.iterrows():
            lim = lrow["LímiteTotal"]
            con = lrow["ConsumoTotal"]
            if lim and lim > 0:
                pct = (con / lim) * 100
                if pct >= 80:
                    alertas_id.append({"linea": lrow["Línea"], "pct": pct, "limite": lim, "consumo": con, "disponible": lrow["DisponibleTotal"]})

        if alertas_id:
            st.markdown(f'<div class="section-title">⚠ {len(alertas_id)} Alertas de Riesgo País</div>', unsafe_allow_html=True)
            st.markdown('<div class="section-sub">Líneas con consumo ≥ 80% del límite asignado al país</div>', unsafe_allow_html=True)
            for a in sorted(alertas_id, key=lambda x: -x["pct"]):
                is_crit   = a["pct"] >= 95
                css_class = "alerta-critica" if is_crit else "alerta-media"
                badge     = "🔴 CRÍTICO" if is_crit else "🟡 ALERTA"
                txt_color = "#dc2626" if is_crit else "#b45309"
                bar_color = "#dc2626" if is_crit else "#f59e0b"
                pct_w     = min(100, a["pct"])
                st.markdown(f"""
                <div class="{css_class}">
                  <div style="display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px">
                    <div>
                      <div style="font-size:12px;font-weight:800;color:{txt_color};margin-bottom:4px">{badge}</div>
                      <div style="font-size:15px;font-weight:700;color:#222">{a['linea']}</div>
                    </div>
                    <div style="text-align:right">
                      <div style="font-size:28px;font-weight:900;color:{txt_color}">{a['pct']:.1f}%</div>
                      <div style="font-size:10px;color:#888">del límite consumido</div>
                    </div>
                  </div>
                  <div style="margin-top:10px;height:8px;background:rgba(0,0,0,0.08);border-radius:4px">
                    <div style="height:8px;border-radius:4px;background:{bar_color};width:{pct_w:.1f}%"></div>
                  </div>
                  <div style="display:flex;gap:24px;margin-top:10px;font-size:12px;color:#555;flex-wrap:wrap">
                    <span>Límite: <strong>{fmt_usd(a['limite'])}</strong></span>
                    <span>Consumo: <strong style="color:{txt_color}">{fmt_usd(a['consumo'])}</strong></span>
                    <span>Disponible: <strong>{fmt_usd(a['disponible'])}</strong></span>
                  </div>
                </div>
                """, unsafe_allow_html=True)
            st.markdown("<br>", unsafe_allow_html=True)

        # ── Tabla Riesgo País
        st.markdown(f'<div class="section-title">Riesgo País — País {pais}</div>', unsafe_allow_html=True)
        st.markdown('<div class="section-sub">Fuente: Riesgo País · convertido a USD</div>', unsafe_allow_html=True)

        if not rp_total.empty:
            rt = rp_total.iloc[0]
            c1, c2, c3 = st.columns(3)
            for col, label, val, bg, fc in [
                (c1, "Límite Total",     rt["LímiteTotal"],     "#fff",   ORANGE_DARK),
                (c2, "Consumo Total",    rt["ConsumoTotal"],    "#fff",   ORANGE_DARK),
                (c3, "Disponible Total", rt["DisponibleTotal"], ORANGE,   "#fff"),
            ]:
                col.markdown(f"""
                <div style="background:{bg};border-radius:12px;padding:16px 22px;box-shadow:0 1px 4px rgba(0,0,0,0.07);margin-bottom:16px">
                  <div style="font-size:10px;color:{'rgba(255,255,255,0.8)' if bg==ORANGE else '#888'};text-transform:uppercase;letter-spacing:1.2px;margin-bottom:5px">{label}</div>
                  <div style="font-size:22px;font-weight:800;color:{fc}">{fmt_usd(val)}</div>
                </div>
                """, unsafe_allow_html=True)

        if not df_rp.empty:
            tabla = df_rp[["Línea","LímiteTotal","ConsumoTotal","DisponibleTotal"]].copy()
            tabla.columns = ["Línea","Límite (USD)","Consumo (USD)","Disponible (USD)"]
            for c in ["Límite (USD)","Consumo (USD)","Disponible (USD)"]:
                tabla[c] = tabla[c].apply(fmt_usd)
            st.dataframe(tabla, use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════════════════════════════════════
# ALERTAS
# ══════════════════════════════════════════════════════════════════════════════
elif st.session_state.page == "alertas":

    all_alertas = get_alertas()
    criticos    = [a for a in all_alertas if a["pct"] >= 95]
    normales    = [a for a in all_alertas if 80 <= a["pct"] < 95]

    # Filter tabs
    filtro = st.radio("Filtrar por:", ["Todas", "Críticas (≥95%)", "Alertas (80-94%)"],
                      horizontal=True, label_visibility="collapsed")

    lista = criticos if "Críticas" in filtro else normales if "Alertas" in filtro else all_alertas

    st.markdown(f"<br><b>{len(lista)} resultado(s)</b>", unsafe_allow_html=True)

    if not lista:
        st.success("✅ Sin alertas en esta categoría")
    else:
        for a in lista:
            is_crit   = a["pct"] >= 95
            css_class = "alerta-critica" if is_crit else "alerta-media"
            badge     = "🔴 CRÍTICO" if is_crit else "🟡 ALERTA"
            txt_color = "#dc2626" if is_crit else "#b45309"
            bar_color = "#dc2626" if is_crit else "#f59e0b"
            pct_w     = min(100, a["pct"])
            st.markdown(f"""
            <div class="{css_class}">
              <div style="display:flex;justify-content:space-between;align-items:flex-start;flex-wrap:wrap;gap:10px">
                <div>
                  <div style="display:flex;align-items:center;gap:10px;margin-bottom:5px">
                    <span style="font-size:12px;font-weight:800;color:{txt_color}">{badge}</span>
                    <span style="font-size:12px;background:#e5e7eb;color:#374151;border-radius:6px;padding:2px 8px;font-weight:700">País {a['pais']}</span>
                  </div>
                  <div style="font-size:16px;font-weight:800;color:#222">{a['linea']}</div>
                </div>
                <div style="text-align:right">
                  <div style="font-size:30px;font-weight:900;color:{txt_color}">{a['pct']:.1f}%</div>
                  <div style="font-size:11px;color:#888">del límite consumido</div>
                </div>
              </div>
              <div style="margin-top:12px;height:9px;background:rgba(0,0,0,0.08);border-radius:4px">
                <div style="height:9px;border-radius:4px;background:{bar_color};width:{pct_w:.1f}%"></div>
              </div>
              <div style="display:flex;gap:28px;margin-top:12px;font-size:13px;color:#555;flex-wrap:wrap">
                <span>Límite: <strong>{fmt_usd(a['limite'])}</strong></span>
                <span>Consumo: <strong style="color:{txt_color}">{fmt_usd(a['consumo'])}</strong></span>
                <span>Disponible: <strong>{fmt_usd(a['disponible'])}</strong></span>
              </div>
            </div>
            """, unsafe_allow_html=True)
