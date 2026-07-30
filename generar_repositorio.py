#!/usr/bin/env python3
"""
generar_repositorio.py — Repositorio de Reportes · Gemco Medical Solutions.
Genera index.html desde el Excel de reportes, leído en vivo desde SharePoint
vía Microsoft Graph API (ver sharepoint.py) — no depende de ningún PC.

Variables de entorno requeridas:
- SP_TENANT_ID, SP_CLIENT_ID, SP_CLIENT_SECRET : credenciales del App
  Registration de Azure AD (las mismas que usa Panel Licitaciones).
- SP_FILE_URL_REPOSITORIO : link de "Compartir" del Excel de reportes.

Uso:
  python generar_repositorio.py
"""

import base64
from datetime import datetime
from pathlib import Path

import os

import pandas as pd

import sharepoint

# ── Rutas ──────────────────────────────────────────────────────────────────
BASE_DIR   = Path(__file__).parent
LOGO_FILE  = BASE_DIR / "Gemco-logo-blanco-v2.png"
OUTPUT_HTML = BASE_DIR / "index.html"

SP_FILE_URL_REPOSITORIO = os.environ["SP_FILE_URL_REPOSITORIO"]

# ── Orden de frecuencia (mayor a menor) ───────────────────────────────────
FREQ_ORDER = {"Diario": 0, "Semanal": 1, "Mensual": 2, "Trimestral": 3}

# ── Paleta de colores por fuente ──────────────────────────────────────────
SOURCE_COLORS = {
    "PENTA":            ("#eeeafe", "#3c2eb5"),
    "SAP":              ("#e4eef8", "#1a4080"),
    "MERCADO PÚBLICO":  ("#fdf3e3", "#7a4b0a"),
    "MERCADO PUBLICO":  ("#fdf3e3", "#7a4b0a"),
    "SALESFORCE":       ("#e3f5ee", "#1a5c3a"),
    "PIPELINE":         ("#fce9f2", "#7a1f50"),
    "INVENTARIO":       ("#eaf5e4", "#2a5c1a"),
    "FALCONSOFT":       ("#e3f3f8", "#1a5c6a"),
}

# ── Paleta de colores por estado ──────────────────────────────────────────
STATUS_COLORS = {
    "ACTIVO":        ("#e4f2eb", "#1a6b47"),
    "EN REVISION":   ("#fef3e2", "#7a4b0a"),
    "EN REVISIÓN":   ("#fef3e2", "#7a4b0a"),
    "DEPRECADO":     ("#fce8e8", "#7a1f1f"),
    "EN DESARROLLO": ("#e8eef8", "#1a3d7a"),
}


# ── Reportes adicionales (no viven en el Excel) ───────────────────────────
EXTRA_REPORTS = [
    {
        "reporte":     "Contratos Incardia",
        "descripcion": "Registro de todos los contratos de la competencia de Incardia con el objetivo de realizar un seguimiento oportuno de los clientes que deben ser visitados y de los insumos respecto de los cuales podrían levantarse futuras licitaciones.",
        "link":        "panel_contratos_incardia.html",
        "fuente":      "Falconsoft",
        "frecuencia":  "Trimestral",
        "estado":      "Activo",
        "ultima":      datetime(2026, 6, 8),
    },
]

# ── Helpers ────────────────────────────────────────────────────────────────

def encode_logo() -> str:
    if not LOGO_FILE.exists():
        return ""
    return base64.b64encode(LOGO_FILE.read_bytes()).decode()


def source_style(name: str) -> tuple:
    key = name.upper().strip()
    for k, v in SOURCE_COLORS.items():
        if key == k or key.startswith(k):
            return v
    return ("#f0f0f0", "#555555")


def status_style(name: str) -> tuple:
    return STATUS_COLORS.get(name.upper().strip(), ("#f0f0f0", "#555555"))


def esc(s: str) -> str:
    """Escapa caracteres especiales para atributos HTML."""
    return (str(s)
            .replace("&", "&amp;")
            .replace('"', "&quot;")
            .replace("'", "&#39;")
            .replace("<", "&lt;")
            .replace(">", "&gt;"))


def fmt_date(val) -> str:
    if val is None:
        return "—"
    try:
        if pd.isna(val):
            return "—"
    except (TypeError, ValueError):
        pass
    if isinstance(val, datetime):
        return val.strftime("%d/%m/%Y")
    s = str(val).strip()
    return s if s and s not in ("nan", "-", "") else "—"


# ── Carga de datos ─────────────────────────────────────────────────────────

def load_reports() -> pd.DataFrame:
    excel_bytes = sharepoint.descargar_excel(
        SP_FILE_URL_REPOSITORIO, "Repositorio_Reportes_PowerBI_Gemco.xlsx"
    )
    df = pd.read_excel(excel_bytes, header=4)
    df = df.dropna(how="all")
    df.columns = ["reporte", "descripcion", "link", "fuente", "frecuencia", "estado", "ultima"]
    df = df[df["reporte"].notna()].copy()

    for col in ["reporte", "descripcion", "link", "fuente", "frecuencia", "estado"]:
        df[col] = df[col].fillna("").astype(str).str.strip()
        df[col] = df[col].replace({"-": "", "nan": ""})

    # Agregar reportes extra (no viven en el Excel)
    if EXTRA_REPORTS:
        extras = pd.DataFrame(EXTRA_REPORTS)
        df = pd.concat([df, extras], ignore_index=True)

    # Ordenar de más frecuente a menos frecuente
    df["_ord"] = df["frecuencia"].apply(lambda v: FREQ_ORDER.get(str(v).strip(), 99))
    df = df.sort_values("_ord").drop(columns=["_ord"]).reset_index(drop=True)
    return df


# ── Renderizado de cada card ───────────────────────────────────────────────

def card_html(row) -> str:
    # Badges de fuente
    sources = [s.strip() for s in str(row["fuente"]).split(";")
               if s.strip() and s.strip() not in ("nan", "")]
    if sources:
        badges = "".join(
            f'<span class="badge" style="background:{source_style(s)[0]};color:{source_style(s)[1]}">{esc(s)}</span>'
            for s in sources
        )
    else:
        badges = '<span class="badge badge-empty">—</span>'

    # Chip de estado
    estado = row["estado"]
    if estado and estado not in ("nan", ""):
        sbg, sfg = status_style(estado)
        chip = f'<span class="chip" style="background:{sbg};color:{sfg}">{esc(estado)}</span>'
    else:
        chip = '<span class="chip chip-empty">—</span>'

    # Campos de texto
    nombre = esc(row["reporte"])
    desc   = esc(row["descripcion"]) if row["descripcion"] not in ("nan", "") else "—"
    freq   = row["frecuencia"] if row["frecuencia"] not in ("nan", "") else "—"
    ultima = fmt_date(row["ultima"])

    # Enlace
    link = str(row["link"]).strip()
    if link and link not in ("nan", "") and ("http" in link or link.endswith(".html")):
        label = "Abrir reporte ↗" if link.endswith(".html") else "Abrir en Power BI ↗"
        btn = f'<a class="btn-open" href="{esc(link)}" target="_blank" rel="noopener noreferrer">{label}</a>'
    else:
        btn = '<span class="btn-open btn-disabled">Sin enlace</span>'

    # Data attributes para filtros (normalizados, sin acentos, minúsculas)
    fuentes_data = esc(";".join(s.lower() for s in sources))
    estado_data  = esc(estado.lower())
    search_data  = esc(f"{row['reporte'].lower()} {row['descripcion'].lower()}")

    return f'''  <div class="card" data-fuentes="{fuentes_data}" data-estado="{estado_data}" data-search="{search_data}">
    <div class="card-top">
      <div class="badges">{badges}</div>
      {chip}
    </div>
    <h3 class="card-name">{nombre}</h3>
    <p class="card-desc">{desc}</p>
    <div class="card-meta">
      <div class="meta-row">
        <span class="meta-label">Frecuencia</span>
        <span class="meta-val">{esc(freq)}</span>
      </div>
      <div class="meta-row">
        <span class="meta-label">Última actualización</span>
        <span class="meta-val">{ultima}</span>
      </div>
    </div>
    {btn}
  </div>'''


# ── Generación del HTML completo ──────────────────────────────────────────

def generate_html() -> str:
    df = load_reports()
    logo_b64 = encode_logo()

    total       = len(df)
    activos     = len(df[df["estado"].str.strip().str.lower() == "activo"])
    en_revision = len(df[df["estado"].str.strip().str.lower().isin(["en revision", "en revisión"])])

    all_fuentes = sorted({
        s.strip()
        for val in df["fuente"]
        for s in str(val).split(";")
        if s.strip() and s.strip() not in ("nan", "")
    })
    all_estados = [e for e in df["estado"].unique() if e and e not in ("nan", "")]

    fuente_opts = "\n        ".join(
        f'<option value="{esc(f.lower())}">{esc(f)}</option>' for f in all_fuentes
    )
    estado_opts = "\n        ".join(
        f'<option value="{esc(e.lower())}">{esc(e)}</option>' for e in all_estados
    )

    cards = "\n".join(card_html(row) for _, row in df.iterrows())

    logo_tag = (
        f'<img src="data:image/png;base64,{logo_b64}" alt="Gemco Medical Solutions" class="logo">'
        if logo_b64 else
        '<span class="logo-fallback">Gemco Medical Solutions</span>'
    )

    gen_date = datetime.now().strftime("%d/%m/%Y %H:%M")

    return f'''<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Repositorio de Reportes — Gemco Medical Solutions</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link href="https://fonts.googleapis.com/css2?family=DM+Sans:opsz,wght@9..40,400;9..40,500;9..40,600;9..40,700&display=swap" rel="stylesheet">
  <style>
    *, *::before, *::after {{
      box-sizing: border-box;
      margin: 0;
      padding: 0;
    }}

    body {{
      font-family: "DM Sans", sans-serif;
      background: #f5f4f0;
      color: #1a1a1a;
      min-height: 100vh;
    }}

    /* ── Header ──────────────────────────────────────────── */
    .header {{
      position: sticky;
      top: 0;
      z-index: 100;
      background: #1a4f8a;
      padding: 0 2rem;
      height: 70px;
      display: flex;
      align-items: center;
      gap: 1rem;
      box-shadow: 0 2px 10px rgba(0, 0, 0, 0.25);
    }}
    .logo {{
      height: 50px;
      width: auto;
    }}
    .logo-fallback {{
      color: #fff;
      font-weight: 600;
      font-size: 1rem;
    }}
    .header-sep {{
      width: 1px;
      height: 28px;
      background: rgba(255, 255, 255, 0.3);
    }}
    .header-title {{
      color: rgba(255, 255, 255, 0.88);
      font-size: 0.95rem;
      font-weight: 500;
      letter-spacing: 0.01em;
    }}

    /* ── Hero ────────────────────────────────────────────── */
    .hero {{
      background: linear-gradient(160deg, #1e5799 0%, #0e2f56 100%);
      color: #fff;
      padding: 3.5rem 2rem 3rem;
      text-align: center;
      border-bottom: 1px solid rgba(0,0,0,0.15);
    }}
    .hero-eyebrow {{
      display: inline-block;
      font-size: 0.7rem;
      font-weight: 700;
      letter-spacing: 0.14em;
      text-transform: uppercase;
      color: rgba(255,255,255,0.55);
      margin-bottom: 0.75rem;
    }}
    .hero h1 {{
      font-size: clamp(1.6rem, 3.5vw, 2.6rem);
      font-weight: 700;
      letter-spacing: -0.025em;
      margin-bottom: 2rem;
      line-height: 1.15;
    }}
    .stats {{
      display: inline-flex;
      gap: 0.5rem;
      flex-wrap: wrap;
      justify-content: center;
    }}
    .stat {{
      background: rgba(255, 255, 255, 0.1);
      border: 1px solid rgba(255, 255, 255, 0.18);
      border-radius: 6px;
      padding: 0.45rem 1.1rem;
      font-size: 0.82rem;
      font-weight: 600;
      letter-spacing: 0.01em;
      display: flex;
      align-items: center;
      gap: 0.4rem;
    }}
    .stat-num {{
      font-size: 1.05rem;
      font-weight: 700;
    }}
    .stat-label {{
      opacity: 0.75;
      font-weight: 500;
    }}

    /* ── Toolbar ─────────────────────────────────────────── */
    .toolbar {{
      background: #fff;
      border-bottom: 1px solid #e5e3de;
      padding: 0.9rem 2rem;
      display: flex;
      gap: 0.65rem;
      align-items: center;
      flex-wrap: wrap;
    }}
    .toolbar input,
    .toolbar select {{
      font-family: inherit;
      font-size: 0.88rem;
      padding: 0.5rem 0.85rem;
      border: 1px solid #d5d3ce;
      border-radius: 8px;
      outline: none;
      background: #fafaf8;
      color: #1a1a1a;
      transition: border-color 0.15s, box-shadow 0.15s;
    }}
    .toolbar input {{
      flex: 1;
      min-width: 200px;
    }}
    .toolbar input:focus,
    .toolbar select:focus {{
      border-color: #1a4f8a;
      box-shadow: 0 0 0 3px rgba(26, 79, 138, 0.12);
    }}
    .toolbar select {{
      cursor: pointer;
    }}
    .filter-count {{
      margin-left: auto;
      font-size: 0.8rem;
      color: #999;
    }}

    /* ── Grid ────────────────────────────────────────────── */
    .grid-wrapper {{
      max-width: 1360px;
      margin: 0 auto;
      padding: 2rem;
    }}
    .grid {{
      display: grid;
      grid-template-columns: repeat(auto-fill, minmax(290px, 1fr));
      gap: 1.25rem;
    }}
    .no-results {{
      display: none;
      text-align: center;
      padding: 4rem 2rem;
      color: #999;
      font-size: 1rem;
      grid-column: 1 / -1;
    }}

    /* ── Card ────────────────────────────────────────────── */
    .card {{
      background: #fff;
      border-radius: 14px;
      padding: 1.35rem;
      box-shadow:
        0 1px 3px rgba(0, 0, 0, 0.06),
        0 4px 16px rgba(0, 0, 0, 0.05);
      display: flex;
      flex-direction: column;
      gap: 0.8rem;
      transition: box-shadow 0.2s, transform 0.2s;
    }}
    .card:hover {{
      box-shadow:
        0 4px 12px rgba(0, 0, 0, 0.1),
        0 10px 28px rgba(0, 0, 0, 0.08);
      transform: translateY(-3px);
    }}
    .card.hidden {{
      display: none;
    }}

    .card-top {{
      display: flex;
      align-items: flex-start;
      justify-content: space-between;
      gap: 0.5rem;
      flex-wrap: wrap;
    }}
    .badges {{
      display: flex;
      flex-wrap: wrap;
      gap: 0.3rem;
      flex: 1;
    }}
    .badge {{
      font-size: 0.7rem;
      font-weight: 700;
      padding: 0.2rem 0.55rem;
      border-radius: 10px;
      white-space: nowrap;
      letter-spacing: 0.01em;
    }}
    .badge-empty {{
      background: #f0f0f0;
      color: #999;
    }}
    .chip {{
      font-size: 0.7rem;
      font-weight: 700;
      padding: 0.2rem 0.6rem;
      border-radius: 10px;
      white-space: nowrap;
      flex-shrink: 0;
      letter-spacing: 0.01em;
    }}
    .chip-empty {{
      background: #f0f0f0;
      color: #999;
    }}

    .card-name {{
      font-size: 0.975rem;
      font-weight: 600;
      color: #111;
      line-height: 1.35;
    }}
    .card-desc {{
      font-size: 0.83rem;
      color: #5a5a5a;
      line-height: 1.55;
      flex: 1;
    }}

    .card-meta {{
      display: flex;
      flex-direction: column;
      gap: 0.3rem;
      padding-top: 0.2rem;
      border-top: 1px solid #f0efec;
    }}
    .meta-row {{
      display: flex;
      align-items: center;
      gap: 0.4rem;
      font-size: 0.78rem;
    }}
    .meta-label {{
      color: #aaa;
      min-width: 130px;
    }}
    .meta-val {{
      font-weight: 500;
      color: #444;
    }}

    .btn-open {{
      display: block;
      text-align: center;
      background: #1a4f8a;
      color: #fff;
      text-decoration: none;
      padding: 0.55rem 1rem;
      border-radius: 9px;
      font-size: 0.84rem;
      font-weight: 600;
      transition: background 0.15s;
      margin-top: 0.2rem;
    }}
    .btn-open:hover {{
      background: #0e2f56;
    }}
    .btn-disabled {{
      background: #e8e8e8;
      color: #bbb;
      cursor: default;
      pointer-events: none;
    }}

    /* ── Footer ──────────────────────────────────────────── */
    .footer {{
      text-align: center;
      padding: 2rem;
      color: #aaa;
      font-size: 0.78rem;
      border-top: 1px solid #e5e3de;
    }}

    /* ── Responsive ──────────────────────────────────────── */
    @media (max-width: 640px) {{
      .header {{
        padding: 0 1rem;
        height: 62px;
      }}
      .logo {{
        height: 40px;
      }}
      .hero {{
        padding: 2.5rem 1.25rem 2rem;
      }}
      .toolbar {{
        padding: 0.75rem 1rem;
      }}
      .toolbar input {{
        min-width: 160px;
      }}
      .grid-wrapper {{
        padding: 1.25rem 1rem;
      }}
      .grid {{
        gap: 1rem;
      }}
      .filter-count {{
        display: none;
      }}
    }}
  </style>
</head>
<body>

<header class="header">
  {logo_tag}
  <div class="header-sep"></div>
  <span class="header-title">Repositorio de Reportes &middot; Gemco Medical</span>
</header>

<section class="hero">
  <span class="hero-eyebrow">Gemco Medical Solutions</span>
  <h1>Repositorio de Reportes</h1>
  <div class="stats">
    <span class="stat"><span class="stat-num">{total}</span><span class="stat-label">Reportes</span></span>
    <span class="stat"><span class="stat-num">{activos}</span><span class="stat-label">Activos</span></span>
    <span class="stat"><span class="stat-num">{en_revision}</span><span class="stat-label">En revisión</span></span>
    <span class="stat"><span class="stat-num">{len(all_fuentes)}</span><span class="stat-label">Fuentes</span></span>
  </div>
</section>

<div class="toolbar">
  <input type="text" id="search" placeholder="🔍  Buscar reporte o descripción..." oninput="applyFilter()">
  <select id="sel-fuente" onchange="applyFilter()">
    <option value="">Todas las fuentes</option>
    {fuente_opts}
  </select>
  <select id="sel-estado" onchange="applyFilter()">
    <option value="">Todos los estados</option>
    {estado_opts}
  </select>
  <span class="filter-count" id="filter-count">{total} de {total} reportes</span>
</div>

<div class="grid-wrapper">
  <div class="grid" id="grid">
{cards}
    <p class="no-results" id="no-results">No se encontraron reportes con ese criterio de búsqueda.</p>
  </div>
</div>

<footer class="footer">
  <p>Gemco Medical Solutions &nbsp;·&nbsp; Generado el {gen_date}</p>
</footer>

<script>
  function normalize(s) {{
    return String(s || "")
      .normalize("NFD")
      .replace(/[\\u0300-\\u036f]/g, "")
      .toLowerCase()
      .trim();
  }}

  function applyFilter() {{
    const q   = normalize(document.getElementById("search").value);
    const fnt = normalize(document.getElementById("sel-fuente").value);
    const est = normalize(document.getElementById("sel-estado").value);

    const cards = document.querySelectorAll(".card");
    let visible = 0;

    cards.forEach(function(card) {{
      const search  = normalize(card.dataset.search);
      const fuentes = normalize(card.dataset.fuentes);
      const estado  = normalize(card.dataset.estado);

      const matchQ = !q   || search.includes(q);
      const matchF = !fnt || fuentes.split(";").some(function(f) {{ return f.trim().includes(fnt); }});
      const matchE = !est || estado.includes(est);

      if (matchQ && matchF && matchE) {{
        card.classList.remove("hidden");
        visible++;
      }} else {{
        card.classList.add("hidden");
      }}
    }});

    document.getElementById("no-results").style.display = visible === 0 ? "block" : "none";
    document.getElementById("filter-count").textContent = visible + " de {total} reportes";
  }}
</script>

</body>
</html>
'''


# ── Build ──────────────────────────────────────────────────────────────────

def build():
    html = generate_html()
    OUTPUT_HTML.write_text(html, encoding="utf-8")
    ts = datetime.now().strftime("%d/%m/%Y %H:%M")
    print(f"[OK] HTML regenerado — {ts}")

    import generar_contratos_incardia as gci
    gci.generar()


if __name__ == "__main__":
    build()
