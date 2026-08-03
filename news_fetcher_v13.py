"""
News Fetcher v13 - Quaker Houghton Strategic News Dashboard
===========================================================
ENDMARKETS   — lado de la demanda: tendencias, estrategias y movimientos
               de mercado en las industrias que usan productos de QH.
PRODUCT LINE — lado de la oferta: nuevos productos, nuevos procesos e
               innovaciones de competidores en química de metales.

REGLAS DE CALIDAD:
  - Descarte político   : lobbying, geopolítica, sanciones, guerras comerciales
  - Descarte por fuente : blogs personales y blogs académicos/institucionales
  - Descarte listicle   : títulos que empiezan con número ("33 things", "10 ways")
  - Endmarkets          : debe contener al menos una palabra de mercado/industria
  - Product Lines       : filtro por niveles —
        Nivel 1 (ideal)  : innovación + sin plásticos
        Nivel 2 (fallback): solo excluye plásticos (para regiones con poca cobertura)
  - Sin duplicados por combinación: deduplicación por URL normalizada
    (detecta paginación: /article-name-40 == /article-name)

PROGRESO POR NIVELES:
  - Primera corrida : intenta cada combinación una sola vez
  - Corridas siguientes: reintenta hasta alcanzar TARGET_ARTICLES (default: 4)
    Prioridad 1 → nunca intentadas
    Prioridad 2 → tienen 1–3 artículos (necesitan más)
    Prioridad 3 → intentadas y vacías
    Saltadas    → ya tienen ≥ TARGET_ARTICLES artículos

CAMBIOS v13 (sobre v12):
  1. Etiquetas Endmarket actualizadas al nombre oficial del dashboard.
  2. Sistema de meta por combinación — reintenta combinaciones con 1–3 artículos
     hasta alcanzar TARGET_ARTICLES=4, sin volver a tocar las completas.

FECHAS: manuales — ajusta MONTH_START, MONTH_END y MONTH_LABEL cada mes.
PROGRESO: se guarda en JSON por mes. Nunca se pierde trabajo entre corridas.

INSTALACIÓN: pip install requests pandas openpyxl
"""

import requests
import pandas as pd
import time
import json
import os
import re
from datetime import datetime

# ─────────────────────────────────────────────
# CONFIGURACIÓN MENSUAL — AJUSTA ESTO CADA MES
# ─────────────────────────────────────────────

MONTH_START = "2026-06-01"
MONTH_END   = "2026-06-30"
MONTH_LABEL = "june_2026"

# ─────────────────────────────────────────────
# API KEYS — reemplaza con las tuyas
# ─────────────────────────────────────────────

ivan_newsapi = "4c68b0edd18b4faba942bdcd9631775a"
ivan_gnews   = "708137e49b93f4344667346622e8ab5c"

ale_newsapi = "fadb456f64a5423f84f41b8b7a0c58f3"
ale_gnews = "1a788fcc5d7f73012c2d4bda1728f9e1"
ale2_newsapi = "2393d8439a5448d9951feb6ad76dc94e"
ale2_gnews = "fdd87921889b9c88257a8c9d0ae6d88a"

emilio_newsapi = "3815e9feba024752be3145d5ca7d17d6"
emilio_gnews = "bdee22f1eca09e66ec757529bedc36f2"



NEWSAPI_KEY = ivan_newsapi
GNEWS_KEY   = ivan_gnews

# ─────────────────────────────────────────────
# PARÁMETROS GENERALES
# ─────────────────────────────────────────────

ARTICLES_PER_COMBO = 5   # cuántos pide la API por request
TARGET_ARTICLES    = 4   # meta por combinación — se sigue reintentando hasta llegar aquí
LIMIT_PER_API      = 100

OUTPUT_FILE   = f"noticias_dashboard_{MONTH_LABEL}.xlsx"
PROGRESS_FILE = f"news_progress_{MONTH_LABEL}.json"

# ─────────────────────────────────────────────
# REGIONES
# ─────────────────────────────────────────────

CONTINENT_KEYWORDS = {
    "Europe":          "Europe",
    "North America":   "America",
    "South America":   "Brazil",
    "Africa":          "Africa",
    "Middle East":     "Middle East",
    "North East Asia": "China",
    "South East Asia": "Asia",
    "Oceania":         "Australia",
}

# ─────────────────────────────────────────────
# ENDMARKETS — queries cortos orientados a demanda
# La brevedad maximiza resultados en la API.
# El filtro is_relevant() descarta lo irrelevante.
# ─────────────────────────────────────────────

ENDMARKET_QUERIES = {
    "Aerospace & Defense":         "aerospace defense industry",
    "Construction":                "construction industry market",
    "E-Mobility":                  "electric vehicle industry",
    "Home Appliances":             "home appliances market",
    "Mining":                      "mining metals industry",
    "Non Ferrous":                 "aluminum copper metals market",
    "Renewal Energy":              "renewable energy market",
    "Semiconductors":              "semiconductor chip industry",
    "Steel":                       "steel industry market",
    "Transportation - Components": "auto parts supply chain",
    "Tube & Pipe":                 "pipeline steel tube",
    "Agriculture & Equipment":     "agriculture equipment market",
    "Containers":                  "beverage aluminum packaging",
    "Industrial":                  "industrial machinery market",
    "Transportation - OEM":        "automotive industry market",
    "Fossil Energy":               "oil gas energy market",
}

# Palabras que debe contener el título/descripción
# para que un artículo de Endmarket sea válido.
ENDMARKET_REQUIRED = [
    "industry", "market", "demand", "production", "manufacturer",
    "growth", "strategy", "sector", "supply", "trade", "export",
    "import", "capacity", "output", "company", "group", "corp",
]

# ─────────────────────────────────────────────
# PRODUCT LINES — queries cortos orientados a oferta
# Primero: topic + metal. La región la agrega build_query().
# ─────────────────────────────────────────────

PRODUCT_LINE_QUERIES = {
    "Cleaning":                "industrial cleaning metal",
    "Die Casting":             "die casting process",
    "Forging":                 "metal forging process",
    "Forming":                 "metal forming process",
    "Greases":                 "industrial grease lubricant",
    "Heat Treatment":          "heat treatment metal",
    "Hydraulics & Lubricants": "hydraulic lubricant industrial",
    "Metal Protection":        "corrosion protection metal",
    "Metal Removal":           "metal machining cutting",
    "Non Ferrous Rolling":     "aluminum rolling process",
    "Specialty Additives":     "specialty chemical additive",
    "Specialty Coatings":      "industrial coating metal",
    "Steel Rolling":           "steel rolling mill",
    "Surface Treatment":       "surface treatment metal",
}

# Nivel 1 — palabras de innovación/novedad (ideal)
PRODUCT_LINE_REQUIRED = [
    "new", "launch", "innovat", "technolog", "process",
    "develop", "introduc", "solution", "release", "advance",
]

# Descarte duro — plásticos/no metales (aplica en ambos niveles)
PRODUCT_LINE_EXCLUDED = [
    "plastic", "polymer", "rubber", "composite", "resin",
    "polyurethane", "polypropylene", "polyethylene",
]

# ── Descarte político — aplica a TODAS las categorías ──────────────────
# Captura artículos cuyo dato central es una acción política,
# no un movimiento de mercado o tendencia industrial.
POLITICAL_EXCLUDED = [
    "lobbying", " lobby ", "geopolit", "diplomatic", "diplomacy",
    "sanction", "trade war", "trade-war", "trade dispute",
    "tariff war", "election", "vote for", "parliament",
    "congress ", "senate ", "legislation",
    "geopolitical", "trade tension",
]

# ── Descarte por fuente no periodística ────────────────────────────────
# Blogs personales, plataformas de opinión y blogs académicos/institucionales.
# Se compara contra la URL del artículo (en minúsculas).
BLOG_URL_PATTERNS = [
    ".blog/", ".blog.",
    "substack.com",
    "medium.com",
    "blogspot.com",
    "wordpress.com",
    "newyorkfed.org",        # Liberty Street Economics (blog académico)
    "federalreserve.gov/pubs",
    "imf.org/en/blogs",
    "worldbank.org/en/blogs",
    "oecd.org/en/blogs",
    "bruegel.org",           # think tank europeo
]

CONTINENTS    = list(CONTINENT_KEYWORDS.keys())
ENDMARKETS    = list(ENDMARKET_QUERIES.keys())
PRODUCT_LINES = list(PRODUCT_LINE_QUERIES.keys())

CONTINENT_ORDER   = {c: i for i, c in enumerate(CONTINENTS)}
ENDMARKET_ORDER   = {t: i for i, t in enumerate(ENDMARKETS)}
PRODUCTLINE_ORDER = {t: i for i, t in enumerate(PRODUCT_LINES)}

# ─────────────────────────────────────────────
# FALLBACKS para combinaciones históricamente vacías
# ─────────────────────────────────────────────

FALLBACK_QUERIES = {
    "Product Line|Europe|Greases":                    "grease lubrication metal Europe",
    "Product Line|South America|Forging":             "forging metal Latin America",
    "Product Line|South America|Greases":             "grease lubrication South America",
    "Product Line|South America|Non Ferrous Rolling": "aluminum rolling South America",
    "Product Line|Africa|Greases":                    "grease lubrication metal Africa",
    "Product Line|Africa|Non Ferrous Rolling":        "aluminum rolling Africa",
    "Product Line|Middle East|Greases":               "grease lubrication Middle East",
    "Product Line|Middle East|Metal Removal":         "machining cutting Middle East",
    "Product Line|Middle East|Non Ferrous Rolling":   "aluminum rolling Middle East",
    "Product Line|Oceania|Forging":                   "forging metal Australia",
    "Product Line|Oceania|Greases":                   "grease lubrication Australia",
}

# ─────────────────────────────────────────────
# FILTROS DE CALIDAD — SISTEMA POR NIVELES
# ─────────────────────────────────────────────

def _text(article: dict) -> str:
    return (
        (article.get("title") or "") + " " +
        (article.get("description") or "")
    ).lower()


def is_political(article: dict) -> bool:
    """
    Descarte político — aplica a TODAS las categorías.
    Rechaza artículos cuyo tema central es una acción política
    (lobbying, sanciones, guerras comerciales, legislación).
    """
    t = _text(article)
    return any(word in t for word in POLITICAL_EXCLUDED)


def is_blog(article: dict) -> bool:
    """
    Descarte por fuente — rechaza blogs personales y blogs académicos/institucionales.
    Se compara contra la URL del artículo.
    """
    url = (article.get("url") or "").lower()
    return any(pat in url for pat in BLOG_URL_PATTERNS)


def is_listicle(article: dict) -> bool:
    """
    Descarte listicle — rechaza artículos cuyo título empieza con un número.
    Ejemplos: "33 magnificently random things...", "10 best ways...", "5 reasons..."
    Estas piezas son contenido de entretenimiento/SEO, no noticias de mercado.
    """
    title = (article.get("title") or "").strip()
    return bool(re.match(r'^\d+\s+', title))


def _is_disqualified(article: dict) -> bool:
    """Descarte previo que aplica a cualquier nivel y categoría."""
    return is_political(article) or is_blog(article) or is_listicle(article)


def has_plastic(article: dict) -> bool:
    """Descarte duro: el artículo habla de plásticos/polímeros."""
    t = _text(article)
    return any(word in t for word in PRODUCT_LINE_EXCLUDED)


def is_relevant_level1(article: dict, category: str) -> bool:
    """
    Nivel 1 — filtro completo (ideal).
      Descarte previo  : político o blog → False (todas las categorías)
      Endmarket        : contiene palabra de mercado/industria.
      Product Line     : contiene palabra de innovación + sin plásticos.
    """
    if _is_disqualified(article):
        return False

    t = _text(article)

    if category == "Endmarket":
        return any(word in t for word in ENDMARKET_REQUIRED)

    # Product Line
    if has_plastic(article):
        return False
    return any(word in t for word in PRODUCT_LINE_REQUIRED)


def is_relevant_level2(article: dict, category: str) -> bool:
    """
    Nivel 2 — filtro relajado (solo para Product Lines con poca cobertura).
      El descarte político y de blogs sigue aplicando.
      Se acepta cualquier artículo que NO hable de plásticos.
    """
    if category != "Product Line":
        return False
    if _is_disqualified(article):
        return False
    return not has_plastic(article)


def filter_articles(raw: list, category: str) -> tuple[list, int]:
    """
    Aplica filtros por nivel. Devuelve (artículos_válidos, nivel_usado).
      nivel 1 = filtro completo
      nivel 2 = solo excluye plásticos
      nivel 0 = sin resultados
    """
    level1 = [a for a in raw if is_relevant_level1(a, category)]
    if level1:
        return level1, 1

    if category == "Product Line":
        level2 = [a for a in raw if is_relevant_level2(a, category)]
        if level2:
            return level2, 2

    return [], 0


# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────

def build_all_combos() -> list:
    combos = []
    for continent in CONTINENTS:
        for topic in ENDMARKETS:
            combos.append(("Endmarket", continent, topic, "endmarket"))
    for continent in CONTINENTS:
        for topic in PRODUCT_LINES:
            combos.append(("Product Line", continent, topic, "product_line"))
    return combos  # 240 total


def combo_key(category: str, continent: str, topic: str) -> str:
    return f"{category}|{continent}|{topic}"


def build_query(category: str, continent: str, topic: str, topic_type: str) -> str:
    ckey = combo_key(category, continent, topic)
    if ckey in FALLBACK_QUERIES:
        return FALLBACK_QUERIES[ckey]

    region = CONTINENT_KEYWORDS.get(continent, continent)
    base   = (
        ENDMARKET_QUERIES.get(topic)
        if topic_type == "endmarket"
        else PRODUCT_LINE_QUERIES.get(topic, topic)
    )
    return f"{base} {region}"


def normalize_url(url: str) -> str:
    """
    Normaliza una URL para deduplicación robusta.
    Elimina sufijos de paginación (-N, -NN) al final del path.
    Ejemplo:
      /article-name-40  →  /article-name
      /article-name-30  →  /article-name   (mismo artículo, URL distinta)
    """
    return re.sub(r'-\d+$', '', url.rstrip('/'))


def extract_global_parent(article: dict) -> str:
    """Usa el nombre de la fuente como Global Parent."""
    source = article.get("source") or ""
    return source if source else "N/A"


# ─────────────────────────────────────────────
# FETCH NEWSAPI
# ─────────────────────────────────────────────

def fetch_newsapi(query: str) -> list:
    params = {
        "q":        query,
        "from":     MONTH_START,
        "to":       MONTH_END,
        "sortBy":   "relevancy",
        "language": "en",
        "pageSize": ARTICLES_PER_COMBO,
        "apiKey":   NEWSAPI_KEY,
    }
    try:
        r    = requests.get("https://newsapi.org/v2/everything", params=params, timeout=10)
        data = r.json()
        if data.get("status") != "ok":
            print(f"    ⚠ NewsAPI: {data.get('message', 'error')}")
            return []
        return [
            {
                "title":        a.get("title") or "",
                "source":       a.get("source", {}).get("name") or "",
                "published_at": (a.get("publishedAt") or "")[:10],
                "url":          a.get("url") or "",
                "description":  a.get("description") or "",
            }
            for a in data.get("articles", [])
        ]
    except Exception as e:
        print(f"    ✗ NewsAPI error: {e}")
        return []


# ─────────────────────────────────────────────
# FETCH GNEWS
# ─────────────────────────────────────────────

def fetch_gnews(query: str) -> list:
    params = {
        "q":      query,
        "lang":   "en",
        "max":    ARTICLES_PER_COMBO,
        "from":   f"{MONTH_START}T00:00:00Z",
        "to":     f"{MONTH_END}T23:59:59Z",
        "apikey": GNEWS_KEY,
    }
    try:
        r    = requests.get("https://gnews.io/api/v4/search", params=params, timeout=10)
        data = r.json()
        if "articles" not in data:
            print(f"    ⚠ GNews: {data.get('errors', 'error')}")
            return []
        return [
            {
                "title":        a.get("title") or "",
                "source":       a.get("source", {}).get("name") or "",
                "published_at": (a.get("publishedAt") or "")[:10],
                "url":          a.get("url") or "",
                "description":  a.get("description") or "",
            }
            for a in data.get("articles", [])
        ]
    except Exception as e:
        print(f"    ✗ GNews error: {e}")
        return []


# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────

def run():
    attempted  = set()
    used_urls  = {}
    art_counts = {}
    all_rows   = []

    if os.path.exists(PROGRESS_FILE):
        with open(PROGRESS_FILE, "r") as f:
            saved = json.load(f)
        all_rows   = saved.get("rows", [])
        attempted  = set(tuple(c) for c in saved.get("attempted", []))
        used_urls  = {k: set(v) for k, v in saved.get("used_urls", {}).items()}
        art_counts = saved.get("art_counts", {})
        with_results = sum(1 for v in art_counts.values() if v > 0)
        print(f"📂 Progreso cargado: {len(attempted)}/240 intentadas, "
              f"{with_results} con artículos\n")

    all_combos = build_all_combos()
    total      = len(all_combos)

    # ── Clasificación de combinaciones ───────────────────────────────────
    # Prioridad 1: nunca intentadas → se procesan siempre primero
    # Prioridad 2: tienen 1–3 artículos → se reintentan hasta llegar a TARGET_ARTICLES
    # Prioridad 3: intentadas y vacías → se reintentan al final
    # Saltadas   : ya tienen ≥ TARGET_ARTICLES artículos
    not_tried   = [
        c for c in all_combos
        if (c[0], c[1], c[2]) not in attempted
    ]
    needs_more  = [
        c for c in all_combos
        if (c[0], c[1], c[2]) in attempted
        and 0 < art_counts.get(combo_key(c[0], c[1], c[2]), 0) < TARGET_ARTICLES
    ]
    empty_retry = [
        c for c in all_combos
        if (c[0], c[1], c[2]) in attempted
        and art_counts.get(combo_key(c[0], c[1], c[2]), 0) == 0
    ]
    skip_count  = sum(
        1 for c in all_combos
        if art_counts.get(combo_key(c[0], c[1], c[2]), 0) >= TARGET_ARTICLES
    )

    # ── Intercalar Endmarket / Product Line ──────────────────────────────
    # Alterna EM → PL para que ambas categorías consuman APIs por igual
    # sin importar cuántas combinaciones queden pendientes.
    def _interleave(combos: list) -> list:
        em = [c for c in combos if c[0] == "Endmarket"]
        pl = [c for c in combos if c[0] == "Product Line"]
        result = []
        for pair in zip(em, pl):
            result.extend(pair)
        result.extend(em[len(pl):])
        result.extend(pl[len(em):])
        return result

    to_process = (
        _interleave(not_tried) +
        _interleave(needs_more) +
        _interleave(empty_retry)
    )

    print(f"Periodo                                 : {MONTH_START} → {MONTH_END}")
    print(f"Meta por combinación                    : {TARGET_ARTICLES} artículos")
    print(f"Total combinaciones                     : {total}")
    print(f"✅ Completas ≥{TARGET_ARTICLES} artículos (se saltan) : {skip_count}")
    print(f"📈 Con 1–{TARGET_ARTICLES-1} artículos (necesitan más)  : {len(needs_more)}")
    print(f"🆕 Sin intentar                         : {len(not_tried)}")
    print(f"↩  Reintentando vacías                  : {len(empty_retry)}")
    print(f"Procesando hoy                          : {min(len(to_process), LIMIT_PER_API * 2)}\n")

    if not to_process:
        print(f"✅ Todas las combinaciones tienen ≥{TARGET_ARTICLES} artículos. Regenerando Excel...")
        export_excel(all_rows, art_counts, total)
        return

    newsapi_count = 0
    gnews_count   = 0
    done_today    = 0

    try:
        for (category, continent, topic, topic_type) in to_process:

            if newsapi_count < LIMIT_PER_API:
                api = "newsapi"
                newsapi_count += 1
            elif gnews_count < LIMIT_PER_API:
                api = "gnews"
                gnews_count += 1
            else:
                print("\n⏸  Límite de ambas APIs alcanzado (200 requests).")
                print("   Corre el script mañana para continuar.")
                break

            done_today += 1
            ckey     = combo_key(category, continent, topic)
            query    = build_query(category, continent, topic, topic_type)
            existing = art_counts.get(ckey, 0)

            print(f"[{done_today}/{min(len(to_process), LIMIT_PER_API * 2)}] "
                  f"[{api.upper()}] {continent} × {topic}  | '{query}'")

            raw = fetch_newsapi(query) if api == "newsapi" else fetch_gnews(query)

            # Deduplicación por combinación — usa URL normalizada para detectar
            # paginación: /article-30, /article-40, etc. se tratan como el mismo artículo.
            if ckey not in used_urls:
                used_urls[ckey] = set()

            raw_deduped = [
                a for a in raw
                if a["url"] and normalize_url(a["url"]) not in used_urls[ckey]
            ]

            # Filtrado por niveles
            new_articles, level = filter_articles(raw_deduped, category)

            # Registrar URLs normalizadas de los artículos que pasaron el filtro
            for a in new_articles:
                used_urls[ckey].add(normalize_url(a["url"]))

            if new_articles:
                # Eliminar fila "No results found" previa si existía
                all_rows = [
                    r for r in all_rows
                    if not (
                        r["Category"] == category
                        and r["Continent"] == continent
                        and r["Topic"] == topic
                        and r["News_Headline"] == "No results found"
                    )
                ]
                for art in new_articles:
                    all_rows.append({
                        "Category":          category,
                        "Continent":         continent,
                        "Topic":             topic,
                        "News_Headline":     art["title"],
                        "News_Link":         art["url"],
                        "Fecha_Publicacion": art["published_at"],
                        "Global_Parent":     extract_global_parent(art),
                    })
                art_counts[ckey] = existing + len(new_articles)

                level_tag = "" if level == 1 else " ⚠ Nivel 2 (sin filtro innovación)"
                print(f"    ✓ {len(new_articles)} artículo(s) — total: {art_counts[ckey]}{level_tag}")

            else:
                if (category, continent, topic) not in attempted:
                    all_rows.append({
                        "Category":          category,
                        "Continent":         continent,
                        "Topic":             topic,
                        "News_Headline":     "No results found",
                        "News_Link":         "",
                        "Fecha_Publicacion": "",
                        "Global_Parent":     "",
                    })
                art_counts[ckey] = existing

                total_raw     = len(raw)
                total_deduped = len(raw_deduped)
                n_political   = sum(1 for a in raw_deduped if is_political(a))
                n_blog        = sum(1 for a in raw_deduped if is_blog(a) and not is_political(a))
                n_listicle    = sum(1 for a in raw_deduped if is_listicle(a) and not is_political(a) and not is_blog(a))
                n_filter      = total_deduped - len(new_articles) - n_political - n_blog - n_listicle

                msg = "sin resultados válidos"
                reasons = []
                if total_raw == 0:
                    msg += " (API no regresó artículos)"
                else:
                    if n_political: reasons.append(f"{n_political} políticos")
                    if n_blog:      reasons.append(f"{n_blog} blogs")
                    if n_listicle:  reasons.append(f"{n_listicle} listicles")
                    if n_filter:    reasons.append(f"{n_filter} por filtro de relevancia")
                    if reasons:     msg += f" ({', '.join(reasons)} descartados)"
                print(f"    ↩ {msg}")

            attempted.add((category, continent, topic))
            time.sleep(0.6)

    finally:
        with open(PROGRESS_FILE, "w") as f:
            json.dump({
                "rows":       all_rows,
                "attempted":  [list(c) for c in attempted],
                "used_urls":  {k: list(v) for k, v in used_urls.items()},
                "art_counts": art_counts,
            }, f)
        print(f"\n💾 Progreso guardado en {PROGRESS_FILE}")

    export_excel(all_rows, art_counts, total)


# ─────────────────────────────────────────────
# EXPORTAR EXCEL
# ─────────────────────────────────────────────

def export_excel(all_rows: list, art_counts: dict, total: int):
    if not all_rows:
        print("⚠ No hay datos para exportar todavía.")
        return

    df = pd.DataFrame(all_rows)

    COLUMNS = [
        "Continent", "Topic", "News_Headline",
        "News_Link", "Fecha_Publicacion", "Global_Parent",
    ]

    df["_topic_order"] = df.apply(
        lambda r: ENDMARKET_ORDER.get(r["Topic"], 999)
        if r["Category"] == "Endmarket"
        else PRODUCTLINE_ORDER.get(r["Topic"], 999),
        axis=1,
    )
    df["_continent_order"] = df["Continent"].map(CONTINENT_ORDER).fillna(999)

    df = (
        df.sort_values(["Category", "_topic_order", "_continent_order"])
          .drop(columns=["_topic_order", "_continent_order"])
    )

    df_end  = df[df["Category"] == "Endmarket"][COLUMNS].rename(columns={"Topic": "End_Market"})
    df_prod = df[df["Category"] == "Product Line"][COLUMNS].rename(columns={"Topic": "Product_Line"})

    col_widths = {
        "Continent": 18, "End_Market": 28, "Product_Line": 28,
        "News_Headline": 65, "News_Link": 70,
        "Fecha_Publicacion": 18, "Global_Parent": 30,
    }

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        df_end.to_excel(writer,  sheet_name="Continents x Endmarkets",   index=False)
        df_prod.to_excel(writer, sheet_name="Continents x Product Lines", index=False)

        for sheet_name, df_sheet in [
            ("Continents x Endmarkets",   df_end),
            ("Continents x Product Lines", df_prod),
        ]:
            ws = writer.sheets[sheet_name]
            for col_idx, col_name in enumerate(df_sheet.columns, start=1):
                col_letter = ws.cell(row=1, column=col_idx).column_letter
                ws.column_dimensions[col_letter].width = col_widths.get(col_name, 20)

    with_results = sum(1 for v in art_counts.values() if v > 0)
    level2_combos = sum(
        1 for r in all_rows
        if r.get("News_Headline") != "No results found"
        # (el nivel 2 no se guarda en la fila, pero se registra en el log)
    )
    print(f"\n✅ Excel exportado: {OUTPUT_FILE}")
    print(f"   {with_results}/{total} combinaciones con artículos")
    print(f"   {total - with_results} sin resultados — corre mañana para reintentar")


# ─────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────

if __name__ == "__main__":
    if NEWSAPI_KEY == "TU_NEWSAPI_KEY_AQUI":
        print("❌ Reemplaza NEWSAPI_KEY con tu key de newsapi.org")
        exit(1)

    print("=" * 60)
    print("  Quaker Houghton — Strategic News Dashboard v13")
    print(f"  {datetime.now().strftime('%Y-%m-%d %H:%M')} — Monterrey, MX")
    print("=" * 60 + "\n")

    run()
