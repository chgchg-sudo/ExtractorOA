"""
buscador_core.py: lógica del buscador de publicaciones (sin interfaz de usuario)

Consulta Crossref, OpenAlex, Unpaywall, CORE, DOAJ, Semantic Scholar y BASE a partir de un
Excel con columnas doi y/o título (y opcionalmente Author, Source, PubYear).

Punto de entrada para cualquier interfaz:
    ejecutar_ambos(ruta_excel, email, hacer_doi, hacer_titulo, fuentes, ...)
    -> {"resumen": str, "doi": (DataFrame, ruta_xlsx) | None, "titulo": (DataFrame, ruta_xlsx) | None}

Variable de entorno opcional BUSCADOR_CACHE: ruta del archivo SQLite de caché.
"""

import html
import json
import os
import re
import sqlite3
import threading
import time
import unicodedata
from concurrent.futures import ThreadPoolExecutor, as_completed
from difflib import SequenceMatcher
from urllib.parse import quote, quote_plus, unquote, urlparse

import pandas as pd
import requests

# =====================================================================
# 1. Utilidades de texto
# =====================================================================
DOI_RE = re.compile(r"10\.\d{4,9}/[^\s\"<>]+")


def norm(t):
    """Minúsculas, sin tildes ni signos (conserva letras de cualquier alfabeto)."""
    t = unicodedata.normalize("NFKD", str(t or ""))
    t = "".join(c for c in t if not unicodedata.combining(c)).lower()
    t = re.sub(r"[^\w]+|_", " ", t)
    return re.sub(r"\s+", " ", t).strip()


def limpiar_html(t):
    return re.sub(r"<[^>]+>", " ", html.unescape(str(t or "")))


ARTICULOS = re.compile(r"^(the|a|an|el|la|los|las|un|una|le|les|der|die|das)\s+")


def norm_titulo(t):
    return ARTICULOS.sub("", norm(limpiar_html(t)))


def titulo_principal(t):
    """Parte del título anterior a un subtítulo (: – — ? o ' - ')."""
    return re.split(r"\s*[:–—]\s*|\s+-\s+|\?\s+", limpiar_html(t), maxsplit=1)[0]


def extraer_doi(texto):
    """Devuelve un DOI limpio en minúsculas o '' si no hay ninguno."""
    m = DOI_RE.search(unquote(str(texto or "")))
    if not m:
        return ""
    d = m.group()
    while d:
        if d[-1] in ".,;:]}>'\"":
            d = d[:-1]
        elif d[-1] == ")" and d.count(")") > d.count("("):  # ')' sobrante, no del DOI
            d = d[:-1]
        else:
            break
    return d.lower()


def limpiar_doi(texto):
    return extraer_doi(texto)


def lista(v):
    if v is None:
        return []
    return [x for x in v if x] if isinstance(v, list) else [v]


def primero(v):
    l = lista(v)
    return str(l[0]) if l else ""


def familia(nombre):
    """Apellido a partir de 'Apellido, Nombre' o 'Nombre Apellido'."""
    nombre = str(nombre or "").strip()
    if "," in nombre:
        return nombre.split(",")[0].strip()
    partes = nombre.split()
    return partes[-1] if partes else ""


def anio(v):
    m = re.search(r"(?:1[5-9]|20)\d{2}", str(v or ""))
    return int(m.group()) if m else None


def si_no(v):
    return {True: "Sí", False: "No"}.get(v, "")


VERSIONES = {"publishedversion": "publicada", "acceptedversion": "aceptada",
             "submittedversion": "preprint/enviada"}


def version(v):
    return VERSIONES.get(norm(v).replace(" ", ""), "")


def consulta_titulo(meta):
    t = re.sub(r'["\\]', " ", limpiar_html(meta.get("titulo", "")))
    return re.sub(r"\s+", " ", t).strip()


# =====================================================================
# 2. HTTP con caché (SQLite), límites por servidor y reintentos
# =====================================================================
class ErrorAPI(Exception):
    pass


RUTA_CACHE = os.environ.get("BUSCADOR_CACHE", "cache_api.sqlite")
TTL_CACHE = 30 * 24 * 3600
_MISS = object()
_lock_cache = threading.Lock()
_con = None
_usar_cache = True
STATS = {"api": 0, "cache": 0}
_lock_stats = threading.Lock()

INTERVALOS = {"api.crossref.org": 0.1, "api.openalex.org": 0.1, "api.unpaywall.org": 0.1,
              "api.core.ac.uk": 6.0, "doaj.org": 0.5, "api.semanticscholar.org": 3.0,
              "api.base-search.net": 1.0}
_lock_rate = threading.Lock()
_proximo = {}
PARAMS_PRIVADOS = {"mailto", "email", "api_key"}

# Protección por servidor: si una API falla varias veces seguidas se suspende un rato
FALLOS_MAX = 5          # fallos consecutivos (tras reintentos) que activan la suspensión
SUSPENSION = 600        # segundos de suspensión
_lock_fallos = threading.Lock()
_fallos, _susp, _motivo, _penal = {}, {}, {}, {}


def _fallo(host, motivo, inmediato=False):
    with _lock_fallos:
        _fallos[host] = _fallos.get(host, 0) + 1
        if inmediato or _fallos[host] >= FALLOS_MAX:
            _susp[host] = time.time() + SUSPENSION
            _motivo[host] = motivo


def _exito(host):
    with _lock_fallos:
        _fallos[host] = 0
        if _penal.get(host, 0.0) > 0:
            _penal[host] = max(0.0, _penal[host] - 0.25)


def _penalizar(host):
    """Tras un 429, se espacian más las siguientes peticiones a ese servidor."""
    with _lock_fallos:
        _penal[host] = min(_penal.get(host, 0.0) + 1.0, 8.0)


def host_suspendido(host):
    with _lock_fallos:
        t = _susp.get(host)
        if t is None:
            return ""
        if time.time() >= t:
            _susp.pop(host, None)
            _fallos[host] = 0
            return ""
        return _motivo.get(host, "fallos repetidos")


def hosts_suspendidos():
    return [(h, host_suspendido(h)) for h in list(_susp) if host_suspendido(h)]


def configurar_cache(activa):
    global _usar_cache
    _usar_cache = bool(activa)


def reiniciar_stats():
    with _lock_stats:
        STATS["api"] = STATS["cache"] = 0
    with _lock_fallos:
        _fallos.clear(); _susp.clear(); _motivo.clear(); _penal.clear()


def _conexion():
    global _con
    if _con is None:
        _con = sqlite3.connect(RUTA_CACHE, check_same_thread=False)
        _con.execute("CREATE TABLE IF NOT EXISTS cache (k TEXT PRIMARY KEY, ts REAL, v TEXT)")
        _con.commit()
    return _con


def cache_get(k):
    if not _usar_cache:
        return _MISS
    try:
        with _lock_cache:
            fila = _conexion().execute("SELECT ts, v FROM cache WHERE k=?", (k,)).fetchone()
        if fila and time.time() - fila[0] < TTL_CACHE:
            return json.loads(fila[1])
    except Exception:
        pass
    return _MISS


def cache_set(k, v):
    if not _usar_cache:
        return
    try:
        with _lock_cache:
            con = _conexion()
            con.execute("INSERT OR REPLACE INTO cache VALUES (?,?,?)", (k, time.time(), json.dumps(v)))
            con.commit()
    except Exception:
        pass


def vaciar_cache():
    try:
        with _lock_cache:
            con = _conexion()
            n = con.execute("SELECT COUNT(*) FROM cache").fetchone()[0]
            con.execute("DELETE FROM cache")
            con.commit()
        return n
    except Exception:
        return 0


def _esperar(host):
    intervalo = INTERVALOS.get(host, 0.3) + _penal.get(host, 0.0)
    with _lock_rate:
        ahora = time.monotonic()
        t = max(ahora, _proximo.get(host, 0.0))
        _proximo[host] = t + intervalo
    if t > ahora:
        time.sleep(t - ahora)


def _detalle(r):
    try:
        return re.sub(r"\s+", " ", r.text or "").strip()[:120]
    except Exception:
        return ""


def get_json(url, params=None, headers=None, reintentos=4):
    """GET con caché, reintentos y protección por servidor. None si 404; ErrorAPI si falla."""
    params = params or {}
    k = url + "?" + "&".join(f"{a}={b}" for a, b in sorted(params.items())
                             if a not in PARAMS_PRIVADOS)
    c = cache_get(k)
    if c is not _MISS:
        with _lock_stats:
            STATS["cache"] += 1
        return c

    host = urlparse(url).netloc
    motivo = host_suspendido(host)
    if motivo:
        raise ErrorAPI(f"{host}: fuente suspendida temporalmente tras fallos repetidos ({motivo})")

    ultimo = "sin respuesta"
    for intento in range(reintentos):
        _esperar(host)
        try:
            r = requests.get(url, params=params, headers=headers, timeout=30)
        except requests.RequestException as e:
            ultimo = f"conexión ({type(e).__name__})"
            time.sleep(2 * (intento + 1))
            continue
        with _lock_stats:
            STATS["api"] += 1
        sc = r.status_code
        if sc == 404:
            _exito(host)
            cache_set(k, None)
            return None
        if sc == 429:
            try:
                espera = float(r.headers.get("Retry-After", 0))
            except ValueError:
                espera = 0.0
            _penalizar(host)
            ultimo = "límite de peticiones (429)"
            if intento < reintentos - 1:
                time.sleep(min(max(espera, 3 * 2 ** intento), 60))
            continue
        if sc in (401, 403):
            msg = f"HTTP {sc}: acceso no autorizado (clave o IP)"
            _fallo(host, msg, inmediato=True)
            raise ErrorAPI(f"{host}: {msg}")
        if sc >= 500:
            ultimo = f"HTTP {sc} {_detalle(r)}".strip()
            if intento < reintentos - 1:
                time.sleep(2 * (intento + 1))
            continue
        if sc >= 400:
            raise ErrorAPI(f"{host}: HTTP {sc} {_detalle(r)}".strip())
        try:
            data = r.json()
        except ValueError:
            _fallo(host, "respuesta no válida")
            raise ErrorAPI(f"respuesta no válida de {host} (¿clave o IP no autorizada?)")
        _exito(host)
        cache_set(k, data)
        return data
    _fallo(host, ultimo)
    raise ErrorAPI(f"{host}: {ultimo}")


# =====================================================================
# 3. Fuentes. Cada una devuelve una lista de "candidatos" normalizados:
#    titulo, autores (apellidos), anio, revista, doi, registro, pdf,
#    es_oa, estado_oa, ubicaciones[{lugar,url,pdf,version}]
# =====================================================================
def _cand(titulo="", autores=None, anio_=None, revista="", doi="", registro="", pdf="",
          es_oa=None, estado_oa="", ubicaciones=None, raw=None):
    return dict(titulo=titulo or "", autores=[a for a in (autores or []) if a], anio=anio_,
                revista=revista or "", doi=doi or "", registro=registro or "", pdf=pdf or "",
                es_oa=es_oa, estado_oa=estado_oa or "", ubicaciones=ubicaciones or [], raw=raw)


def _es_pdf(u):
    u = str(u or "").lower().split("?")[0]
    return u.endswith(".pdf") or "core.ac.uk/download/" in u


def _dedup_ubic(ubic):
    vistos, salida = set(), []
    for u in ubic:
        k = _norm_url(u["url"])
        if k not in vistos:
            vistos.add(k)
            salida.append(u)
    return salida


def fmt_ubic(ubic, maximo=30):
    """'Repositorio [versión]: URL | Repositorio: URL ...'"""
    return " | ".join(f"{u['lugar']}" + (f" [{u['version']}]" if u.get("version") else "")
                      + f": {u['url']}" for u in ubic[:maximo])


def _primer_pdf(ubic):
    return next((u["url"] for u in ubic if u.get("pdf")), "")


# ---- Crossref ----
def _cr_cand(m):
    a = None
    for k in ("issued", "published-print", "published-online"):
        dp = (m.get(k) or {}).get("date-parts") or [[None]]
        if dp and dp[0] and dp[0][0]:
            a = dp[0][0]
            break
    return _cand(primero(m.get("title")),
                 [x.get("family") or x.get("name") or "" for x in m.get("author") or []],
                 a, primero(m.get("container-title")), (m.get("DOI") or "").lower(),
                 m.get("URL", ""))


def cr_doi(doi, ctx):
    d = get_json(f"https://api.crossref.org/works/{quote(doi, safe='/')}", {"mailto": ctx["email"]})
    m = (d or {}).get("message")
    return [_cr_cand(m)] if m else []


def cr_titulo(meta, ctx):
    d = get_json("https://api.crossref.org/works",
                 {"query.bibliographic": consulta_titulo(meta), "rows": 8, "mailto": ctx["email"]})
    return [_cr_cand(m) for m in ((d or {}).get("message") or {}).get("items", [])]


# ---- OpenAlex ----
def _oa_params(ctx, **extra):
    p = {"mailto": ctx["email"]}
    if ctx.get("key_oa"):
        p["api_key"] = ctx["key_oa"]
    p.update(extra)
    return p


def _oa_cand(w):
    oa = w.get("open_access") or {}
    ubic = []
    for loc in w.get("locations") or []:
        if not loc.get("is_oa"):
            continue
        src = loc.get("source") or {}
        url = loc.get("pdf_url") or loc.get("landing_page_url")
        if url:
            ubic.append(dict(lugar=src.get("display_name") or "desconocido", url=url,
                             pdf=bool(loc.get("pdf_url")), version=version(loc.get("version"))))
    prim = (w.get("primary_location") or {}).get("source") or {}
    return _cand(w.get("title") or w.get("display_name"),
                 [familia((x.get("author") or {}).get("display_name")) for x in w.get("authorships") or []],
                 w.get("publication_year"), prim.get("display_name"), limpiar_doi(w.get("doi")),
                 w.get("id", ""), _primer_pdf(ubic), oa.get("is_oa"), oa.get("oa_status"), ubic)


def oa_doi(doi, ctx):
    w = get_json(f"https://api.openalex.org/works/doi:{quote(doi, safe='/')}", _oa_params(ctx))
    return [_oa_cand(w)] if w else []


def oa_titulo(meta, ctx):
    d = get_json("https://api.openalex.org/works",
                 _oa_params(ctx, search=consulta_titulo(meta), **{"per-page": 8}))
    return [_oa_cand(w) for w in (d or {}).get("results", [])]


# ---- Unpaywall (solo DOI) ----
def up_doi(doi, ctx):
    d = get_json(f"https://api.unpaywall.org/v2/{quote(doi, safe='/')}", {"email": ctx["email"]})
    if not d:
        return []
    locs = []
    if d.get("best_oa_location"):
        locs.append(d["best_oa_location"])
    locs += [l for l in d.get("oa_locations") or [] if l != d.get("best_oa_location")]
    ubic = []
    for l in locs:
        url = l.get("url_for_pdf") or l.get("url")
        if not url:
            continue
        if l.get("host_type") == "publisher":
            lugar = d.get("publisher") or "Editorial"
        else:
            lugar = l.get("repository_institution") or l.get("host_type") or "desconocido"
        ubic.append(dict(lugar=lugar, url=url, pdf=bool(l.get("url_for_pdf")),
                         version=version(l.get("version"))))
    return [_cand(d.get("title"), [x.get("family") or "" for x in d.get("z_authors") or []],
                  d.get("year"), d.get("journal_name"), limpiar_doi(d.get("doi")),
                  d.get("doi_url", ""), _primer_pdf(ubic), d.get("is_oa"), d.get("oa_status"), ubic)]


# ---- CORE ----
def _core_cand(r):
    links = {l.get("type"): l.get("url") for l in r.get("links") or [] if l.get("url")}
    web = links.get("display") or (f"https://core.ac.uk/works/{r.get('id')}" if r.get("id") else "")
    pdf = links.get("download") or r.get("downloadUrl") or ""
    provs = [p for p in r.get("dataProviders") or [] if isinstance(p, dict)]
    nombres = " / ".join(str(p.get("name")).strip("'") for p in provs if p.get("name"))
    ubic = [dict(lugar=nombres or "Repositorio (CORE)", url=u, pdf=_es_pdf(u), version="")
            for u in [u for u in lista(r.get("sourceFulltextUrls")) if u][:5]]
    if pdf and _norm_url(pdf) not in {_norm_url(u["url"]) for u in ubic}:
        ubic.append(dict(lugar="CORE (copia alojada)" if "core.ac.uk" in pdf else (nombres or "Repositorio (CORE)"),
                         url=pdf, pdf=_es_pdf(pdf), version=""))
    if not ubic and web:
        ubic.append(dict(lugar="CORE", url=web, pdf=False, version=""))
    jr = (r.get("journals") or [{}])[0] if r.get("journals") else {}
    return _cand(r.get("title"), [familia(a.get("name")) for a in r.get("authors") or []],
                 r.get("yearPublished"), jr.get("title", ""), limpiar_doi(r.get("doi")),
                 web, pdf, True if pdf else None, "", ubic,
                 raw=dict(outputs=lista(r.get("outputs")), providers=provs))


CORE_MAX_OUTPUTS = 5


def _core_hdr(ctx):
    return {"Authorization": f"Bearer {ctx['key_core']}"}


def _core_nombre(p, ctx):
    """Nombre del repositorio; en los resultados de búsqueda a menudo viene vacío y hay que consultarlo."""
    if p.get("name"):
        return str(p["name"]).strip().strip("'")
    pid = p.get("id")
    if not pid:
        return ""
    try:
        d = get_json(f"https://api.core.ac.uk/v3/data-providers/{pid}", None, _core_hdr(ctx), reintentos=2)
    except ErrorAPI:
        return f"Repositorio CORE #{pid}"
    return str((d or {}).get("name") or f"Repositorio CORE #{pid}").strip("'")


def core_expandir(cand, meta, ctx):
    """Una localización por repositorio: cada 'output' de CORE es la copia de la obra en un repositorio."""
    if not ctx.get("core_detalle", True):
        return
    raw = cand.get("raw") or {}
    ids = [m.group(1) for u in raw.get("outputs", []) for m in [re.search(r"(\d+)\s*$", str(u))] if m]
    nuevas = []
    for oid in ids[:CORE_MAX_OUTPUTS]:
        try:
            o = get_json(f"https://api.core.ac.uk/v3/outputs/{oid}", None, _core_hdr(ctx), reintentos=2)
        except ErrorAPI:
            break  # probablemente límite de uso: no se insiste
        if not o:
            continue
        provs = [p for p in o.get("dataProviders") or [] if isinstance(p, dict)]
        nombre = " / ".join(n for n in (_core_nombre(p, ctx) for p in provs) if n) or "Repositorio (CORE)"
        urls = [u for u in lista(o.get("sourceFulltextUrls")) if u] or \
               ([o["downloadUrl"]] if o.get("downloadUrl") else [])
        for u in urls[:2]:
            nuevas.append(dict(lugar=nombre, url=u, pdf=_es_pdf(u), version=""))
    if nuevas:
        propias = [u for u in cand["ubicaciones"] if u["lugar"].startswith("CORE")]
        cand["ubicaciones"] = _dedup_ubic(nuevas + propias)
    else:
        # sin detalle por output: al menos se resuelven los nombres de repositorio de la obra
        nombres = " / ".join(n for n in (_core_nombre(p, ctx) for p in raw.get("providers", [])[:CORE_MAX_OUTPUTS]) if n)
        if nombres:
            for u in cand["ubicaciones"]:
                if u["lugar"] == "Repositorio (CORE)":
                    u["lugar"] = nombres


def _core_buscar(q, ctx, n, reintentos=2):
    # La barra final es necesaria: sin ella CORE redirige (301)
    d = get_json("https://api.core.ac.uk/v3/search/works/", {"q": q, "limit": n},
                 {"Authorization": f"Bearer {ctx['key_core']}"}, reintentos=reintentos)
    return [_core_cand(r) for r in (d or {}).get("results", [])]


def _reintentable(e):
    t = str(e)
    return ("suspendida" not in t) and ("HTTP 5" in t or "respuesta no válida" in t)


def core_doi(doi, ctx):
    try:
        return _core_buscar(f'doi:"{doi}"', ctx, 3)
    except ErrorAPI as e:
        if not _reintentable(e):
            raise
        # alternativa: el DOI como texto; solo se aceptan registros con ese DOI exacto
        res = [c for c in _core_buscar(f'"{doi}"', ctx, 5) if c["doi"] == doi]
        _exito("api.core.ac.uk")  # la alternativa funcionó: el servidor está bien
        return res


def core_titulo(meta, ctx):
    # Sin signos que el buscador interpreta como operadores (: ( ) [ ] ^ ~ * ? ...)
    limpio = re.sub(r"\s+", " ", re.sub(r"[^\w\s]|_", " ", limpiar_html(meta.get("titulo", "")))).strip()
    try:
        return _core_buscar(f'title:"{limpio}"', ctx, 5)
    except ErrorAPI as e:
        if not _reintentable(e):
            raise
        res = _core_buscar(" ".join(limpio.split()[:12]), ctx, 5)
        _exito("api.core.ac.uk")  # la alternativa funcionó: el servidor está bien
        return res


# ---- DOAJ ----
def _doaj_cand(r):
    b = r.get("bibjson") or {}
    jr = b.get("journal") or {}
    doi = next((i.get("id") for i in b.get("identifier") or []
                if str(i.get("type")).lower() == "doi"), "")
    full = next((l.get("url") for l in b.get("link") or []
                 if l.get("type") == "fulltext" and l.get("url")), "")
    ubic = [dict(lugar=jr.get("title") or "DOAJ", url=full, pdf=False, version="publicada")] if full else []
    return _cand(b.get("title"), [familia(a.get("name")) for a in b.get("author") or []],
                 anio(b.get("year")), jr.get("title"), limpiar_doi(doi),
                 f"https://doaj.org/article/{r.get('id')}" if r.get("id") else "", "",
                 True, "gold (revista en DOAJ)", ubic)


def _doaj_buscar(q):
    d = get_json("https://doaj.org/api/search/articles/" + quote(q, safe=""), {"pageSize": 8})
    return [_doaj_cand(r) for r in (d or {}).get("results", [])]


def doaj_doi(doi, ctx):
    return _doaj_buscar(f'bibjson.identifier.id:"{doi}"')


def doaj_titulo(meta, ctx):
    return _doaj_buscar(f'bibjson.title:"{consulta_titulo(meta)}"')


# ---- Semantic Scholar ----
S2_CAMPOS = "title,year,authors,venue,externalIds,isOpenAccess,openAccessPdf"


def _s2_cand(p):
    pdf = (p.get("openAccessPdf") or {}).get("url") or ""
    ubic = [dict(lugar="Semantic Scholar (copia en abierto)", url=pdf, pdf=True, version="")] if pdf else []
    return _cand(p.get("title"), [familia(a.get("name")) for a in p.get("authors") or []],
                 p.get("year"), p.get("venue"), limpiar_doi((p.get("externalIds") or {}).get("DOI")),
                 f"https://www.semanticscholar.org/paper/{p.get('paperId')}" if p.get("paperId") else "",
                 pdf, p.get("isOpenAccess"), "", ubic)


def _s2_hdr(ctx):
    return {"x-api-key": ctx["key_s2"]} if ctx.get("key_s2") else None


def s2_doi(doi, ctx):
    d = get_json(f"https://api.semanticscholar.org/graph/v1/paper/DOI:{quote(doi, safe='/')}",
                 {"fields": S2_CAMPOS}, _s2_hdr(ctx))
    return [_s2_cand(d)] if d and d.get("title") else []


def s2_titulo(meta, ctx):
    d = get_json("https://api.semanticscholar.org/graph/v1/paper/search",
                 {"query": consulta_titulo(meta), "limit": 8, "fields": S2_CAMPOS}, _s2_hdr(ctx))
    return [_s2_cand(p) for p in (d or {}).get("data", [])]


# ---- BASE (requiere IP autorizada por BASE) ----
BASE_API = "https://api.base-search.net/cgi-bin/BaseHttpSearchInterface.fcgi"


def _base_cand(d):
    enlaces = [str(u) for u in lista(d.get("dclink")) if u]
    if not enlaces:
        enlaces = [str(u) for u in lista(d.get("dcidentifier")) if str(u).startswith("http")][:1]
    lugar = (primero(d.get("dcprovider")) or primero(d.get("collname"))
             or primero(d.get("dccollection")) or "BASE")
    ubic = [dict(lugar=lugar, url=u, pdf=_es_pdf(u), version="") for u in enlaces[:3]]
    return _cand(primero(d.get("dctitle")), [familia(x) for x in lista(d.get("dccreator"))],
                 anio(primero(d.get("dcyear")) or primero(d.get("dcdate"))), primero(d.get("dcsource")),
                 limpiar_doi(primero(d.get("dcdoi"))), enlaces[0] if enlaces else "", _primer_pdf(ubic),
                 {"1": True, "2": False}.get(str(primero(d.get("dcoa")))), "", ubic)


def _base_buscar(q):
    d = get_json(BASE_API, {"func": "PerformSearch", "query": q, "format": "json", "hits": 15})
    return [_base_cand(x) for x in ((d or {}).get("response") or {}).get("docs", [])]


def base_doi(doi, ctx):
    return _base_buscar(f'dcdoi:"{doi}"')


def base_titulo(meta, ctx):
    return _base_buscar(f'dctitle:"{consulta_titulo(meta)}"')


FUENTES = {
    "Crossref": dict(clave="crossref", host="api.crossref.org", doi=cr_doi, titulo=cr_titulo),
    "OpenAlex": dict(clave="openalex", host="api.openalex.org", doi=oa_doi, titulo=oa_titulo),
    "Unpaywall": dict(clave="unpaywall", host="api.unpaywall.org", doi=up_doi, titulo=None),
    "CORE": dict(clave="core", host="api.core.ac.uk", doi=core_doi, titulo=core_titulo,
                 requiere_clave="key_core", limitada=True, expandir=core_expandir, localizaciones=True),
    "DOAJ": dict(clave="doaj", host="doaj.org", doi=doaj_doi, titulo=doaj_titulo),
    "Semantic Scholar": dict(clave="semanticscholar", host="api.semanticscholar.org",
                             doi=s2_doi, titulo=s2_titulo, limitada=True),
    "BASE": dict(clave="base", host="api.base-search.net", doi=base_doi, titulo=base_titulo,
                 limitada=True, multi=True, localizaciones=True),
}
LIMITADAS = {n for n, f in FUENTES.items() if f.get("limitada")}
PRIORIDAD = ["Unpaywall", "OpenAlex", "DOAJ", "Semantic Scholar", "CORE", "BASE", "Crossref"]
PRIORIDAD_DOI = ["Crossref", "OpenAlex", "Semantic Scholar", "CORE", "DOAJ", "BASE"]


# =====================================================================
# 4. Coincidencia: título + autor + año + revista/libro
# =====================================================================
PESOS = dict(titulo=0.6, autor=0.2, anio=0.1, fuente=0.1)


def sim_titulo(a, b):
    na, nb = norm_titulo(a), norm_titulo(b)
    if not na or not nb:
        return None
    s = SequenceMatcher(None, na, nb, autojunk=False).ratio()
    if s < 1:
        pa, pb = norm_titulo(titulo_principal(a)), norm_titulo(titulo_principal(b))
        if len(pa.split()) >= 4 and len(pb.split()) >= 4:  # subtítulo presente en uno solo
            s = max(s, 0.97 * SequenceMatcher(None, pa, pb, autojunk=False).ratio())
    return s


def score_autor(autor_excel, apellidos):
    tokens = {t for t in norm(autor_excel).split() if len(t) >= 2}
    cands = [norm(a).split() for a in (apellidos or [])[:5]]
    cands = [c[-1] for c in cands if c]
    if not tokens or not cands:
        return None
    aciertos = [c in tokens for c in cands]
    return max(1.0 if aciertos[0] else 0.0, sum(aciertos) / len(aciertos))


def score_anio(a, b):
    a, b = anio(a), anio(b)
    if a is None or b is None:
        return None
    return {0: 1.0, 1: 0.8, 2: 0.4}.get(abs(a - b), 0.0)


def sim_fuente(a, b):
    na, nb = norm(a), norm(b)
    if not na or not nb:
        return None
    s = SequenceMatcher(None, na, nb, autojunk=False).ratio()
    corto, largo = sorted((na, nb), key=len)
    if len(corto) >= 6 and corto in largo:
        s = max(s, 0.95)
    return s


def evaluar(meta, cand):
    """(puntuación 0-1, detalle) o None si no hay nada que comparar."""
    det = {"titulo": sim_titulo(meta.get("titulo"), cand.get("titulo")) if meta.get("titulo") else None,
           "autor": score_autor(meta.get("autor"), cand.get("autores")),
           "anio": score_anio(meta.get("anio"), cand.get("anio")),
           "fuente": sim_fuente(meta.get("fuente"), cand.get("revista"))}
    usados = {k: v for k, v in det.items() if v is not None}
    if not usados:
        return None
    return sum(PESOS[k] * v for k, v in usados.items()) / sum(PESOS[k] for k in usados), det


def clasificar(ev, umbral, titulo_excel):
    if ev is None:
        return "Sí"  # nada con qué comparar: se acepta
    score, det = ev
    t = det["titulo"]
    if score >= umbral:
        estado = "Sí"
    elif score >= max(umbral - 0.15, 0.5) or (t is not None and t >= 0.95):
        estado = "Dudoso"
    else:
        estado = "No"
    corr = [det[k] for k in ("autor", "anio", "fuente") if det[k] is not None]
    if estado == "Sí" and len(norm_titulo(titulo_excel).split()) <= 3 and (not corr or max(corr) < 0.5):
        estado = "Dudoso"  # título muy corto/genérico sin nada que lo corrobore
    return estado


def consultar_fuente(nombre, modo, meta, ctx, umbral):
    f = FUENTES[nombre]
    if f.get("requiere_clave") and not ctx.get(f["requiere_clave"]):
        return dict(estado="Sin clave", score=None, cand=None, error="")
    try:
        if modo == "doi":
            cands = f["doi"](meta["doi"], ctx)
            cands = [c for c in cands if not c["doi"] or c["doi"] == meta["doi"]]
        else:
            cands = [c for c in f["titulo"](meta, ctx) if c["titulo"]]
    except ErrorAPI as e:
        return dict(estado="Error API", score=None, cand=None, error=str(e))
    except Exception as e:
        return dict(estado="Error API", score=None, cand=None, error=f"{type(e).__name__}: {e}")
    if not cands:
        return dict(estado="No", score=None, cand=None, error="")

    mejor = None
    for c in cands:
        ev = evaluar(meta, c)
        clave = ev[0] if ev else 1.1  # sin datos para comparar = máxima prioridad
        if mejor is None or clave > mejor[0]:
            mejor = (clave, ev, c)
    _, ev, cand = mejor
    estado = clasificar(ev, umbral, meta.get("titulo", ""))
    if modo == "doi" and estado == "No":
        estado = "Dudoso"  # el DOI existe pero los datos del Excel no encajan
    if estado in ("Sí", "Dudoso"):
        if f.get("multi"):
            _fusionar_candidatos(cand, cands, meta, umbral, estado, modo)
        if f.get("expandir"):
            try:
                f["expandir"](cand, meta, ctx)
            except Exception:
                pass  # el detalle de localizaciones es un extra: nunca debe romper la fuente
    return dict(estado=estado, score=ev[0] if ev else None, cand=cand, error="")


def _fusionar_candidatos(cand, cands, meta, umbral, estado, modo):
    """BASE devuelve un registro por repositorio: se reúnen las localizaciones de todos los que son la misma obra."""
    ubic, oa = list(cand["ubicaciones"]), [cand["es_oa"]]
    for c in cands:
        if c is cand:
            continue
        if modo != "doi":
            ev = evaluar(meta, c)
            est = clasificar(ev, umbral, meta.get("titulo", "")) if ev else "Sí"
            if not (est == "Sí" or (estado == "Dudoso" and est == "Dudoso")):
                continue
        ubic += c["ubicaciones"]
        oa.append(c["es_oa"])
    cand["ubicaciones"] = _dedup_ubic(ubic)
    vals = [x for x in oa if x is not None]
    if vals:
        cand["es_oa"] = any(vals)
    cand["pdf"] = cand["pdf"] or _primer_pdf(cand["ubicaciones"])


# =====================================================================
# 5. Procesado de una fila y resumen global
# =====================================================================
def _norm_url(u):
    return re.sub(r"^https?://(www\.)?", "", u.lower()).rstrip("/")


def agregar(resultados):
    def con_estado(est):
        return [(n, resultados[n]["cand"]) for n in PRIORIDAD
                if n in resultados and resultados[n]["estado"] == est and resultados[n]["cand"]]

    buenos, dudosos = con_estado("Sí"), con_estado("Dudoso")
    ubic, vistos = [], set()
    for dudosa, grupo in ((False, buenos), (True, dudosos)):
        for _, c in grupo:
            for u in c["ubicaciones"]:
                k = _norm_url(u["url"])
                if k not in vistos:
                    vistos.add(k)
                    ubic.append(dict(u, dudosa=dudosa))
    rank = {"publicada": 0, "aceptada": 1, "preprint/enviada": 2, "": 3}
    mejor = min(enumerate(ubic), key=lambda iu: (1 if iu[1]["dudosa"] else 0, rank.get(iu[1]["version"], 3),
                                                  0 if iu[1]["pdf"] else 1, iu[0]))[1] if ubic else None

    # Acceso abierto: se usan primero las fuentes fiables; si no dicen nada, las dudosas (marcadas)
    base, suf = (buenos, "") if buenos else (dudosos, " (dudoso)")
    oas = [c["es_oa"] for _, c in base if c["es_oa"] is not None]
    if not oas and buenos and dudosos:
        oas = [c["es_oa"] for _, c in dudosos if c["es_oa"] is not None]
        suf = " (dudoso)" if oas else ""
    if not (buenos or dudosos):
        en_abierto = ""
    else:
        en_abierto = ("Sí" if any(oas) else "No" if oas else "Desconocido") + (suf if oas or not buenos else "")
    estado_oa = next((c["estado_oa"] for _, c in buenos if c["estado_oa"]), "")
    if not estado_oa:
        estado_oa = next((c["estado_oa"] + " (dudoso)" for _, c in dudosos if c["estado_oa"]), "")

    def etiqueta(u):
        return (f"{u['lugar']}" + (f" [{u['version']}]" if u["version"] else "")
                + (" (coincidencia dudosa)" if u["dudosa"] else "") + f": {u['url']}")

    return dict(
        en_abierto=en_abierto, estado_oa=estado_oa,
        versiones_disponibles=", ".join(sorted({u["version"] for u in ubic if u["version"]})),
        mejor_url_acceso_abierto=mejor["url"] if mejor else "",
        donde_esta=" | ".join(etiqueta(u) for u in ubic[:20]),
        fuentes_que_lo_encuentran=", ".join(n for n in FUENTES if resultados.get(n, {}).get("estado") == "Sí"),
        fuentes_con_coincidencia_dudosa=", ".join(n for n in FUENTES if resultados.get(n, {}).get("estado") == "Dudoso"),
        datos_incluyen_dudosas="Sí" if dudosos else ("No" if buenos else ""))


def estado_global(resultados):
    est = [r["estado"] for r in resultados.values()]
    utiles = [e for e in est if e not in ("Sin clave", "Sin DOI", "No consultada")]
    if not utiles:
        return "Sin fuentes disponibles"
    if "Sí" in est:
        return "OK"
    if "Dudoso" in est:
        return "Revisar"
    if "Error API" in est:
        return "Error API" if all(e == "Error API" for e in utiles) \
            else "Revisar (error de API en alguna fuente)"
    return "No encontrado"


def columnas_fuente(clave, modo, detalle, localizaciones=False):
    cols = [f"{clave}_encontrado", f"{clave}_puntuacion"]
    if detalle:
        cols.append(f"{clave}_titulo_encontrado")
        if modo == "titulo":
            cols.append(f"{clave}_doi")
        cols += [f"{clave}_registro_url", f"{clave}_pdf_url", f"{clave}_en_abierto", f"{clave}_estado_oa"]
    if localizaciones:
        cols.append(f"{clave}_localizaciones")
    return cols


def columnas_resultado(modo, fuentes, detalle):
    cols = ["estado_global", "observaciones",
            "doi_usado" if modo == "doi" else "doi_recuperado", "en_abierto", "estado_oa",
            "versiones_disponibles", "mejor_url_acceso_abierto", "donde_esta",
            "fuentes_que_lo_encuentran", "fuentes_con_coincidencia_dudosa", "datos_incluyen_dudosas",
            "busqueda_manual_BASE"]
    for n in fuentes:
        cols += columnas_fuente(FUENTES[n]["clave"], modo, detalle, bool(FUENTES[n].get("localizaciones")))
    return cols


def procesar_fila(meta, modo, fuentes, ctx, umbral, detalle, escalonar=False):
    col_doi = "doi_usado" if modo == "doi" else "doi_recuperado"
    if modo == "doi" and not meta["doi"]:
        return {"estado_global": "DOI no válido" if meta["doi_original"] else "Sin DOI",
                "observaciones": f"DOI no reconocible: {meta['doi_original']}" if meta["doi_original"] else ""}
    if modo == "titulo" and not meta["titulo"]:
        return {"estado_global": "Sin título"}

    resultados = {}
    for n in [f for f in fuentes if f != "Unpaywall" and f not in LIMITADAS]:
        resultados[n] = consultar_fuente(n, modo, meta, ctx, umbral)

    def doi_actual():
        """(doi, 'Sí'|'Dudoso', fuente): primero DOIs de coincidencias fiables, luego de dudosas."""
        if modo == "doi":
            return meta["doi"], "Sí", ""
        for est in ("Sí", "Dudoso"):
            for n in PRIORIDAD_DOI:
                r = resultados.get(n)
                if r and r["estado"] == est and r["cand"] and r["cand"]["doi"]:
                    return r["cand"]["doi"], est, n
        return "", "", ""

    unpaywall_hecho = False
    if "Unpaywall" in fuentes and doi_actual()[0]:
        resultados["Unpaywall"] = consultar_fuente("Unpaywall", "doi", dict(meta, doi=doi_actual()[0]),
                                                   ctx, umbral)
        unpaywall_hecho = True

    # Fuentes con límites estrictos: si ya hay confirmación, se ahorra la consulta
    confirmada = any(r["estado"] == "Sí" for r in resultados.values())
    for n in [f for f in fuentes if f in LIMITADAS]:
        if escalonar and confirmada:
            resultados[n] = dict(estado="No consultada", score=None, cand=None, error="")
        else:
            resultados[n] = consultar_fuente(n, modo, meta, ctx, umbral)

    doi, doi_est, doi_fuente = doi_actual()
    if "Unpaywall" in fuentes and not unpaywall_hecho:
        resultados["Unpaywall"] = (consultar_fuente("Unpaywall", "doi", dict(meta, doi=doi), ctx, umbral)
                                   if doi else dict(estado="Sin DOI", score=None, cand=None, error=""))

    obs = []
    for n, r in resultados.items():
        if r["estado"] == "Error API":
            obs.append(f"{n}: error de API ({r['error'][:100]})")
        elif r["estado"] == "Dudoso":
            obs.append(f"{n}: coincidencia dudosa" + (f" ({r['score']:.2f})" if r["score"] is not None else ""))
        elif r["estado"] == "Sin clave":
            obs.append(f"{n}: sin clave API")
    if modo == "doi" and meta.get("doi_otra_columna"):
        obs.append("DOI extraído de otra columna")
    if modo == "doi" and meta["doi_original"] and not meta["doi"]:
        obs.append("DOI con formato no válido")
    if modo == "titulo" and meta["doi"] and doi and meta["doi"] != doi:
        obs.append(f"El DOI recuperado difiere del DOI del Excel ({meta['doi']})")
    if modo == "titulo" and doi and doi_est == "Dudoso":
        obs.append(f"DOI recuperado de una coincidencia dudosa ({doi_fuente}): verifícalo")
    agg = agregar(resultados)
    if agg["datos_incluyen_dudosas"] == "Sí":
        obs.append("Las columnas resumen incluyen datos de coincidencias dudosas (marcados como dudosos)")

    res = {"estado_global": estado_global(resultados), "observaciones": "; ".join(obs),
           col_doi: doi, **agg,
           "busqueda_manual_BASE": "https://www.base-search.net/Search/Results?lookfor="
                                   + quote_plus(f'"{meta["doi"]}"' if modo == "doi" else f'"{meta["titulo"]}"')}
    for n in fuentes:
        k, r = FUENTES[n]["clave"], resultados[n]
        c = r["cand"] or {}
        mostrar = r["estado"] in ("Sí", "Dudoso")
        res[f"{k}_encontrado"] = r["estado"]
        res[f"{k}_puntuacion"] = round(r["score"], 2) if r["score"] is not None else ""
        if detalle:
            res[f"{k}_titulo_encontrado"] = c.get("titulo", "")
            if modo == "titulo":
                res[f"{k}_doi"] = c.get("doi", "") if mostrar else ""
            res[f"{k}_registro_url"] = c.get("registro", "") if mostrar else ""
            res[f"{k}_pdf_url"] = c.get("pdf", "") if mostrar else ""
            res[f"{k}_en_abierto"] = si_no(c.get("es_oa")) if mostrar else ""
            res[f"{k}_estado_oa"] = c.get("estado_oa", "") if mostrar else ""
        if FUENTES[n].get("localizaciones"):
            res[f"{k}_localizaciones"] = fmt_ubic(c.get("ubicaciones", [])) if mostrar else ""
    return res


# =====================================================================
# 6. Lectura del Excel de entrada
# =====================================================================
ALIAS = {"titulo": ["titulo", "title", "article title", "document title"],
         "doi": ["doi"],
         "autor": ["author", "authors", "autor", "autores", "author s"],
         "fuente": ["source", "source title", "revista", "journal", "publication"],
         "anio": ["pubyear", "year", "anio", "ano", "publication year", "publication date", "fecha", "date"]}


def _col(df, clave):
    mapa = {norm(c): c for c in df.columns}
    return next((mapa[a] for a in ALIAS[clave] if a in mapa), None)


def _val(fila, col):
    if col is None:
        return ""
    v = fila[col]
    try:
        if pd.isna(v):
            return ""
    except (TypeError, ValueError):
        pass
    return str(v).strip()


def leer_entrada(ruta, modo):
    df = pd.read_excel(ruta)
    cols = {k: _col(df, k) for k in ALIAS}
    if modo == "titulo" and not cols["titulo"]:
        raise ValueError("No encuentro la columna 'título' (títulos admitidos: título, titulo, title).")
    metas = []
    for _, f in df.iterrows():
        doi_orig = _val(f, cols["doi"])
        doi = extraer_doi(doi_orig)
        otra = False
        if not doi and not doi_orig:
            for v in f.values:
                d = extraer_doi(v)
                if d:
                    doi, otra = d, True
                    break
        metas.append(dict(titulo=_val(f, cols["titulo"]), doi=doi, doi_original=doi_orig,
                          doi_otra_columna=otra, autor=_val(f, cols["autor"]),
                          fuente=_val(f, cols["fuente"]), anio=anio(_val(f, cols["anio"]))))
    if modo == "doi" and not any(m["doi"] or m["doi_original"] for m in metas):
        raise ValueError("No encuentro ningún DOI (ni en la columna 'doi' ni en otras columnas).")
    return df, metas, cols


# =====================================================================
# 7. Excel de salida con formato
# =====================================================================
_ILEGAL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def _celda(v):
    return _ILEGAL.sub("", v)[:32000] if isinstance(v, str) else v


def guardar_excel(df, ruta):
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    df = df.copy()
    for c in df.columns:
        df[c] = df[c].map(_celda)
    eg = df["estado_global"].astype(str) if "estado_global" in df else pd.Series([""] * len(df))
    hojas = {
        "Resultados": df,
        "En abierto": df[df["en_abierto"].astype(str).str.startswith("Sí")] if "en_abierto" in df else df.iloc[0:0],
        "Revisar": df[eg.str.startswith(("Revisar", "Error", "Sin", "DOI no válido"))],
        "No encontrados": df[eg == "No encontrado"],
    }
    relleno = {"OK": "C6EFCE", "Sí": "C6EFCE", "Revisar": "FFEB9C", "Dudoso": "FFEB9C",
               "No encontrado": "FFC7CE", "Error API": "F8CBAD"}
    with pd.ExcelWriter(ruta, engine="openpyxl") as w:
        for nombre, d in hojas.items():
            d.to_excel(w, sheet_name=nombre, index=False)
            ws = w.sheets[nombre]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for cel in ws[1]:
                cel.font = Font(bold=True, color="FFFFFF")
                cel.fill = PatternFill("solid", fgColor="305496")
                cel.alignment = Alignment(wrap_text=True, vertical="center")
            for i, col in enumerate(ws.columns, 1):
                largo = max((len(str(c.value)) for c in list(col)[:200] if c.value is not None), default=8)
                ws.column_dimensions[get_column_letter(i)].width = min(largo + 2, 60)
            cabeceras = [c.value for c in ws[1]]
            colorear = {i for i, h in enumerate(cabeceras)
                        if h == "estado_global" or str(h).endswith("_encontrado")}
            enlaces = 0
            for fila in ws.iter_rows(min_row=2):
                for i, cel in enumerate(fila):
                    v = cel.value
                    if i in colorear and isinstance(v, str):
                        clave = "Revisar" if v.startswith("Revisar") else v
                        if clave in relleno:
                            cel.fill = PatternFill("solid", fgColor=relleno[clave])
                    elif (isinstance(v, str) and v.startswith(("http://", "https://"))
                          and " " not in v and len(v) < 2000 and enlaces < 60000):
                        cel.hyperlink = v
                        cel.font = Font(color="0563C1", underline="single")
                        enlaces += 1


# =====================================================================
# 8. Proceso completo (usado por la interfaz)
# =====================================================================
def ejecutar_proceso(modo, ruta_entrada, email, fuentes, umbral=0.85, workers=3, usar_cache=True,
                     detalle=True, key_core="", key_oa="", key_s2="", escalonar=True,
                     core_detalle=True, progreso=lambda f, d: None):
    if not email or "@" not in email:
        raise ValueError("Indica un email válido (lo requieren las APIs).")
    fuentes = [n for n in FUENTES if n in (fuentes or [])]
    if not fuentes:
        raise ValueError("Selecciona al menos una fuente.")
    if modo == "titulo" and fuentes == ["Unpaywall"]:
        raise ValueError("Unpaywall solo funciona con DOI; en modo título elige otra fuente además.")

    df, metas, _ = leer_entrada(ruta_entrada, modo)
    ctx = dict(email=email.strip(), key_core=(key_core or "").strip(),
               key_oa=(key_oa or "").strip(), key_s2=(key_s2 or "").strip(),
               core_detalle=bool(core_detalle))
    configurar_cache(usar_cache)
    reiniciar_stats()
    # Semantic Scholar: con clave 1 consulta/s; sin clave comparten cupo global y conviene ir más lento
    INTERVALOS["api.semanticscholar.org"] = 1.1 if ctx["key_s2"] else 3.0

    n = len(metas)
    resultados = [None] * n

    def tarea(meta):
        return procesar_fila(meta, modo, fuentes, ctx, umbral, detalle, escalonar)

    with ThreadPoolExecutor(max_workers=max(1, int(workers))) as ex:
        futs = {ex.submit(tarea, m): i for i, m in enumerate(metas)}
        hechos = 0
        for fut in as_completed(futs):
            i = futs[fut]
            try:
                resultados[i] = fut.result()
            except Exception as e:
                resultados[i] = {"estado_global": "Error API", "observaciones": f"Error inesperado: {e}"}
            hechos += 1
            progreso(hechos / max(n, 1), f"Consultando {hechos}/{n}")

    cols = columnas_resultado(modo, fuentes, detalle)
    res = pd.DataFrame(resultados).reindex(columns=cols).fillna("")
    res.columns = [f"{c}_resultado" if c in df.columns else c for c in res.columns]
    salida = pd.concat([df.reset_index(drop=True), res], axis=1)

    import tempfile
    ruta = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False,
                                       prefix=f"resultado_{modo.upper()}_").name
    guardar_excel(salida, ruta)

    eg = res["estado_global"] if "estado_global" in res else res.iloc[:, 0]
    cuenta = lambda f: int(eg.map(f).sum())
    resumen = (f"{n} filas: {cuenta(lambda e: e == 'OK')} OK, "
               f"{cuenta(lambda e: str(e).startswith(('Revisar', 'Error', 'Sin', 'DOI no')))} a revisar, "
               f"{cuenta(lambda e: e == 'No encontrado')} no encontradas. "
               f"En abierto: {int(res['en_abierto'].astype(str).str.startswith('Sí').sum()) if 'en_abierto' in res else 0}. "
               f"Consultas a API: {STATS['api']}; resueltas desde caché: {STATS['cache']}.")
    if "CORE" in fuentes and not ctx["key_core"]:
        resumen += " CORE omitido: falta la clave API."
    nombre_host = {FUENTES[n]["host"]: n for n in FUENTES}
    susp = [f"{nombre_host.get(h, h)} ({m})" for h, m in hosts_suspendidos()]
    if susp:
        resumen += (" FUENTES SUSPENDIDAS por fallos repetidos: " + "; ".join(susp)
                    + ". Sus filas quedaron como 'Error API': relanza más tarde el mismo Excel "
                      "(la caché conserva lo ya consultado).")
    return salida, ruta, resumen


# =====================================================================
# 9. Ejecución conjunta (DOI + título) y textos de ayuda
# =====================================================================
def ejecutar_ambos(ruta_entrada, email, hacer_doi, hacer_titulo, fuentes, umbral=0.85, workers=3,
                   usar_cache=True, detalle=True, key_core="", key_oa="", key_s2="", escalonar=True,
                   core_detalle=True, progreso=lambda f, d: None):
    """Ejecuta las búsquedas pedidas. Devuelve {'resumen', 'doi': (df, ruta)|None, 'titulo': ...}."""
    import shutil
    import tempfile

    if not (hacer_doi or hacer_titulo):
        raise ValueError("Marca al menos una búsqueda (por DOI y/o por título).")
    if not email or "@" not in email:
        raise ValueError("Indica un email válido (lo requieren las APIs).")
    if not [n for n in FUENTES if n in (fuentes or [])]:
        raise ValueError("Selecciona al menos una fuente.")

    modos = [m for m, h in (("doi", hacer_doi), ("titulo", hacer_titulo)) if h]
    nombres = {"doi": "DOI", "titulo": "título"}
    salida, lineas = {"doi": None, "titulo": None}, []
    for idx, modo in enumerate(modos):
        def prog(f, d, idx=idx, modo=modo):
            progreso((idx + f) / len(modos), f"Búsqueda por {nombres[modo]}: {d}")
        try:
            df, ruta_tmp, resumen = ejecutar_proceso(
                modo, ruta_entrada, email, fuentes, umbral, workers, usar_cache, detalle,
                key_core, key_oa, key_s2, escalonar, core_detalle, prog)
        except ValueError as e:
            lineas.append(f"Búsqueda por {nombres[modo]} OMITIDA: {e}")
            continue
        destino = os.path.join(tempfile.mkdtemp(), f"resultado_busqueda_por_{nombres[modo].replace('í', 'i')}.xlsx")
        shutil.copy(ruta_tmp, destino)
        salida[modo] = (df, destino)
        lineas.append(f"Por {nombres[modo]}: {resumen}")
    if not any(salida.values()):
        raise ValueError(" | ".join(lineas))
    salida["resumen"] = "\n".join(lineas)
    return salida


AYUDA_FUNCIONAMIENTO = """
### Qué hace este programa
Para cada publicación de tu Excel comprueba **si existe, si está en acceso abierto y dónde se puede leer**,
consultando varias bases de datos. Hace **dos búsquedas independientes** y te devuelve **dos Excel**:

1. **Búsqueda por DOI**: consulta cada fuente con el DOI exacto. Es la más fiable.
2. **Búsqueda por título**: busca por título y elige el mejor candidato de cada fuente. Sirve sobre todo
   para publicaciones **sin DOI**, y también para recuperar el DOI que falta.

**Cómo decide que una coincidencia es buena (búsqueda por título).** Compara lo que devuelve cada fuente
con tu fila y calcula una *puntuación* de 0 a 1: título (60 %), autor (20 %), año (10 %) y revista/libro (10 %).
Solo cuentan los datos que tengas en el Excel. Ignora tildes, mayúsculas, etiquetas HTML y subtítulos.
- Puntuación ≥ umbral → **Sí**.
- Algo por debajo del umbral (hasta 0,15 menos), o título casi idéntico con otros datos distintos → **Dudoso** (revísalo a mano). Los datos de las coincidencias dudosas **también se recogen** en todas las columnas, pero **marcados como dudosos**.
- Resto → **No**. En ese caso solo se muestra el título del mejor candidato y su puntuación, para que puedas revisarlo.

**Verificación en la búsqueda por DOI.** Si tu Excel también trae título, autor, año o revista, se comparan con
lo que devuelve la fuente. Si el DOI existe pero los datos no encajan, sale **Dudoso**: probablemente el DOI
del Excel es erróneo. Si el DOI está vacío, se intenta extraer de otras columnas (por ejemplo una URL `doi.org`).

**Fuentes disponibles**

| Fuente | Qué aporta | Requisitos |
|---|---|---|
| Crossref | Confirma que el DOI/obra existe; editorial y revista | Ninguno |
| OpenAlex | Si está en abierto, tipo de acceso y ubicaciones (repositorios, editorial) | Ninguno (clave opcional) |
| Unpaywall | Acceso abierto por DOI: mejor URL y copias. **Solo trabaja con DOI**: en la búsqueda por título usa el DOI recuperado | Ninguno |
| CORE | **Cada repositorio donde está la obra con su URL**, y enlaces al registro y al PDF | **Clave API gratuita**; cuota baja (ver «Límites») |
| DOAJ | Artículos de revistas en acceso abierto | Ninguno |
| Semantic Scholar | Metadatos y PDF en abierto si lo hay | Sin clave comparte un límite global y suele dar error 429; con clave gratuita, 1 consulta/s |
| BASE | **Cada repositorio donde está la obra con su URL** (reúne todos los registros que son la misma obra) | **Tu IP debe estar autorizada por BASE** |

**Otros ajustes**
- **Filas en paralelo**: acelera el proceso; con valores altos las APIs pueden limitar más.
- **Caché**: guarda las respuestas en `cache_api.sqlite`. Si se interrumpe, vuelve a lanzar el mismo Excel y solo
  se repetirán las consultas pendientes. Los errores no se guardan en caché.
- **CORE: detallar cada repositorio**: para cada obra encontrada, CORE se consulta una vez por cada copia en un
  repositorio (hasta 5) y una vez por cada nombre de repositorio desconocido (se guardan en caché). Da la lista
  completa `repositorio: URL`, pero gasta más cuota. Desmárcalo si te quedas sin cupo.
- **Detalle por fuente**: si lo desmarcas, el Excel solo incluye el resultado y la puntuación de cada fuente.
- Tu **email** se envía a las APIs como identificación de cortesía (lo piden Crossref, OpenAlex y Unpaywall).

### Límites de las APIs y errores 429 / 500
- **429 = demasiadas peticiones.** El programa espera, reintenta y espacia más las siguientes consultas a ese servidor.
- **500 = fallo del servidor** (a veces transitorio o por una consulta que no le gusta). En CORE se reintenta
  con una consulta más simple antes de darlo por fallido.
- **Suspensión automática:** si una fuente falla 5 veces seguidas, se **suspende 10 minutos** para no gastar
  tiempo ni cuota. Sus filas salen como `Error API` y el resumen te lo indica.
- **Qué hacer:** relanza el mismo Excel pasado un rato. La caché conserva todo lo ya consultado, así que solo se
  repite lo que falló.
- **CORE** tiene cuotas bajas: según su documentación pública, del orden de 10-25 consultas por minuto y
  1.000-5.000 al día según el tipo de cuenta (confírmalo en tu panel de CORE). Por eso el programa va a
  una consulta cada 6 s.
- **Semantic Scholar** sin clave comparte un cupo con todos los usuarios y limita sobre todo la búsqueda por
  título. Lo más fiable es pedir una clave gratuita (1 consulta/s), o dejarlo desmarcado.
- **Casilla «solo si las demás no confirman»** (activada por defecto): CORE, Semantic Scholar y BASE solo se
  consultan para las filas que Crossref, OpenAlex, Unpaywall o DOAJ no han confirmado. Reduce muchísimo las
  consultas a las fuentes limitadas. Desmárcala si quieres sus datos para **todas** las filas.
"""

AYUDA_EXCEL_ENTRADA = """
### Cómo preparar el Excel de entrada
Una fila por publicación, en la **primera hoja**. Se reconocen estas columnas (da igual mayúsculas o tildes):

| Dato | Nombres admitidos | ¿Obligatoria? |
|---|---|---|
| Título | `título`, `titulo`, `title` | Para la búsqueda por título |
| DOI | `doi` (admite `10.xxxx/...`, `doi:...` o URL de doi.org) | Para la búsqueda por DOI |
| Autor | `Author`, `Authors`, `Autor`, `Autores` | No, pero mejora la coincidencia |
| Revista o libro | `Source`, `Source Title`, `Revista`, `Journal` | No, pero mejora la coincidencia |
| Año | `PubYear`, `Year`, `Año`, `Publication Year` | No, pero mejora la coincidencia |

- Si falta la columna de título, la búsqueda por título se **omite** (y se avisa); lo mismo con el DOI.
- Se **conservan todas tus columnas originales**; los resultados se añaden a la derecha.
- En `Author` basta con los apellidos en cualquier formato (`Smith, J.; Doe, A.` o `John Smith and Ann Doe`).
"""

AYUDA_EXCEL_SALIDA = """
### Cómo leer los Excel resultantes
Cada Excel tiene **cuatro hojas**: `Resultados` (todo), `En abierto`, `Revisar` y `No encontrados`.
Tiene cabecera fija, filtros y enlaces clicables. Tus columnas originales van primero; después estas:

**Columnas resumen (una por fila)**

| Columna | Significado |
|---|---|
| `estado_global` | **OK**: al menos una fuente lo confirma. **Revisar**: hay coincidencias dudosas. **Revisar (error de API…)**: ninguna fuente lo confirma y alguna falló, así que el "no" no es fiable. **No encontrado**: ninguna fuente lo tiene. **Error API**: todas fallaron. **Sin DOI / Sin título / DOI no válido**: faltaba el dato |
| `observaciones` | Explica los avisos: fuente dudosa (con su puntuación), errores de API, DOI recuperado distinto del de tu Excel, DOI extraído de otra columna… |
| `doi_usado` / `doi_recuperado` | DOI consultado (búsqueda por DOI) o DOI hallado a partir del título (búsqueda por título). **Verifícalo antes de copiarlo a tu Excel** |
| `en_abierto` | **Sí** / **No** / **Desconocido** según las fuentes fiables. Si solo hay coincidencias dudosas, se indica con el sufijo **(dudoso)**, p. ej. `Sí (dudoso)` |
| `estado_oa` | Tipo de acceso abierto: *gold* (revista abierta), *hybrid* (artículo abierto en revista de suscripción), *green* (copia en repositorio), *bronze* (gratis en la web de la editorial sin licencia abierta), *closed* (cerrado) |
| `versiones_disponibles` | Versiones con copia en abierto: publicada, aceptada, preprint/enviada |
| `mejor_url_acceso_abierto` | Enlace recomendado: prefiere la versión publicada y el PDF directo |
| `donde_esta` | Todas las ubicaciones encontradas (lugar [versión]: enlace), sin duplicados y hasta 20. Las que vienen de coincidencias dudosas llevan la marca `(coincidencia dudosa)` y van al final |
| `fuentes_que_lo_encuentran` | Fuentes que dan coincidencia **Sí** |
| `fuentes_con_coincidencia_dudosa` | Fuentes cuya coincidencia es **Dudoso** (sus datos también están en las columnas resumen) |
| `datos_incluyen_dudosas` | **Sí** si las columnas resumen incluyen datos de alguna coincidencia dudosa; **No** si todo viene de coincidencias fiables |
| `busqueda_manual_BASE` | Enlace para repetir la búsqueda a mano en BASE |

**Columnas por fuente** (prefijo `crossref_`, `openalex_`, `unpaywall_`, `core_`, `doaj_`, `semanticscholar_`, `base_`)

| Columna | Significado |
|---|---|
| `…_encontrado` | **Sí** / **Dudoso** / **No** / **Error API** (la fuente falló, repite más tarde) / **Sin clave** / **Sin DOI** (Unpaywall sin DOI recuperado) / **No consultada** (fuente limitada que no hizo falta consultar) |
| `…_puntuacion` | Coincidencia de 0 a 1 con los datos de tu Excel (vacía si no había nada que comparar) |
| `…_titulo_encontrado` | Título del mejor candidato; útil para revisar los dudosos y los "No" |
| `…_doi` | DOI del candidato (solo búsqueda por título) |
| `…_registro_url` / `…_pdf_url` | Ficha del registro y PDF directo (CORE: `core.ac.uk/works/…` y `core.ac.uk/download/…`) |
| `…_en_abierto` / `…_estado_oa` | Acceso abierto según esa fuente |
| `core_localizaciones` / `base_localizaciones` | **Todos los repositorios donde está la publicación**, con su URL: `Repositorio: URL \\| Repositorio: URL …`. Se rellenan también con coincidencias dudosas |

**Colores:** verde = OK/Sí · amarillo = Revisar/Dudoso · rojo = No encontrado · naranja = Error API.

**Consejos de interpretación**
- En la búsqueda **por título**, un "No encontrado" **no prueba** que la obra no exista: puede estar con otro título o idioma.
- Los datos de las coincidencias **dudosas** se recogen en todas las columnas para que no pierdas información, pero **pueden pertenecer a otra obra**: compruébalos antes de usarlos (mira `observaciones`, `datos_incluyen_dudosas` y la puntuación de cada fuente).
- `mejor_url_acceso_abierto` prefiere siempre las ubicaciones de coincidencias fiables sobre las dudosas.
- En CORE, a veces el nombre del repositorio no está disponible y aparece como `Repositorio CORE #<número>`.
- Si hay muchos "Error API", espera unos minutos y relanza el mismo Excel: la caché evita repetir lo ya consultado.
"""
