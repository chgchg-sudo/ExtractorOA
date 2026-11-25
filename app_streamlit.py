"""
Interfaz Streamlit del buscador de publicaciones.

Local:   pip install -r requirements.txt  &&  streamlit run app_streamlit.py
Nube:    Streamlit Community Cloud (archivo principal: app_streamlit.py). Ver README.md
"""
import hmac
import os
import tempfile
import threading

# La caché debe apuntar a una ruta escribible; se fija ANTES de importar el núcleo.
os.environ.setdefault("BUSCADOR_CACHE", os.path.join(tempfile.gettempdir(), "buscador_cache_api.sqlite"))

import pandas as pd
import streamlit as st

from buscador_core import (AYUDA_EXCEL_ENTRADA, AYUDA_EXCEL_SALIDA, AYUDA_FUNCIONAMIENTO,
                           FUENTES, ejecutar_ambos, vaciar_cache)

st.set_page_config(page_title="Buscador de publicaciones", page_icon="📚", layout="wide")
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def secreto(nombre, defecto=""):
    """Lee un secreto de Streamlit (o devuelve el valor por defecto si no hay)."""
    try:
        return st.secrets.get(nombre, defecto)
    except Exception:
        return defecto


@st.cache_resource
def candado():
    """Un único candado para todo el servidor: el núcleo comparte estado entre sesiones
    (límites de velocidad, suspensión de fuentes), así que solo se permite una consulta a la vez."""
    return threading.Lock()


# ---------- Contraseña opcional (secreto APP_PASSWORD) ----------
clave_app = str(secreto("APP_PASSWORD"))
if clave_app and not st.session_state.get("autenticado"):
    st.title("Buscador de publicaciones")
    pw = st.text_input("Contraseña de acceso", type="password")
    if pw:
        if hmac.compare_digest(pw, clave_app):
            st.session_state["autenticado"] = True
            st.rerun()
        else:
            st.error("Contraseña incorrecta.")
    st.stop()

# ---------- Barra lateral: configuración ----------
with st.sidebar:
    st.header("Configuración")
    email = st.text_input("Tu email (obligatorio)", placeholder="nombre@institucion.es",
                          help="Se envía a Crossref, OpenAlex y Unpaywall como identificación de cortesía.")
    fuentes = st.multiselect("Fuentes", list(FUENTES), default=["Crossref", "OpenAlex", "Unpaywall"],
                             help="BASE exige que la IP del servidor esté autorizada por BASE: "
                                  "en la nube normalmente no funcionará. CORE necesita clave API.")
    umbral = st.slider("Umbral de coincidencia", 0.5, 1.0, 0.85, 0.01,
                       help="Puntuación mínima (título + autor + año + revista) para dar una coincidencia por buena.")
    workers = st.slider("Filas en paralelo", 1, 5, 3)
    usar_cache = st.checkbox("Usar caché (permite reanudar)", True)
    detalle = st.checkbox("Incluir detalle por fuente", True)
    escalonar = st.checkbox("Consultar CORE, Semantic Scholar y BASE solo si las demás no confirman", True,
                            help="Estas fuentes tienen límites de uso estrictos.")
    core_detalle = st.checkbox("CORE: detallar cada repositorio", True,
                               help="Una consulta extra por repositorio: gasta más cuota de CORE.")
    with st.expander("Claves API (opcionales salvo CORE)"):
        key_core = st.text_input("CORE", type="password")
        key_oa = st.text_input("OpenAlex", type="password")
        key_s2 = st.text_input("Semantic Scholar", type="password")
        if any(secreto(k) for k in ("CORE_API_KEY", "OPENALEX_API_KEY", "SEMANTIC_SCHOLAR_API_KEY")):
            st.caption("Los campos vacíos usan las claves configuradas en el servidor.")
    if st.button("Vaciar caché"):
        st.success(f"Caché vaciada ({vaciar_cache()} entradas).")

# ---------- Cuerpo ----------
st.title("Buscador de publicaciones: por DOI y por título")
st.write("Comprueba si cada publicación de tu Excel existe, si está en acceso abierto y dónde. "
         "Devuelve **dos Excel**: uno con la búsqueda por DOI y otro con la búsqueda por título.")

tab_buscar, tab_ayuda = st.tabs(["Buscar", "Instrucciones"])

with tab_ayuda:
    t1, t2, t3 = st.tabs(["Cómo funciona", "Preparar el Excel", "Leer el resultado"])
    with t1:
        st.markdown(AYUDA_FUNCIONAMIENTO)
    with t2:
        st.markdown(AYUDA_EXCEL_ENTRADA)
    with t3:
        st.markdown(AYUDA_EXCEL_SALIDA)

with tab_buscar:
    archivo = st.file_uploader("Excel de entrada (.xlsx)", type=["xlsx", "xls"])
    c1, c2 = st.columns(2)
    hacer_doi = c1.checkbox("Búsqueda por DOI", True)
    hacer_titulo = c2.checkbox("Búsqueda por título", True)
    lanzar = st.button("Consultar", type="primary", disabled=archivo is None)

    if lanzar:
        # Los valores vacíos se completan con los secretos del servidor (nunca se envían al navegador)
        email_ef = email.strip() or str(secreto("CONTACT_EMAIL"))
        k_core = key_core.strip() or str(secreto("CORE_API_KEY"))
        k_oa = key_oa.strip() or str(secreto("OPENALEX_API_KEY"))
        k_s2 = key_s2.strip() or str(secreto("SEMANTIC_SCHOLAR_API_KEY"))
        try:
            max_filas = int(secreto("MAX_FILAS", 500))
        except (TypeError, ValueError):
            max_filas = 500

        # Se guarda el Excel subido en un archivo temporal (se lee una vez por tipo de búsqueda)
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            tmp.write(archivo.getvalue())
            ruta = tmp.name
        try:
            try:
                n_filas = len(pd.read_excel(ruta))
            except Exception as e:
                st.error(f"No he podido leer el Excel: {e}")
                st.stop()
            if n_filas > max_filas:
                st.error(f"El Excel tiene {n_filas} filas y el máximo en esta aplicación es {max_filas}. "
                         "Divídelo en varios archivos.")
                st.stop()

            cerrojo = candado()
            if not cerrojo.acquire(blocking=False):
                st.warning("Hay otra consulta en curso. Espera a que termine e inténtalo de nuevo.")
                st.stop()
            barra = st.progress(0.0, text="Preparando…")
            try:
                def progreso(frac, desc):
                    barra.progress(min(max(float(frac), 0.0), 1.0), text=desc)

                r = ejecutar_ambos(ruta, email_ef, hacer_doi, hacer_titulo, fuentes, umbral, workers,
                                   usar_cache, detalle, k_core, k_oa, k_s2, escalonar, core_detalle,
                                   progreso)
                almacen = {"resumen": r["resumen"]}
                for modo in ("doi", "titulo"):
                    if r[modo]:
                        df, ruta_xlsx = r[modo]
                        with open(ruta_xlsx, "rb") as f:
                            almacen[modo] = (df.head(200).fillna("").astype(str), f.read(), len(df))
                st.session_state["resultado"] = almacen
            except ValueError as e:
                st.error(str(e))
            except Exception as e:
                st.error(f"Error inesperado: {type(e).__name__}: {e}")
            finally:
                cerrojo.release()
                barra.empty()
        finally:
            try:
                os.unlink(ruta)
            except OSError:
                pass

    # Los resultados se guardan en la sesión para que sobrevivan a los clics en «Descargar»
    res = st.session_state.get("resultado")
    if res:
        st.info(res["resumen"])
        tabs = st.tabs(["Resultado por DOI", "Resultado por título"])
        for tab, modo, etiqueta, nombre in ((tabs[0], "doi", "DOI", "DOI"),
                                            (tabs[1], "titulo", "título", "titulo")):
            with tab:
                if modo not in res:
                    st.caption(f"No se ha ejecutado la búsqueda por {etiqueta}.")
                    continue
                vista, datos, n = res[modo]
                st.download_button(f"Descargar Excel (búsqueda por {etiqueta})", datos,
                                   file_name=f"resultado_busqueda_por_{nombre}.xlsx", mime=XLSX,
                                   key=f"descarga_{modo}")
                st.caption(f"Vista previa: primeras {min(n, 200)} de {n} filas.")
                st.dataframe(vista, use_container_width=True)
