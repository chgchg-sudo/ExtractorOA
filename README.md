# Buscador de publicaciones: por DOI y por título

Comprueba, a partir de un Excel, si cada publicación **existe**, si está en **acceso abierto** y **en qué repositorios** se puede leer. Hace dos búsquedas independientes (por DOI y por título) y devuelve **dos Excel** con formato.

Fuentes: Crossref, OpenAlex, Unpaywall, CORE, DOAJ, Semantic Scholar y BASE.

## Qué hace

- **Búsqueda por DOI**: consulta cada fuente con el DOI exacto y, si tu Excel trae título, autor, año o revista, verifica que el DOI corresponde a la obra correcta.
- **Búsqueda por título**: elige el mejor candidato de cada fuente según título, autor, año y revista, y recupera el DOI que falta.
- **Localizaciones**: lista cada repositorio donde está la obra con su URL (CORE y BASE), versiones disponibles (publicada, aceptada, preprint) y la mejor URL de acceso abierto.
- **Coincidencias dudosas**: se recogen igualmente, pero marcadas como dudosas.
- **Robustez**: reintentos, límites de velocidad por servidor, suspensión automática de fuentes que fallan, y caché SQLite que permite reanudar un proceso interrumpido.

## Formato del Excel de entrada

Una fila por publicación, en la primera hoja. Ver `ejemplo_entrada.xlsx`.

| Dato | Nombres admitidos |
|---|---|
| Título | `título`, `titulo`, `title` |
| DOI | `doi` (también `doi:...` o URL de doi.org) |
| Autor | `Author`, `Authors`, `Autor`, `Autores` |
| Revista o libro | `Source`, `Source Title`, `Revista`, `Journal` |
| Año | `PubYear`, `Year`, `Año`, `Publication Year` |

Basta con DOI o título; el resto mejora la coincidencia. Las instrucciones completas (y el significado de cada columna del resultado) están dentro de la aplicación, en la pestaña **Instrucciones**.

## Estructura del proyecto

```
buscador_core.py        # lógica (sin interfaz)
app_streamlit.py        # interfaz Streamlit
app_gradio.py           # interfaz Gradio (alternativa)
requirements.txt        # dependencias de la versión Streamlit
requirements-gradio.txt # dependencias de la versión Gradio
ejemplo_entrada.xlsx    # Excel de ejemplo
.streamlit/             # configuración y plantilla de secretos
```

## Uso en local

```bash
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt
streamlit run app_streamlit.py
```

Versión Gradio:

```bash
pip install -r requirements-gradio.txt
python app_gradio.py
```

### Secretos en local (opcional)

Copia `.streamlit/secrets.toml.example` como `.streamlit/secrets.toml` y rellénalo. Está en `.gitignore`: **no lo subas nunca a GitHub**.

## Publicar en GitHub

```bash
git init
git add .
git commit -m "Primera versión"
git branch -M main
git remote add origin https://github.com/<tu-usuario>/<tu-repositorio>.git
git push -u origin main
```

Antes de subir, comprueba con `git status` que **no** aparece `.streamlit/secrets.toml` ni ningún `.sqlite`. Añade además una licencia (por ejemplo desde <https://choosealicense.com>); este repositorio no incluye ninguna.

## Desplegar en Streamlit Community Cloud

1. Sube el repositorio a GitHub (arriba).
2. Entra en <https://share.streamlit.io> con tu cuenta de GitHub y crea una app nueva.
3. Elige el repositorio, la rama `main` y el archivo principal **`app_streamlit.py`**.
4. En la configuración avanzada, pega tus secretos en formato TOML (contenido de `secrets.toml.example`). Es opcional.
5. Despliega.

### Qué tener en cuenta si la app es pública

- **Cuota de tus claves**: si pones claves de API en los secretos, cualquier visitante las usará indirectamente. Pon `APP_PASSWORD`, o no pongas claves y deja que cada usuario use las suyas (los campos vacíos usan las del servidor; los rellenados, las del usuario).
- **BASE no funcionará**: exige que BASE autorice la IP del servidor, y la nube no ofrece una IP fija.
- **Una consulta a la vez**: el estado de límites y suspensiones se comparte entre sesiones, así que la app rechaza una segunda consulta mientras otra está en curso.
- **Límite de filas**: por defecto 500 por Excel (`MAX_FILAS` en secretos). Un proceso largo puede interrumpirse si cierras la pestaña; al relanzar, la caché evita repetir lo ya consultado.
- **Caché efímera**: el disco de la nube se reinicia cuando la app se reinicia o duerme.
- **Privacidad**: los Excel subidos se procesan en el servidor de Streamlit. No subas datos sensibles.
- **Condiciones de las APIs**: respeta los límites y condiciones de uso de cada servicio (Crossref, OpenAlex, Unpaywall, CORE, DOAJ, Semantic Scholar, BASE).

## Límites de las APIs

- **Semantic Scholar** sin clave comparte un cupo global y suele dar error 429. Con clave gratuita permite una consulta por segundo.
- **CORE** exige clave gratuita y tiene cuotas bajas; el programa va a una consulta cada 6 s.
- Por eso, por defecto CORE, Semantic Scholar y BASE solo se consultan cuando las demás fuentes no confirman la publicación.
