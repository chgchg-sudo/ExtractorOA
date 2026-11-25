"""
Interfaz Gradio del buscador de publicaciones.
Uso:  pip install -r requirements-gradio.txt  &&  python app_gradio.py
"""
from buscador_core import (AYUDA_EXCEL_ENTRADA, AYUDA_EXCEL_SALIDA, AYUDA_FUNCIONAMIENTO,
                           FUENTES, ejecutar_ambos, vaciar_cache)


def crear_app():
    import gradio as gr

    def ejecutar(archivo, email, hacer_doi, hacer_titulo, fuentes, umbral, workers, usar_cache,
                 detalle, escalonar, core_detalle, key_core, key_oa, key_s2, progress=gr.Progress()):
        if archivo is None:
            raise gr.Error("Sube un archivo Excel.")
        ruta = archivo if isinstance(archivo, str) else getattr(archivo, "name", archivo)
        try:
            r = ejecutar_ambos(ruta, email, hacer_doi, hacer_titulo, fuentes, umbral, workers,
                               usar_cache, detalle, key_core, key_oa, key_s2, escalonar,
                               core_detalle, lambda f, d: progress(f, desc=d))
        except ValueError as e:
            raise gr.Error(str(e))
        d, t = r["doi"], r["titulo"]
        return (r["resumen"],
                d[0].head(200) if d else None, d[1] if d else None,
                t[0].head(200) if t else None, t[1] if t else None)

    def limpiar():
        return f"Caché vaciada ({vaciar_cache()} entradas)."

    with gr.Blocks(title="Buscador de publicaciones") as demo:
        gr.Markdown("# Buscador de publicaciones: por DOI y por título\n"
                    "Comprueba si cada publicación de tu Excel existe, si está en acceso abierto y dónde. "
                    "Devuelve **dos Excel**: uno con la búsqueda por DOI y otro con la búsqueda por título.")
        with gr.Accordion("Instrucciones (léelas antes de empezar)", open=False):
            with gr.Tabs():
                with gr.Tab("Cómo funciona"):
                    gr.Markdown(AYUDA_FUNCIONAMIENTO)
                with gr.Tab("Preparar el Excel"):
                    gr.Markdown(AYUDA_EXCEL_ENTRADA)
                with gr.Tab("Leer el resultado"):
                    gr.Markdown(AYUDA_EXCEL_SALIDA)

        archivo = gr.File(label="Excel de entrada (.xlsx)", file_types=[".xlsx", ".xls"])
        email = gr.Textbox(label="Tu email (obligatorio)", placeholder="nombre@institucion.es")
        with gr.Row():
            hacer_doi = gr.Checkbox(True, label="Búsqueda por DOI")
            hacer_titulo = gr.Checkbox(True, label="Búsqueda por título")
        fuentes = gr.CheckboxGroup(list(FUENTES), value=["Crossref", "OpenAlex", "Unpaywall"],
                                   label="Fuentes (BASE requiere IP autorizada; CORE requiere clave)")
        umbral = gr.Slider(0.5, 1.0, value=0.85, step=0.01,
                           label="Umbral de coincidencia (título + autor + año + revista)")
        with gr.Row():
            workers = gr.Slider(1, 5, value=3, step=1, label="Filas en paralelo")
            usar_cache = gr.Checkbox(True, label="Usar caché (permite reanudar)")
            detalle = gr.Checkbox(True, label="Incluir detalle por fuente")
        escalonar = gr.Checkbox(
            True, label="Consultar CORE, Semantic Scholar y BASE solo si las demás fuentes no confirman "
                        "la publicación (recomendado: tienen límites de uso estrictos)")
        core_detalle = gr.Checkbox(
            True, label="CORE: detallar cada repositorio donde está la publicación "
                        "(una consulta extra por repositorio; gasta más cuota de CORE)")
        with gr.Accordion("Claves API", open=False):
            key_core = gr.Textbox(label="CORE (necesaria para usar CORE)", type="password")
            key_oa = gr.Textbox(label="OpenAlex (opcional)", type="password")
            key_s2 = gr.Textbox(label="Semantic Scholar (opcional)", type="password")
        with gr.Row():
            boton = gr.Button("Consultar", variant="primary")
            vaciar = gr.Button("Vaciar caché")

        estado = gr.Textbox(label="Resumen", interactive=False, lines=3)
        with gr.Tabs():
            with gr.Tab("Resultado por DOI"):
                descarga_doi = gr.File(label="Descargar Excel (búsqueda por DOI)")
                tabla_doi = gr.Dataframe(label="Vista previa (primeras 200 filas)")
            with gr.Tab("Resultado por título"):
                descarga_titulo = gr.File(label="Descargar Excel (búsqueda por título)")
                tabla_titulo = gr.Dataframe(label="Vista previa (primeras 200 filas)")

        boton.click(ejecutar,
                    [archivo, email, hacer_doi, hacer_titulo, fuentes, umbral, workers, usar_cache,
                     detalle, escalonar, core_detalle, key_core, key_oa, key_s2],
                    [estado, tabla_doi, descarga_doi, tabla_titulo, descarga_titulo])
        vaciar.click(limpiar, None, estado)
    return demo



demo = crear_app()

if __name__ == "__main__":
    demo.launch()
