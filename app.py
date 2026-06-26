import streamlit as st
import os
import json
import re
import tempfile
from datetime import datetime # <--- AGREGA ESTA LÍNEA AQUÍ
import google.generativeai as genai
import io
from docx import Document
# ==========================================
# PROMPT MAESTRO INTEGRADO
# ==========================================
PROMPT_MAESTRO = """
**Rol y Objetivo:**
Actúa como un Experto Evaluador en Protección Civil del área de inspecciones del Estado de Michoacán. Tu tarea es analizar exhaustivamente el documento adjunto correspondiente a un Programa Interno de Protección Civil (PIPC) de Alto Riesgo.

**PARÁMETROS ESTRICTOS DE AUDITORÍA:**
* **Fecha actual de evaluación:** 26 de junio de 2026. Úsala matemáticamente para determinar vigencias.
* **COTEJO DE ÍNDICE Y CARACTERÍSTICAS (OBLIGATORIO):** Compara la estructura del documento evaluado contra los dos índices oficiales de Michoacán: 
  1) "GUIA TÉCNICA PIPC michRV9" (que exige plan operativo, contingencias, continuidad, y 30 características estrictas).
  2) "INDICE_DE_PIPC 2016-2021".
  Identifica cuál de los dos utilizó el consultor. Revisa estrictamente que se cumplan las características: firmas en el acta constitutiva, carta de corresponsabilidad, INE, constancias de capacitación con registro vigente, y croquis completos.
* **ANÁLISIS FOTOGRÁFICO:** Rechaza fotos de internet (marcas de agua, baja resolución) o recicladas (fechas antiguas, desgaste ilógico). Exige calidad in situ.

**ESTRUCTURA DE RESPUESTA REQUERIDA:**

**--- BLOQUE A: DICTAMEN INTERNO COMPLETO ---**
**FASE 1: Identificación y Marco Normativo**
* Nombre de la empresa/inmueble y giro específico.
* Marco normativo aplicable.

**FASE 2: Verificación de Índice y Estructura**
* **Índice Utilizado:** (Indica si el documento se basó en la Guía RV9, en el Índice 2016-2021, o en Ninguno/Desordenado).
* **Apartados Faltantes:** (Enlista los capítulos o subprogramas que el consultor omitió según el índice oficial).
* **Cumplimiento de Características:** (Indica si faltan firmas originales, carta de corresponsabilidad, acreditaciones, etc.).

**FASE 3: Evaluación Exhaustiva de Dictámenes Técnicos**
Por CADA dictamen técnico anexado, extrae y presenta:
* **Tipo de Dictamen:** (Eléctrico, Estructural, Gas, etc.).
* **SERVER GUID:** (Exclusivo eléctrico. Si no existe, escribe: "ALERTA: NO PRESENTA SERVER GUID").
* **Nombre del Dictaminador:** (Persona física que firma).
* **Registro Institucional / UVSEIE:** (Número de registro o unidad).
* **Vigencia del Registro:** (VIGENTE o VENCIDO al 26 de junio de 2026).
* **Fecha de Elaboración:** (Emisión del dictamen).
* **Vigencia del Dictamen:** (VIGENTE o VENCIDO al 26 de junio de 2026).
* **Observaciones y Sentencia:** (Anomalías y conclusión del perito).

**FASE 4: Matriz de Observaciones (Formato Markdown)**
Genera una tabla estricta. Aquí DEBES incluir como filas independientes todas las omisiones de índice (Fase 2), anomalías en dictámenes (Fase 3), y fallas fotográficas.
* **Página:** (Número).
* **Ubicación:** (Apartado/Título).
* **Observación:** (Describe la discrepancia o faltante).
* **Justificación Normativa:** (Cita el apartado del Índice oficial, Guía RV9, o Ley/NOM aplicable).
* **Acción Solicitada:** (Instrucción correctiva estricta para el consultor).

**FASE 5: Resolución**
Párrafo indicando si el programa es Aprobado, Aprobado con Condicionantes o Rechazado.

**--- BLOQUE B: OFICIO DE OBSERVACIONES (Formato Legal Michoacán) ---**
Genera un oficio formal dirigido al promovente, redactado estrictamente con esta plantilla:

Morelia, Michoacán a 26 de junio de 2026.
[NOMBRE DEL CONSULTOR RESPONSABLE, o escribir: CONSULTOR EXTERNO]
P R E S E N T E.

En relación con la revisión de los Programas internos de Protección Civil promovidos solicitando “Su validación y en su caso la expedición de la constancia de cumplimiento”, de la empresa [NOMBRE DE LA EMPRESA Y SUCURSAL], al respecto le comunico a Usted lo siguiente:

Habiendo realizado una valoración respecto a las condiciones y Apego al contenido del programa Interno y de haber realizado la verificación de las instalaciones se ha encontrado lo siguiente al momento de la revisión:

R E S U L T A D O:
PRIMERO. - Del resultado de la revisión realizada al Programa Interno en mención, se observó que los documentos presentados, NO CUMPLEN con los contenidos del índice para la elaboración de programas internos de protección civil.
SEGUNDO. - por lo tanto, esta Coordinación Estatal de Protección Civil Se hace saber al promovente que previo a resolver en cuanto al fondo, se le REQUIERE mediante NOTIFICACIÓN PERSONAL para que dentro del plazo de quince días cumpla con las modificaciones que se señalan a continuación:

[AQUÍ INSERTA LAS OBSERVACIONES EN FORMA DE LISTA EJECUTIVA. Extrae las fallas de la Fase 4 (Matriz), incluyendo omisiones del índice, problemas con dictámenes y fotografías].
* Al no presentar completo el contenido del Programa Interno no se ha evaluado completo por lo que al solventar pudieran generarse nuevas observaciones.

D E R E C H O:
Por lo antes expuesto, con fundamento legal en los Artículos 1, 2, 3, 4, 60, 61, 62, y demás relativos aplicables de la Ley de Protección Civil del Estado de Michoacán de Ocampo, esta Coordinación Estatal de Protección Civil determina otorgar la presente Opinión.
Así mismo, se le informa que la revisión física que se llevó a cabo al documento de referencia, no lo exime de las demás observaciones y recomendaciones que le hagan los demás organismos normativos de los diversos órdenes de gobierno.

Sin otro particular, hago propicia la ocasión para enviarle un cordial saludo.

A T E N T A M E N T E

______________________________________
DEPARTAMENTO DE INSPECCIÓN
"""

# ==========================================
# CONFIGURACIÓN DE LA INTERFAZ
# ==========================================
st.set_page_config(page_title="Auditor PIPC", page_icon="📋", layout="wide")
st.title("📋 Sistema Automatizado de Evaluación PIPC")
st.markdown("Sube el documento del Programa Interno, añade tus notas y genera el dictamen estructurado automáticamente.")

# Barra lateral para configuración
with st.sidebar:
    st.header("Configuración del Sistema")
    api_key = st.text_input("Ingresa tu API Key de Gemini", type="password")
    
    st.markdown("---")
    st.subheader("Configuración de Nube (Drive)")
    st.info("Asegúrate de tener instalado 'Google Drive para Escritorio' en tu PC.")
    # Cambia el 'value' por la ruta real de tu Drive en tu computadora
    ruta_base_drive = st.text_input(
        "Ruta local de Google Drive:", 
        value="G:/Mi unidad/Biblioteca_PIPC" 
    )

    if not api_key:
        st.warning("⚠️ Necesitas ingresar tu API Key para continuar.")
# Área principal de entrada
col1, col2 = st.columns([1, 1])

with col1:
    st.markdown("### 📂 Carga de Documentos")
    metodo_carga = st.radio(
        "Selecciona cómo deseas subir el PIPC:",
        ["Seleccionar archivos manualmente", "Ruta de carpeta local (Recomendado)"]
    )

    archivos_pdf_manuales = []
    rutas_pdf_locales = []

    if metodo_carga == "Seleccionar archivos manualmente":
        archivos_pdf_manuales = st.file_uploader("Sube los archivos PDF", type=["pdf"], accept_multiple_files=True)
    else:
        ruta_carpeta = st.text_input(
            "Pega la ruta de la carpeta principal:", 
            placeholder="Ej: C:\\Users\\pgss1\\Documentos\\PIPC_Empresa"
        )
        if ruta_carpeta and os.path.exists(ruta_carpeta):
            # Buscar todos los PDFs incluso en subcarpetas
            for root, dirs, files in os.walk(ruta_carpeta):
                for file in files:
                    if file.lower().endswith('.pdf'):
                        rutas_pdf_locales.append(os.path.join(root, file))
            st.success(f"✅ Se encontraron {len(rutas_pdf_locales)} archivos PDF listos para analizar.")
        elif ruta_carpeta:
            st.error("⚠️ La ruta no existe. Verifica que esté bien escrita.")

with col2:
    st.markdown("### 📝 Notas de Inspección")
    directivas = st.text_area(
        "Directivas Previas del Inspector:", 
        height=180,
        placeholder="Ej: Revisar detalladamente las responsivas técnicas. Si el inmueble declara usar gas LP, ser estricto con las evidencias fotográficas de los tanques."
    )

# Botón de ejecución
if st.button("Ejecutar Revisión Completa", type="primary"):
    if not api_key:
        st.error("Por favor, ingresa tu API Key en el panel izquierdo.")
    elif not archivos_pdf_manuales and not rutas_pdf_locales:
        st.error("Por favor, proporciona al menos un archivo PDF o una ruta válida.")
    else:
        with st.status("Iniciando auditoría técnica...", expanded=True) as status:
            try:
                client = genai.Client(api_key=api_key)
                documentos_nube = []
                rutas_temporales = []

                st.write("📥 Preparando archivos...")
                
                # Procesar archivos subidos manualmente
                if archivos_pdf_manuales:
                    for archivo in archivos_pdf_manuales:
                        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
                            tmp_file.write(archivo.getvalue())
                            rutas_temporales.append(tmp_file.name)
                        st.write(f"☁️ Subiendo '{archivo.name}'...")
                        doc_nube = client.files.upload(file=tmp_file.name)
                        documentos_nube.append(doc_nube)
                
                # Procesar ruta de carpeta (inmune a acentos en los nombres)
                elif rutas_pdf_locales:
                    for ruta_pdf in rutas_pdf_locales:
                        nombre_base = os.path.basename(ruta_pdf)
                        # Le quitamos los acentos temporalmente solo para el letrero de la pantalla
                        nombre_seguro = nombre_base.encode('ascii', 'ignore').decode('ascii')
                        st.write(f"☁️ Subiendo '{nombre_seguro}'...")
                        
                        # Creamos una copia temporal segura para que la IA no choque con los acentos
                        with open(ruta_pdf, 'rb') as archivo_original:
                            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
                                tmp_file.write(archivo_original.read())
                                rutas_temporales.append(tmp_file.name)
                                
                        doc_nube = client.files.upload(file=tmp_file.name)
                        documentos_nube.append(doc_nube)
                        fecha_hoy = datetime.now().strftime("%d de %B de %Y")
                inputs_usuario = f"**Fecha Actual:** {fecha_hoy}\n**Directivas Previas:**\n{directivas}"
                st.write("🧠 Analizando Focos Rojos y evaluando marco normativo...")
                elementos_a_evaluar = documentos_nube + [PROMPT_MAESTRO, inputs_usuario]
                       

                # Sistema anti-saturación
                max_reintentos = 3
                for intento in range(max_reintentos):
                    try:
                        response = client.models.generate_content(
                            model='gemini-2.5-flash', 
                            contents=elementos_a_evaluar
                        )
                        texto_resultado = response.text
                        break
                    except Exception as api_error:
                        if "503" in str(api_error) and intento < max_reintentos - 1:
                            st.warning("⏳ Tráfico alto en servidores. Reintentando en 15 segundos...")
                            time.sleep(15)
                        else:
                            raise api_error

                st.write("💾 Estructurando dictamen...")
                
                nombre_archivo_match = re.search(r"DICTAMEN_[A-Z0-9_]+_[A-Z0-9_]+_[A-Z0-9_]+_\d{4}-\d{2}", texto_resultado)
                nombre_archivo = nombre_archivo_match.group(0) if nombre_archivo_match else "DICTAMEN_PROVISIONAL"
                
                def limpiar_texto(texto): return re.sub(r'[\\/*?:"<>|]', "", texto).strip()

                ficha_datos = {}
                for linea in texto_resultado.split("\n"):
                    if "Empresa:" in linea: ficha_datos["empresa"] = limpiar_texto(linea.split(":")[1])
                    if "Municipio:" in linea: ficha_datos["municipio"] = limpiar_texto(linea.split(":")[1])

                try:
                    fase1 = re.search(r"(?s)(\*\*FASE 1: Identificación y Marco Normativo.*?)(\*\*FASE 2)", texto_resultado).group(1)
                    # Ahora la matriz es la FASE 4
                    fase4 = re.search(r"(?s)(\*\*FASE 4: Matriz de Observaciones.*?)(\*\*FASE 5)", texto_resultado).group(1)
                    # Ahora la resolución es la FASE 5
                    fase5 = re.search(r"(?s)(\*\*FASE 5: Resolución.*?)(\*\*--- BLOQUE B)", texto_resultado).group(1)
                    
                    # El bloque B (Oficio) ya viene redactado completo desde el prompt, así que lo extraemos directamente:
                    reporte_breve_match = re.search(r"(?s)\*\*--- BLOQUE B: OFICIO DE OBSERVACIONES.*?\*\*(.*)", texto_resultado)
                    reporte_breve = reporte_breve_match.group(1).strip() if reporte_breve_match else "Error al extraer el Oficio."
                    
                except AttributeError:
                    reporte_breve = "Error al estructurar el reporte breve.\n\n" + texto_resultado
                # Guardado en Google Drive
                municipio_limpio = ficha_datos.get('municipio', 'General')
                empresa_limpia = ficha_datos.get('empresa', 'Sin_Nombre')
                carpeta_expediente = os.path.join(ruta_base_drive, municipio_limpio, empresa_limpia)
                os.makedirs(carpeta_expediente, exist_ok=True)
                
                ruta_dictamen = os.path.join(carpeta_expediente, f"{nombre_archivo}_COMPLETO.md")
                ruta_oficio = os.path.join(carpeta_expediente, f"OFICIO_{empresa_limpia}.md")
                
                with open(ruta_dictamen, "w", encoding="utf-8") as f: f.write(texto_resultado)
                with open(ruta_oficio, "w", encoding="utf-8") as f: f.write(reporte_breve)

                st.write("🧹 Limpiando archivos temporales en la nube...")
                for doc in documentos_nube: client.files.delete(name=doc.name)
                for ruta in rutas_temporales: os.remove(ruta)

                status.update(label="¡Auditoría Completada!", state="complete", expanded=False)

            except Exception as e:
                status.update(label="Error en el proceso", state="error", expanded=True)
                st.error(f"⚠️ Detalle: {e}")
                
        # Mostrar resultados y Botones de Descarga en la interfaz
        if 'texto_resultado' in locals():
            st.success("Evaluación finalizada. Revisa los resultados a continuación o descárgalos directamente.")
            
            # --- FUNCIÓN PARA CREAR EL WORD EN MEMORIA ---
            doc = Document()
            for linea in reporte_breve.split('\n'):
                # Agrega el texto al documento línea por línea
                if linea.strip(): 
                    doc.add_paragraph(linea.replace("**", "")) # Quitamos los asteriscos de negritas de markdown
            
            archivo_word = io.BytesIO()
            doc.save(archivo_word)
            archivo_word.seek(0)
            # ----------------------------------------------
            
            # Botones de descarga organizados en 3 columnas
            col_descarga1, col_descarga2, col_descarga3 = st.columns(3)
            
            with col_descarga1:
                st.download_button(
                    label="📄 Descargar Oficio en WORD",
                    data=archivo_word,
                    file_name=f"OFICIO_{empresa_limpia}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            with col_descarga2:
                st.download_button(
                    label="⬇️ Oficio en Texto (.md)",
                    data=reporte_breve,
                    file_name=f"OFICIO_{empresa_limpia}.md",
                    mime="text/markdown",
                    use_container_width=True
                )
            with col_descarga3:
                st.download_button(
                    label="⬇️ Dictamen Interno (.md)",
                    data=texto_resultado,
                    file_name=f"DICTAMEN_{empresa_limpia}.md",
                    mime="text/markdown",
                    use_container_width=True
                )
                
            tab1, tab2 = st.tabs(["Oficio para Promovente", "Dictamen Interno Completo"])
            with tab1: st.markdown(reporte_breve)
            with tab2: st.markdown(texto_resultado)