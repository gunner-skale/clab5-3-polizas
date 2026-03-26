# app.py - Versión simplificada solo para comparación de Excel vs Excel
import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from google import genai
from google.genai import types
from io import BytesIO
import os
import time
from dotenv import load_dotenv

# Cargar variables de entorno
load_dotenv()

# ============================================
# CONFIGURACIÓN
# ============================================
st.set_page_config(page_title="Comparador de Pólizas AI", page_icon="🤖", layout="wide")

# CSS
st.markdown("""
<style>
    .main-header { font-size: 2.5rem; font-weight: bold; color: #1f77b4; text-align: center; margin-bottom: 1rem; }
    .sub-header { font-size: 1.2rem; color: #666; text-align: center; margin-bottom: 2rem; }
    .stButton>button { width: 100%; background-color: #1f77b4; color: white; font-weight: bold; padding: 0.75rem; }
    .stButton>button:hover { background-color: #145a8a; }
    .success-box { background-color: #d4edda; color: #155724; padding: 1rem; border-radius: 0.5rem; border-left: 5px solid #28a745; }
    .error-box { background-color: #f8d7da; color: #721c24; padding: 1rem; border-radius: 0.5rem; border-left: 5px solid #dc3545; }
    .warning-box { background-color: #fff3cd; color: #856404; padding: 1rem; border-radius: 0.5rem; border-left: 5px solid #ffc107; }
    .info-box { background-color: #d1ecf1; color: #0c5460; padding: 1rem; border-radius: 0.5rem; border-left: 5px solid #17a2b8; }
</style>
""", unsafe_allow_html=True)

# ============================================
# FUNCIONES UTILITARIAS
# ============================================

def sanitizar_texto(texto):
    """Limpia texto para evitar errores de encoding"""
    if not texto:
        return ""
    texto = str(texto)
    texto = ''.join(char for char in texto if ord(char) >= 32 or char in ['\n', '\t', '\r'])
    texto = ' '.join(texto.split())
    return texto.strip().upper()

def detectar_ok_directo(texto_respuesta):
    """
    Detecta si la respuesta indica OK sin necesidad de consultar IA
    """
    texto = sanitizar_texto(texto_respuesta)

    indicadores_ok = [
        "SE OTORGA", "OTORGA", "SE ACEPTA", "ACEPTA", "CUBRE", "INCLUYE",
        "AMPARA", "PROTEGE", "CONCEDE", "AUTORIZA", "APRUEBA", "SI", "SÍ"
    ]

    indicadores_no = [
        "NO SE OTORGA", "NO OTORGA", "NO SE ACEPTA", "NO ACEPTA",
        "NO CUBRE", "NO INCLUYE", "EXCLUYE", "RECHAZA", "DENIEGA", "NO APLICA"
    ]

    for indicador in indicadores_ok:
        if indicador in texto:
            negado = False
            for neg in indicadores_no:
                if neg in texto:
                    negado = True
                    break
            if not negado:
                return True, f"✅ {indicador.title()}"

    for indicador in indicadores_no:
        if indicador in texto:
            return True, f"DIFERENCIA: {indicador.title()}"

    return False, None

# ============================================
# FUNCIONES DE IA
# ============================================

def inicializar_cliente():
    """Inicializa cliente Gemini desde variable de entorno"""
    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        st.error("❌ No se encontró GEMINI_API_KEY en el archivo .env")
        st.stop()
    return genai.Client(api_key=api_key)

def comparar_lote_con_mejoras(items_lote, client, max_retries=2):
    """
    Compara lote de items detectando MEJORAS, RETROCESOS y EQUIVALENCIAS
    """
    if not items_lote:
        return []

    resultados_preliminares = []
    items_para_ia = []

    # Primero, detección rápida de OK/DIFERENCIA básica
    for idx, orig, resp in items_lote:
        es_ok_directo, motivo = detectar_ok_directo(resp)
        if es_ok_directo:
            if "DIFERENCIA" in motivo:
                resultados_preliminares.append((idx, "DIFERENCIA", motivo, motivo))
            else:
                resultados_preliminares.append((idx, "OK", motivo, "Condiciones equivalentes"))
        else:
            items_para_ia.append((idx, orig, resp))
            resultados_preliminares.append((idx, "PENDIENTE", None, None))

    if not items_para_ia:
        return [(idx, est, obs, detalle) for idx, est, obs, detalle in resultados_preliminares if est != "PENDIENTE"]

    # Prompt mejorado para detectar mejoras y retrocesos
    prompt_items = ""
    mapeo_ia = {}

    for i, (idx, orig, resp) in enumerate(items_para_ia, 1):
        mapeo_ia[i] = idx
        prompt_items += f"""
--- ITEM {i} ---
PÓLIZA ANTERIOR (NUESTRA PROPUESTA): {orig[:500]}
NUEVA PÓLIZA (RESPUESTA ASEGURADORA): {resp[:400]}
"""

    prompt = f"""Actúa como AUDITOR SENIOR DE SEGUROS con 15 años de experiencia. Tu tarea es COMPARAR la póliza anterior con la nueva póliza y determinar si hay MEJORAS, RETROCESOS o son EQUIVALENTES.

{prompt_items}

CRITERIOS DE EVALUACIÓN DETALLADOS:

1. **MEJORA (MEJORA)**: La nueva póliza es SUPERIOR en al menos uno de estos aspectos:
   - Monto asegurado MAYOR
   - Deducible MENOR
   - Cobertura MÁS AMPLIA (incluye más riesgos)
   - Exclusiones REDUCIDAS o ELIMINADAS
   - Condiciones MÁS FAVORABLES para el asegurado
   - Plazo MÁS LARGO
   - Prima MENOR (mismo o mejor cobertura)

2. **RETROCESO (RETROCESO)**: La nueva póliza es INFERIOR en al menos uno de estos aspectos:
   - Monto asegurado MENOR
   - Deducible MAYOR
   - Cobertura MÁS RESTRINGIDA (excluye riesgos que antes cubría)
   - Nuevas exclusiones o condiciones restrictivas
   - Condiciones MENOS FAVORABLES
   - Plazo MÁS CORTO
   - Prima MAYOR (misma o peor cobertura)

3. **EQUIVALENTE (OK)**: Las condiciones son ESENCIALMENTE IGUALES, con diferencias no sustanciales

4. **DIFERENCIA (DIFERENCIA)**: Hay cambios significativos pero no claramente mejor o peor (ej: cambia estructura, no comparable directamente)

REGLAS DE ANÁLISIS:
- Si hay MEJORA en un aspecto pero RETROCESO en otro, prioriza el análisis más relevante para el asegurado
- Para montos: considera significativo >10% de diferencia
- Para exclusiones: cualquier exclusión nueva es un RETROCESO
- Para inclusiones: cualquier cobertura nueva es una MEJORA

FORMATO OBLIGATORIO - UNA LÍNEA POR ITEM CON 3 PARTES SEPARADAS POR "|":
RESULTADO_{i}: TIPO|MENSAJE_CORTO|ANÁLISIS_DETALLADO

Donde:
- TIPO: MEJORA, RETROCESO, OK, DIFERENCIA
- MENSAJE_CORTO: Frase breve de 15-30 palabras resumiendo el cambio
- ANÁLISIS_DETALLADO: Explicación extensa de 50-150 palabras

EJEMPLOS:
RESULTADO_1: MEJORA|Monto asegurado aumentó de $100,000 a $150,000 (+50%)|Se incrementó significativamente la suma asegurada, ofreciendo mayor protección al asegurado sin cambios en las exclusiones.
RESULTADO_2: RETROCESO|Se eliminó cobertura por inundación que sí existía antes|La nueva póliza excluye explícitamente daños por inundación, un riesgo relevante que estaba cubierto en la póliza anterior.
RESULTADO_3: OK|Condiciones equivalentes, solo cambios menores de redacción|La cobertura y condiciones son esencialmente las mismas, no hay cambios sustanciales en montos, deducibles o exclusiones.
RESULTADO_4: DIFERENCIA|Cambio en estructura de coberturas|La nueva póliza reorganiza las coberturas en un formato diferente, difícil de comparar directamente. Se requiere análisis adicional de equivalencia real."""

    for intento in range(max_retries):
        try:
            response = client.models.generate_content(
                model="models/gemini-3.1-flash-lite-preview",
                contents=prompt,
                config=types.GenerateContentConfig(
                    temperature=0.1,
                    max_output_tokens=1024,
                    top_p=0.95,
                )
            )

            if not response or not response.text:
                if intento < max_retries - 1:
                    time.sleep(0.5)
                    continue
                # Fallback: marcar como error
                for i, (idx, orig, resp) in enumerate(items_para_ia, 1):
                    for j, (id_res, est, obs, det) in enumerate(resultados_preliminares):
                        if id_res == idx and est == "PENDIENTE":
                            resultados_preliminares[j] = (idx, "ERROR", "Sin respuesta API", "No se pudo obtener análisis de la IA")
                break

            # Parsear resultados
            lineas = response.text.strip().split('\n')
            resultados_ia = {}

            for linea in lineas:
                linea = linea.strip()
                if not linea.startswith("RESULTADO_"):
                    continue

                try:
                    # Extraer número de item
                    num_str = linea.split("_")[1].split(":")[0]
                    num = int(num_str)

                    # Extraer las tres partes separadas por "|"
                    contenido = linea.split(":", 1)[1].strip() if ":" in linea else linea

                    if "|" in contenido:
                        partes = contenido.split("|")
                        tipo = partes[0].strip().upper()
                        mensaje_corto = partes[1].strip() if len(partes) > 1 else "Sin resumen"
                        detalle = partes[2].strip() if len(partes) > 2 else "Sin detalles"

                        # Validar tipo
                        if tipo not in ["MEJORA", "RETROCESO", "OK", "DIFERENCIA"]:
                            tipo = "DIFERENCIA"  # Default

                        resultados_ia[num] = (tipo, mensaje_corto, detalle)
                    else:
                        # Fallback si no hay separadores
                        resultados_ia[num] = ("DIFERENCIA", contenido[:100], contenido)

                except Exception as e:
                    continue

            # Aplicar resultados a los preliminares
            for i, (idx, orig, resp) in enumerate(items_para_ia, 1):
                if i in resultados_ia:
                    tipo, mensaje, detalle = resultados_ia[i]
                    for j, (id_res, est, obs, det) in enumerate(resultados_preliminares):
                        if id_res == idx and est == "PENDIENTE":
                            resultados_preliminares[j] = (idx, tipo, mensaje, detalle)
                            break
                else:
                    for j, (id_res, est, obs, det) in enumerate(resultados_preliminares):
                        if id_res == idx and est == "PENDIENTE":
                            resultados_preliminares[j] = (idx, "ERROR", "No procesado", "El item no pudo ser analizado correctamente")
                            break

            break

        except Exception as e:
            if intento < max_retries - 1:
                time.sleep(1)
                continue
            for i, (idx, orig, resp) in enumerate(items_para_ia, 1):
                for j, (id_res, est, obs, det) in enumerate(resultados_preliminares):
                    if id_res == idx and est == "PENDIENTE":
                        resultados_preliminares[j] = (idx, "ERROR", f"Error técnico", f"Error: {str(e)[:100]}")

    # Retornar solo los que no están pendientes
    return [(idx, tipo, mensaje, detalle) for idx, tipo, mensaje, detalle in resultados_preliminares if tipo != "PENDIENTE"]

# ============================================
# PROCESAMIENTO DE EXCEL - SOLO HOJAS "PP*"
# ============================================

def procesar_excel_modo1_mejorado(file, col_propuesta, col_respuesta, fila_inicio, progress_bar, status_text, tamanio_lote=3):
    """
    MODO 1 MEJORADO: Compara Excel vs Excel identificando MEJORAS, RETROCESOS y EQUIVALENCIAS
    """
    # Colores para mejor visualización
    VERDE = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")      # OK / Equivalente
    ROJO = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")        # Retroceso / Diferencia
    AMARILLO = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")    # Advertencia
    AZUL_CLARO = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")  # Mejora significativa

    client = inicializar_cliente()
    wb = load_workbook(file)

    resultados_totales = []
    hojas_procesadas = []

    hojas_pp = [name for name in wb.sheetnames if name.startswith("PP")]

    if not hojas_pp:
        status_text.text("❌ No hay hojas 'PP'")
        return wb, []

    total_hojas = len(hojas_pp)

    for idx_hoja, sheet_name in enumerate(hojas_pp, 1):
        ws = wb[sheet_name]

        # Crear columnas adicionales para mejor análisis
        col_resultado = ws.max_column + 1
        col_tipo_cambio = ws.max_column + 2
        col_observacion_detallada = ws.max_column + 3

        # Headers
        headers = [
            ("RESULTADO COMPARACIÓN", col_resultado),
            ("TIPO DE CAMBIO", col_tipo_cambio),
            ("OBSERVACIÓN DETALLADA", col_observacion_detallada)
        ]

        for header_value, col_num in headers:
            header_cell = ws.cell(row=fila_inicio-1, column=col_num)
            header_cell.value = header_value
            header_cell.font = Font(bold=True, size=11, color="FFFFFF")
            header_cell.fill = PatternFill(start_color="1f77b4", end_color="1f77b4", fill_type="solid")

        # Recolectar filas
        filas_datos = []
        for row in range(fila_inicio, ws.max_row + 1):
            val_prop = str(ws.cell(row=row, column=col_propuesta).value or "").strip()
            val_resp = str(ws.cell(row=row, column=col_respuesta).value or "").strip()
            if val_prop and val_resp and len(val_prop) > 5:
                filas_datos.append((row, val_prop, val_resp))

        if not filas_datos:
            continue

        total_filas = len(filas_datos)
        procesadas = 0

        # Procesar por lotes
        for i in range(0, len(filas_datos), tamanio_lote):
            lote = filas_datos[i:i + tamanio_lote]

            filas_str = ", ".join([str(r) for r, _, _ in lote])
            status_text.text(f"Hoja {idx_hoja}/{total_hojas} '{sheet_name}': Analizando filas {filas_str} ({procesadas+1}-{min(procesadas+len(lote), total_filas)}/{total_filas})")

            items_lote = [(row, prop, resp) for row, prop, resp in lote]

            # Usar el prompt mejorado para detectar mejoras/retrocesos
            resultados_lote = comparar_lote_con_mejoras(items_lote, client)

            # Aplicar resultados con colores y observaciones detalladas
            for resultado_analisis in resultados_lote:
                idx, tipo, mensaje_corto, detalle_largo = resultado_analisis

                cell_res = ws.cell(row=idx, column=col_resultado)
                cell_tipo = ws.cell(row=idx, column=col_tipo_cambio)
                cell_detalle = ws.cell(row=idx, column=col_observacion_detallada)

                # Asignar valores según el tipo
                if tipo == "MEJORA":
                    cell_res.value = f"✅ MEJORA: {mensaje_corto}"
                    cell_res.fill = AZUL_CLARO
                    cell_res.font = Font(color="0066CC", bold=True, size=9)

                    cell_tipo.value = "MEJORA SIGNIFICATIVA"
                    cell_tipo.fill = AZUL_CLARO
                    cell_tipo.font = Font(color="0066CC", bold=True)

                elif tipo == "RETROCESO":
                    cell_res.value = f"⚠️ RETROCESO: {mensaje_corto}"
                    cell_res.fill = ROJO
                    cell_res.font = Font(color="9C0006", bold=True, size=9)

                    cell_tipo.value = "RETROCESO EN CONDICIONES"
                    cell_tipo.fill = ROJO
                    cell_tipo.font = Font(color="9C0006", bold=True)

                elif tipo == "OK":
                    cell_res.value = f"✅ EQUIVALENTE: {mensaje_corto}"
                    cell_res.fill = VERDE
                    cell_res.font = Font(color="006100", bold=True, size=9)

                    cell_tipo.value = "CONDICIONES EQUIVALENTES"
                    cell_tipo.fill = VERDE
                    cell_tipo.font = Font(color="006100", bold=True)

                elif tipo == "DIFERENCIA":
                    cell_res.value = f"⚠️ DIFERENCIA: {mensaje_corto}"
                    cell_res.fill = ROJO
                    cell_res.font = Font(color="9C0006", bold=True, size=9)

                    cell_tipo.value = "DIFERENCIA SIGNIFICATIVA"
                    cell_tipo.fill = ROJO
                    cell_tipo.font = Font(color="9C0006", bold=True)

                else:  # ERROR
                    cell_res.value = f"❌ ERROR: {mensaje_corto}"
                    cell_res.fill = AMARILLO
                    cell_res.font = Font(color="9C5700", bold=True, size=9)

                    cell_tipo.value = "ERROR DE PROCESAMIENTO"
                    cell_tipo.fill = AMARILLO
                    cell_tipo.font = Font(color="9C5700", bold=True)

                # Observación detallada
                cell_detalle.value = detalle_largo
                cell_detalle.alignment = Alignment(wrap_text=True, vertical="top")
                cell_detalle.font = Font(size=8)

                # Guardar en resultados
                resultados_totales.append({
                    'hoja': sheet_name,
                    'fila': idx,
                    'tipo': tipo,
                    'resumen': mensaje_corto,
                    'detalle': detalle_largo
                })

            procesadas += len(lote)

            prog_hoja = (idx_hoja - 1) / total_hojas
            prog_fila = (procesadas / total_filas) / total_hojas
            progress_bar.progress(min(prog_hoja + prog_fila, 0.99))

            time.sleep(0.1)

        # Ajustar anchos
        from openpyxl.utils import get_column_letter
        ws.column_dimensions[get_column_letter(col_resultado)].width = 45
        ws.column_dimensions[get_column_letter(col_tipo_cambio)].width = 25
        ws.column_dimensions[get_column_letter(col_observacion_detallada)].width = 80
        hojas_procesadas.append(f"{sheet_name} ({total_filas})")

    progress_bar.progress(1.0)
    status_text.text(f"✅ {len(hojas_procesadas)} hojas procesadas con análisis de mejoras")
    return wb, resultados_totales

# ============================================
# INTERFAZ
# ============================================

def main():
    st.markdown('<div class="main-header">🤖 Comparador Inteligente de Pólizas</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-header">Compara pólizas anteriores vs nuevas identificando mejoras y retrocesos</div>', unsafe_allow_html=True)

    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        st.markdown('<div class="error-box"><strong>❌ Error:</strong> Crear archivo <code>.env</code> con GEMINI_API_KEY=tu_key</div>', unsafe_allow_html=True)
        st.stop()

    st.markdown('<div class="success-box"><strong>✅ API Key configurada</strong> | Usando Gemini 2.5 Flash</div>', unsafe_allow_html=True)

    st.sidebar.header("⚙️ Configuración")

    st.sidebar.info("""
    **📝 Modo de operación:**
    Compara tu póliza anterior (Propuesta) vs nueva póliza (Respuesta)

    **🎯 Detección automática:**
    - ✅ MEJORAS: Coberturas superiores
    - ⚠️ RETROCESOS: Coberturas inferiores
    - 🔄 EQUIVALENTES: Condiciones similares
    """)

    col_prop = st.sidebar.selectbox("Columna PÓLIZA ANTERIOR (tu propuesta)", ["A","B","C","D","E","F","G","H","I","J"], index=2)
    col_prop_num = ord(col_prop) - ord('A') + 1

    col_resp = st.sidebar.selectbox("Columna NUEVA PÓLIZA (respuesta aseguradora)", ["A","B","C","D","E","F","G","H","I","J"], index=4)
    col_resp_num = ord(col_resp) - ord('A') + 1

    fila_inicio = st.sidebar.number_input("Fila donde inician los datos", min_value=1, value=16)

    tamanio_lote = st.sidebar.selectbox("Tamaño de lote", ["Individual (1)", "Parejas (2)", "Tríos (3)"], index=2)
    tamanio_lote_num = int(tamanio_lote.split("(")[1].split(")")[0])

    st.sidebar.markdown("---")
    st.sidebar.markdown("""
    **📊 Columnas de resultado:**
    - **Columna 1**: Resultado resumido
    - **Columna 2**: Tipo de cambio
    - **Columna 3**: Análisis detallado

    **🎨 Colores:**
    - 🔵 Azul: MEJORA significativa
    - 🔴 Rojo: RETROCESO o DIFERENCIA
    - 🟢 Verde: EQUIVALENTE
    - 🟡 Amarillo: ERROR
    """)

    st.subheader("📁 Subir Archivo Excel")

    uploaded_excel = st.file_uploader("Excel con pólizas a comparar (.xlsx)", type=['xlsx', 'xls'], key="excel")

    if uploaded_excel:
        try:
            df_preview = pd.read_excel(uploaded_excel, header=None).head(8)
            with st.expander("👁️ Vista previa del Excel"):
                st.dataframe(df_preview, use_container_width=True)
                st.caption(f"📌 Las hojas deben comenzar con 'PP' para ser procesadas")
            uploaded_excel.seek(0)
        except Exception as e:
            st.error(f"Error leyendo Excel: {e}")
            st.stop()

        if st.button("🚀 INICIAR COMPARACIÓN", type="primary"):
            progress_bar = st.progress(0)
            status_text = st.empty()

            with st.spinner("Procesando comparación con análisis de mejoras..."):
                try:
                    wb, resultados = procesar_excel_modo1_mejorado(
                        uploaded_excel, col_prop_num, col_resp_num, fila_inicio,
                        progress_bar, status_text, tamanio_lote=tamanio_lote_num
                    )

                    if not resultados:
                        st.warning("No se procesaron datos. Verificar que existan hojas 'PP' con datos válidos")
                        st.stop()

                    output = BytesIO()
                    wb.save(output)
                    output.seek(0)

                    # Estadísticas
                    total_mejoras = sum(1 for r in resultados if r['tipo'] == 'MEJORA')
                    total_retrocesos = sum(1 for r in resultados if r['tipo'] == 'RETROCESO')
                    total_ok = sum(1 for r in resultados if r['tipo'] == 'OK')
                    total_diff = sum(1 for r in resultados if r['tipo'] == 'DIFERENCIA')
                    total_err = sum(1 for r in resultados if r['tipo'] == 'ERROR')

                    hojas_unicas = list(set([r['hoja'] for r in resultados]))

                    st.info(f"📊 Hojas procesadas: {', '.join(hojas_unicas)}")

                    col1, col2, col3, col4, col5 = st.columns(5)
                    col1.metric("Total items", len(resultados))
                    col2.metric("✅ MEJORAS", total_mejoras, delta="+", delta_color="normal")
                    col3.metric("⚠️ RETROCESOS", total_retrocesos, delta="-", delta_color="inverse")
                    col4.metric("🔄 EQUIVALENTES", total_ok)
                    col5.metric("📊 DIFERENCIAS", total_diff)

                    with st.expander("📋 Detalle completo de resultados"):
                        df_res = pd.DataFrame(resultados)
                        st.dataframe(df_res, use_container_width=True)

                    st.download_button(
                        "📥 DESCARGAR EXCEL CON ANÁLISIS",
                        data=output,
                        file_name=f"COMPARACION_POLIZAS_{uploaded_excel.name}",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                    st.balloons()
                    st.success(f"✅ Análisis completado: {len(resultados)} coberturas evaluadas")
                    st.markdown(f"""
                    <div class="info-box">
                    <strong>📈 Resumen del análisis:</strong><br>
                    • 🟢 Mejoras identificadas: {total_mejoras}<br>
                    • 🔴 Retrocesos identificados: {total_retrocesos}<br>
                    • 🟡 Condiciones equivalentes: {total_ok}<br>
                    • 📊 Diferencias no clasificadas: {total_diff}
                    </div>
                    """, unsafe_allow_html=True)

                except Exception as e:
                    st.error(f"Error en procesamiento: {e}")
                    st.exception(e)

if __name__ == "__main__":
    main()
