import streamlit as st
import pandas as pd
import math
import re
from io import BytesIO
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

st.set_page_config(page_title="Generador de Minutas Topográficas", layout="wide", page_icon="📐")

def limpiar_numero(val):
    if pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    if '.' in s and ',' in s:
        s = s.replace('.', '').replace(',', '.')
    elif ',' in s:
        s = s.replace(',', '.')
    s_clean = re.sub(r'[^\d.-]', '', s)
    try:
        return float(s_clean)
    except:
        return 0.0

def calcular_distancia(n1, e1, n2, e2):
    return math.sqrt((n2 - n1)**2 + (e2 - e1)**2)

def calcular_azimut(n1, e1, n2, e2):
    dn = n2 - n1
    de = e2 - e1
    rad = math.atan2(de, dn)
    grados = math.degrees(rad)
    if grados < 0:
        grados += 360
    
    g = int(grados)
    m = int((grados - g) * 60)
    s = round(((grados - g) * 60 - m) * 60, 1)
    return g, m, s, grados

def calcular_rumbo(azimut_grados):
    if 0 <= azimut_grados <= 90:
        cuadrante = "NE"
        angulo = azimut_grados
    elif 90 < azimut_grados <= 180:
        cuadrante = "SE"
        angulo = 180 - azimut_grados
    elif 180 < azimut_grados <= 270:
        cuadrante = "SW"
        angulo = azimut_grados - 180
    else:
        cuadrante = "NW"
        angulo = 360 - azimut_grados
    
    g = int(angulo)
    m = int((angulo - g) * 60)
    s = round(((angulo - g) * 60 - m) * 60, 1)
    return f"{cuadrante[0]} {g}°{m:02d}'{s:04.1f}\" {cuadrante}"

def calcular_area_gauss(n, e):
    num_puntos = len(n)
    if num_puntos < 3:
        return 0.0
    area = 0.0
    for i in range(num_puntos):
        j = (i + 1) % num_puntos
        area += e[i] * n[j] - e[j] * n[i]
    return abs(area) / 2.0

st.title("📐 Generador Automatizado de Minutas Topográficas y Linderos")
st.markdown("Genera la descripción técnica y legal de linderos, cuadro de coordenadas, cálculo de distancias, rumbos, azimuts y descarga el informe en Word (`.docx`).")

with st.expander("📝 1. Datos Generales del Predio y Profesional", expanded=True):
    col1, col2, col3 = st.columns(3)
    with col1:
        nombre_predio = st.text_input("Nombre del Predio / Lote", value="LOTE EL MIRADOR")
        municipio = st.text_input("Municipio", value="Guatavita")
        vereda = st.text_input("Vereda", value="Centro")
    with col2:
        departamento = st.text_input("Departamento", value="Cundinamarca")
        cedula_catastral = st.text_input("Cédula Catastral", value="25-000-00-00-0000-0000-000")
        matricula_inmobiliaria = st.text_input("Matrícula Inmobiliaria", value="150-00000")
    with col3:
        propietario = st.text_input("Propietario / Solicitante", value="CARLOS PÉREZ")
        profesional = st.text_input("Topógrafo / Perito", value="DOUGLAS CHAPETON GOMEZ")
        matricula_prof = st.text_input("Matrícula Profesional CPNT", value="01-19914 CPNT")

st.markdown("### 📍 2. Coordenadas y Puntos del Polígono")

archivo_subido = st.file_uploader("Cargar archivo Excel (.xlsx, .xls) o CSV con coordenadas", type=["xlsx", "xls", "csv"])

datos_por_defecto = {
    "Punto": ["1", "2", "3", "4"],
    "Norte": [2119344.68, 2119347.01, 2119247.20, 2119236.88],
    "Este": [4840129.36, 4840156.42, 4840144.94, 4840135.55],
    "Colindante": ["Servidumbre de acceso", "Predio Lote 2", "Camino Veredal", "Predio Lote 4"]
}

if archivo_subido is not None:
    try:
        if archivo_subido.name.endswith('.csv'):
            df_raw = pd.read_csv(archivo_subido)
        else:
            df_raw = pd.read_excel(archivo_subido)
        st.success(f"✅ Archivo '{archivo_subido.name}' cargado correctamente ({len(df_raw)} filas detectadas).")
    except Exception as e:
        st.error(f"Error al leer el archivo: {e}")
        df_raw = pd.DataFrame(datos_por_defecto)
else:
    df_raw = pd.DataFrame(datos_por_defecto)

cols_disponibles = list(df_raw.columns)

def buscar_columna_defecto(lista_cols, opciones):
    for col in lista_cols:
        col_clean = str(col).lower().strip()
        for op in opciones:
            if op in col_clean:
                return col
    return lista_cols[0] if lista_cols else None

col_pto_def = buscar_columna_defecto(cols_disponibles, ['punto', 'pto', 'id', 'name', 'vertice', 'est', 'estacion', 'item', 'no'])
col_nor_def = buscar_columna_defecto(cols_disponibles, ['norte', 'north', 'lat', 'y'])
col_est_def = buscar_columna_defecto(cols_disponibles, ['este', 'east', 'lon', 'x'])
col_col_def = buscar_columna_defecto(cols_disponibles, ['colindante', 'colindancia', 'vecino', 'limite', 'desc', 'observ'])

with st.expander("⚙️ Asignación de Columnas del Archivo", expanded=(archivo_subido is not None)):
    st.markdown("Confirma qué columna corresponde a cada dato:")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        sel_punto = st.selectbox("Columna de Punto / ID", cols_disponibles, index=cols_disponibles.index(col_pto_def) if col_pto_def in cols_disponibles else 0)
    with c2:
        sel_norte = st.selectbox("Columna Norte (Y)", cols_disponibles, index=cols_disponibles.index(col_nor_def) if col_nor_def in cols_disponibles else 0)
    with c3:
        sel_este = st.selectbox("Columna Este (X)", cols_disponibles, index=cols_disponibles.index(col_est_def) if col_est_def in cols_disponibles else 0)
    with c4:
        opciones_colindante = ["(Ninguna / Asignar automáticamente)"] + cols_disponibles
        index_colindante = (cols_disponibles.index(col_col_def) + 1) if (col_col_def and col_col_def in cols_disponibles and col_col_def not in [sel_punto, sel_norte, sel_este]) else 0
        sel_colindante = st.selectbox("Columna Colindante (Opcional)", opciones_colindante, index=index_colindante)

df_procesado = pd.DataFrame()
df_procesado['Punto'] = df_raw[sel_punto].astype(str)
df_procesado['Norte'] = df_raw[sel_norte].apply(limpiar_numero)
df_procesado['Este'] = df_raw[sel_este].apply(limpiar_numero)

if sel_colindante != "(Ninguna / Asignar automáticamente)":
    df_procesado['Colindante'] = df_raw[sel_colindante].fillna("").astype(str)
else:
    df_procesado['Colindante'] = "Colindancia según levantamiento"

st.markdown("#### Vista previa y edición de puntos:")
df_editor = st.data_editor(df_procesado, num_rows="dynamic", use_container_width=True)

if st.button("🚀 Generar Minuta y Cálculos", type="primary"):
    if len(df_editor) < 3:
        st.error("Se requieren al menos 3 vértices para conformar un polígono y calcular linderos.")
    else:
        tramos = []
        perimetro_total = 0.0
        n_puntos = len(df_editor)
        
        n_vals = df_editor['Norte'].values
        e_vals = df_editor['Este'].values

        for i in range(n_puntos):
            sig = (i + 1) % n_puntos
            p_origen = str(df_editor.iloc[i]['Punto'])
            n1 = float(n_vals[i])
            e1 = float(e_vals[i])
            colindante = str(df_editor.iloc[i]['Colindante']).strip()
            if not colindante:
                colindante = "colindancia según predio vecino"
            
            p_destino = str(df_editor.iloc[sig]['Punto'])
            n2 = float(n_vals[sig])
            e2 = float(e_vals[sig])
            
            dist = calcular_distancia(n1, e1, n2, e2)
            perimetro_total += dist
            g, m, s, az = calcular_azimut(n1, e1, n2, e2)
            rumbo_str = calcular_rumbo(az)
            
            tramos.append({
                "Origen": p_origen,
                "Destino": p_destino,
                "Norte_Orig": n1,
                "Este_Orig": e1,
                "Distancia_m": dist,
                "Azimut": f"{g}°{m:02d}'{s:04.1f}\"",
                "Rumbo": rumbo_str,
                "Colindante": colindante
            })

        df_tramos = pd.DataFrame(tramos)
        area_m2 = calcular_area_gauss(n_vals, e_vals)
        hectareas = int(area_m2 // 10000)
        metros_restantes = round(area_m2 % 10000, 2)
        fanegadas = area_m2 / 6400.0

        st.success("✅ Cálculos realizados exitosamente.")

        m1, m2, m3 = st.columns(3)
        m1.metric("Área en Hectáreas", f"{hectareas} Has + {metros_restantes:,.2f} m²", f"{area_m2:,.2f} m² total")
        m2.metric("Perímetro Total", f"{perimetro_total:,.2f} m")
        m3.metric("Fanegadas / Plazas", f"{fanegadas:.2f} Fg")

        texto_minuta = (
            f"DESCRIPCIÓN TÉCNICA Y DETERMINACIÓN DE LINDEROS\n\n"
            f"PREDIO: {nombre_predio.upper()}\n"
            f"UBICACIÓN: Vereda {vereda}, Municipio de {municipio}, Departamento de {departamento}.\n"
            f"PROPIETARIO / SOLICITANTE: {propietario.upper()}\n"
            f"CÉDULA CATASTRAL: {cedula_catastral}\n"
            f"MATRÍCULA INMOBILIARIA: {matricula_inmobiliaria}\n"
            f"ÁREA TOTAL: {area_m2:,.2f} M² ({hectareas} Has. + {metros_restantes:,.2f} M² / {fanegadas:.2f} Fanegadas).\n"
            f"PERÍMETRO: {perimetro_total:,.2f} Metros.\n\n"
            f"ALINDERACIÓN TÉCNICA:\n"
            f"El inmueble se delimita a partir del Punto {tramos[0]['Origen']} con coordenadas de origen "
            f"Norte: {tramos[0]['Norte_Orig']:,.2f} m y Este: {tramos[0]['Este_Orig']:,.2f} m. "
        )

        for t in tramos:
            texto_minuta += (
                f"Del Punto {t['Origen']} continúa en línea recta con rumbo {t['Rumbo']} (Azimut {t['Azimut']}) "
                f"y una distancia de {t['Distancia_m']:.2f} metros hasta llegar al Punto {t['Destino']}, "
                f"colindando con {t['Colindante']}. "
            )

        texto_minuta += (
            f"Punto donde encierra y cierra el polígono.\n\n"
            f"Levantamiento y memoria técnica elaborada por:\n"
            f"{profesional.upper()}\n"
            f"Matrícula Profesional No: {matricula_prof}\n"
        )

        st.subheader("📄 Minuta Técnica Generada")
        st.text_area("Texto listo para copiar o insertar en escrituras/informes periciales:", value=texto_minuta, height=280)

        st.subheader("📊 Cuadro Técnico de Coordenadas, Azimuts y Distancias")
        st.dataframe(df_tramos[["Origen", "Destino", "Distancia_m", "Azimut", "Rumbo", "Colindante"]], use_container_width=True)

        # Generar archivo Word (.docx)
        doc = Document()
        
        titulo = doc.add_heading("MEMORIA TÉCNICA Y MINUTA DE LINDEROS", 0)
        titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        p_sub = doc.add_paragraph(f"PREDIO: {nombre_predio.upper()}\nMunicipio de {municipio} ({departamento}) - Vereda {vereda}")
        p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        doc.add_heading("1. INFORMACIÓN GENERAL", level=1)
        doc.add_paragraph(f"• Propietario / Solicitante: {propietario}")
        doc.add_paragraph(f"• Cédula Catastral: {cedula_catastral}")
        doc.add_paragraph(f"• Matrícula Inmobiliaria: {matricula_inmobiliaria}")
        doc.add_paragraph(f"• Área Calculada: {area_m2:,.2f} m² ({hectareas} Has + {metros_restantes:,.2f} m²)")
        doc.add_paragraph(f"• Perímetro Total: {perimetro_total:,.2f} m")

        doc.add_heading("2. DESCRIPCIÓN TÉCNICA DE LINDEROS", level=1)
        doc.add_paragraph(texto_minuta)

        doc.add_heading("3. CUADRO DE TRAMOS, DISTANCIAS Y RUMBOS", level=1)
        tabla = doc.add_table(rows=1, cols=6)
        tabla.style = 'Table Grid'
        hdr_cells = tabla.rows[0].cells
        hdr_cells[0].text = 'Origen'
        hdr_cells.text = 'Destino'
        hdr_cells.text = 'Distancia (m)'
        hdr_cells.text = 'Azimut'
        hdr_cells.text = 'Rumbo'
        hdr_cells.text = 'Colindancia'
        
        for t in tramos:
            row_cells = tabla.add_row().cells
            row_cells[0].text = str(t['Origen'])
            row_cells.text = str(t['Destino'])
            row_cells.text = f"{t['Distancia_m']:.2f}"
            row_cells.text = str(t['Azimut'])
            row_cells.text = str(t['Rumbo'])
            row_cells.text = str(t['Colindante'])

        doc.add_paragraph(f"\n\n____________________________________\n{profesional.upper()}\nTopógrafo / Perito\nMatrícula Profesional: {matricula_prof}")

        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        st.download_button(
            label="📥 Descargar Informe Completo en Word (.docx)",
            data=buffer,
            file_name=f"Minuta_{nombre_predio.replace(' ', '_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
