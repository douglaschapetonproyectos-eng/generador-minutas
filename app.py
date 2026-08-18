import streamlit as st
import pandas as pd
import math
from io import BytesIO
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

st.set_page_config(page_title="Generador de Minutas Topográficas", layout="wide")

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
    return f"{cuadrante[0]} {g}°{m:02d}'{s:04.1f}\" {cuadrante[1]}"

def calcular_area(df):
    # Fórmula de Gauss (Shoelace formula)
    n = df['Norte'].values
    e = df['Este'].values
    num_puntos = len(n)
    if num_puntos < 3:
        return 0.0
    area = 0.0
    for i in range(num_puntos):
        j = (i + 1) % num_puntos
        area += e[i] * n[j] - e[j] * n[i]
    return abs(area) / 2.0

st.title("📐 Generador Automatizado de Minutas Topográficas y Linderos")
st.markdown("Genera la descripción técnica y legal de linderos a partir de coordenadas y datos del predio.")

# Formulario de datos generales
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

# Ingreso de coordenadas
st.markdown("### 📍 2. Coordenadas y Linderos")
st.info("Ingresa los puntos en orden de recorrido perimetral (sentido horario o antihorario).")

# Opción de carga de archivo
archivo_subido = st.file_uploader("Opcional: Cargar archivo Excel o CSV (Columnas requeridas: Punto, Norte, Este, Colindante, Orientacion)", type=["xlsx", "csv"])

datos_base = {
    "Punto": [1, 2, 3, 4],
    "Norte": [2119344.68, 2119347.01, 2119247.20, 2119236.88],
    "Este": [4840129.36, 4840156.42, 4840144.94, 4840135.55],
    "Colindante": ["Servidumbre de acceso", "Predio Lote 2", "Camino Veredal", "Predio Lote 4"],
    "Orientacion": ["NORTE", "ORIENTE", "SUR", "OCCIDENTE"]
}

if archivo_subido is not None:
    try:
        if archivo_subido.name.endswith('.csv'):
            df_input = pd.read_csv(archivo_subido)
        else:
            df_input = pd.read_excel(archivo_subido)
    except Exception as e:
        st.error(f"Error al leer el archivo: {e}")
        df_input = pd.DataFrame(datos_base)
else:
    df_input = pd.DataFrame(datos_base)

df_puntos = st.data_editor(df_input, num_rows="dynamic", use_container_width=True)

if st.button("🚀 Generar Minuta y Cálculos", type="primary"):
    if len(df_puntos) < 3:
        st.error("Se requieren al menos 3 vértices para conformar un polígono.")
    else:
        # Cálculos de tramos
        tramos = []
        perimetro_total = 0.0
        n_puntos = len(df_puntos)

        for i in range(n_puntos):
            sig = (i + 1) % n_puntos
            p1 = df_puntos.iloc[i]
            p2 = df_puntos.iloc[sig]
            
            p_origen = str(p1['Punto'])
            n1 = float(p1['Norte'])
            e1 = float(p1['Este'])
            colindante = str(p1['Colindante']) if pd.notna(p1['Colindante']) else "Sin colindante"
            orientacion = str(p1['Orientacion']).upper() if 'Orientacion' in p1 and pd.notna(p1['Orientacion']) else ""
            
            p_destino = str(p2['Punto'])
            n2 = float(p2['Norte'])
            e2 = float(p2['Este'])
            
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
                "Colindante": colindante,
                "Orientacion": orientacion
            })

        df_tramos = pd.DataFrame(tramos)
        area_m2 = calcular_area(df_puntos)
        hectareas = int(area_m2 // 10000)
        metros_restantes = round(area_m2 % 10000, 2)

        # Generar texto de la minuta
        texto_minuta = f"DESCRIPCIÓN TÉCNICA Y DETERMINACIÓN DE LINDEROS\n"
        texto_minuta += f"PREDIO: {nombre_predio.upper()}\n"
        texto_minuta += f"UBICACIÓN: Vereda {vereda}, Municipio de {municipio}, Departamento de {departamento}.\n"
        texto_minuta += f"PROPIETARIO / SOLICITANTE: {propietario.upper()}\n"
        texto_minuta += f"CÉDULA CATASTRAL: {cedula_catastral}\n"
        texto_minuta += f"MATRÍCULA INMOBILIARIA: {matricula_inmobiliaria}\n"
        texto_minuta += f"ÁREA TOTAL: {area_m2:,.2f} m² ({hectareas} Has + {metros_restantes:,.2f} m²)\n"
        texto_minuta += f"PERÍMETRO: {perimetro_total:,.2f} m\n\n"
        texto_minuta += "ALINDERACIÓN:\n"
        
        texto_minuta += f"El predio se delimita a partir del Punto {tramos[0]['Origen']} de coordenadas Norte: {tramos[0]['Norte_Orig']:,.2f} m y Este: {tramos[0]['Este_Orig']:,.2f} m. "
        
        for t in tramos:
            texto_minuta += (
                f"Del Punto {t['Origen']} continúa en línea recta con rumbo {t['Rumbo']} "
                f"y una distancia de {t['Distancia_m']:.2f} metros hasta el Punto {t['Destino']} "
                f"(colindando con {t['Colindante']}). "
            )
        
        texto_minuta += f"Punto de partida y donde encierra el polígono.\n\n"
        texto_minuta += f"Levantamiento realizado por: {profesional} (M.P. {matricula_prof}).\n"

        st.subheader("📄 Minuta Técnica Generada")
        st.text_area("Texto de la Minuta", value=texto_minuta, height=280)

        st.subheader("📊 Cuadro de Coordenadas y Distancias")
        st.dataframe(df_tramos[["Origen", "Destino", "Distancia_m", "Azimut", "Rumbo", "Colindante"]], use_container_width=True)

        # Generación de archivo Word (.docx)
        doc = Document()
        doc.add_heading(f"INFORME TÉCNICO DE LINDEROS - {nombre_predio.upper()}", 0)
        
        doc.add_paragraph(f"Municipio: {municipio} ({departamento}) | Vereda: {vereda}")
        doc.add_paragraph(f"Cédula Catastral: {cedula_catastral} | Matrícula: {matricula_inmobiliaria}")
        doc.add_paragraph(f"Área: {area_m2:,.2f} m² ({hectareas} Has + {metros_restantes:,.2f} m²) | Perímetro: {perimetro_total:,.2f} m")
        
        doc.add_heading("1. Descripción de Linderos", level=1)
        doc.add_paragraph(texto_minuta)
        
        doc.add_heading("2. Cuadro Técnico de Coordenadas y Tramos", level=1)
        tabla = doc.add_table(rows=1, cols=6)
        hdr_cells = tabla.rows[0].cells
        hdr_cells[0].text = 'Origen'
        hdr_cells[1].text = 'Destino'
        hdr_cells[2].text = 'Distancia (m)'
        hdr_cells[3].text = 'Azimut'
        hdr_cells[4].text = 'Rumbo'
        hdr_cells[5].text = 'Colindante'
        
        for t in tramos:
            row_cells = tabla.add_row().cells
            row_cells[0].text = str(t['Origen'])
            row_cells[1].text = str(t['Destino'])
            row_cells[2].text = f"{t['Distancia_m']:.2f}"
            row_cells[3].text = str(t['Azimut'])
            row_cells[4].text = str(t['Rumbo'])
            row_cells[5].text = str(t['Colindante'])
            
        doc.add_paragraph(f"\nTopógrafo / Perito: {profesional}\nLicencia / Matrícula: {matricula_prof}")

        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        st.download_button(
            label="📥 Descargar Minuta en Word (.docx)",
            data=buffer,
            file_name=f"Minuta_{nombre_predio.replace(' ', '_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )