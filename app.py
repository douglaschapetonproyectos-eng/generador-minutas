import streamlit as st
import pandas as pd
import math
import re
import io
from io import BytesIO
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from PIL import Image

try:
    import pypdf
    HAY_PYPDF = True
except:
    HAY_PYPDF = False

st.set_page_config(page_title="Generador de Minutas Topográficas", layout="wide", page_icon="📐")

def limpiar_numero(val):
    if pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    if not s:
        return 0.0
    if '.' in s and ',' in s:
        if s.rfind(',') > s.rfind('.'):
            s = s.replace('.', '').replace(',', '.')
        else:
            s = s.replace(',', '')
    elif ',' in s:
        s = s.replace(',', '.')
    s_clean = re.sub(r'[^\d.-]', '', s)
    try:
        return float(s_clean)
    except:
        return 0.0

def calcular_distancia(n1, e1, n2, e2):
    return math.sqrt((n2 - n1)**2 + (e2 - e1)**2)

def obtener_rumbo_cardinal(n1, e1, n2, e2):
    dn = n2 - n1
    de = e2 - e1
    if abs(dn) < 1e-7 and abs(de) < 1e-7:
        return "Mismo punto"
    rad = math.atan2(de, dn)
    grados = math.degrees(rad)
    if grados < 0:
        grados += 360
    
    # 8 sectores cardinales
    if 337.5 <= grados or grados < 22.5:
        return "Norte"
    elif 22.5 <= grados < 67.5:
        return "Nor-Oriente"
    elif 67.5 <= grados < 112.5:
        return "Oriente"
    elif 112.5 <= grados < 157.5:
        return "Sur-Oriente"
    elif 157.5 <= grados < 202.5:
        return "Sur"
    elif 202.5 <= grados < 247.5:
        return "Sur-Occidente"
    elif 247.5 <= grados < 292.5:
        return "Occidente"
    else:
        return "Nor-Occidente"

def calcular_area_gauss(n, e):
    num_puntos = len(n)
    if num_puntos < 3:
        return 0.0
    area = 0.0
    for i in range(num_puntos):
        j = (i + 1) % num_puntos
        area += e[i] * n[j] - e[j] * n[i]
    return abs(area) / 2.0

def obtener_tramos_entre(p_inicio, p_fin, df):
    df_clean = df.reset_index(drop=True)
    n_p = len(df_clean)
    if n_p < 2:
        return []
    
    pts_list = [str(x).strip() for x in df_clean['Punto'].tolist()]
    p_inicio_str = str(p_inicio).strip()
    p_fin_str = str(p_fin).strip()
    
    if p_inicio_str not in pts_list or p_fin_str not in pts_list:
        return []
    
    idx_ini = pts_list.index(p_inicio_str)
    idx_fin = pts_list.index(p_fin_str)
    
    if idx_ini == idx_fin:
        return []
    
    seq_indices = []
    curr = idx_ini
    while True:
        seq_indices.append(curr)
        if curr == idx_fin:
            break
        curr = (curr + 1) % n_p
        if curr == idx_ini:
            break
            
    tramos_costado = []
    for k in range(len(seq_indices) - 1):
        i1 = seq_indices[k]
        i2 = seq_indices[k+1]
        p1 = str(df_clean.iloc[i1]['Punto']).strip()
        n1 = float(limpiar_numero(df_clean.iloc[i1]['Norte']))
        e1 = float(limpiar_numero(df_clean.iloc[i1]['Este']))
        
        p2 = str(df_clean.iloc[i2]['Punto']).strip()
        n2 = float(limpiar_numero(df_clean.iloc[i2]['Norte']))
        e2 = float(limpiar_numero(df_clean.iloc[i2]['Este']))
        
        dist_calc = calcular_distancia(n1, e1, n2, e2)
        if 'Distancia' in df_clean.columns and pd.notna(df_clean.iloc[i1]['Distancia']) and limpiar_numero(df_clean.iloc[i1]['Distancia']) > 0:
            dist = float(limpiar_numero(df_clean.iloc[i1]['Distancia']))
        else:
            dist = dist_calc
            
        rumbo_cardinal = obtener_rumbo_cardinal(n1, e1, n2, e2)
        
        tramos_costado.append({
            "origen": p1, "destino": p2,
            "n1": n1, "e1": e1, "n2": n2, "e2": e2,
            "dist": dist, "rumbo": rumbo_cardinal
        })
    return tramos_costado

def redactar_costado(nombre_costado, p_ini, p_fin, colindante, elemento_lindero, df):
    tramos = obtener_tramos_entre(p_ini, p_fin, df)
    if not tramos:
        return f"POR EL {nombre_costado}: Tramo no especificado o puntos no válidos."
    
    dist_total = sum(t['dist'] for t in tramos)
    p_orig = tramos[0]
    p_dest = tramos[-1]
    
    tipo_linea = "línea recta" if len(tramos) == 1 else "línea quebrada"
    colind_txt = str(colindante).strip() if str(colindante).strip() else "predio colindante según levantamiento"
    elem_txt = str(elemento_lindero).strip() if str(elemento_lindero).strip() else "límite material"
    
    txt = f"POR EL {nombre_costado}: Inicia en el Punto {p_orig['origen']} de coordenadas (Norte: {p_orig['n1']:,.2f} m, Este: {p_orig['e1']:,.2f} m), "
    txt += f"continúa en {tipo_linea} "
    
    if len(tramos) == 1:
        txt += f"en sentido {p_orig['rumbo']} en una distancia de {dist_total:.2f} metros "
    else:
        puntos_intermedios = [f"Punto {t['destino']} (N: {t['n2']:,.2f} m, E: {t['e2']:,.2f} m, sentido {t['rumbo']}, distancia de {t['dist']:.2f} m)" for t in tramos[:-1]]
        if puntos_intermedios:
            txt += f"pasando por {', '.join(puntos_intermedios)}, con una distancia total acumulada de {dist_total:.2f} metros "
        else:
            txt += f"con una distancia total de {dist_total:.2f} metros "
            
    txt += f"hasta llegar al Punto {p_dest['destino']}, colindando con {colind_txt}, teniendo como elemento delimitador {elem_txt}."
    return txt

st.title("📐 Generador Automatizado de Minutas Topográficas y Linderos")
st.markdown("Herramienta técnica para la generación de minutas periciales y notariales, cálculo de polígonos, colindancias y anexo de planos.")

# 1. Datos Generales
with st.expander("📝 1. Datos Generales del Predio y Profesional", expanded=True):
    col1, col2, col3 = st.columns(3)
    with col1:
        nombre_predio = st.text_input("Nombre del Predio / Lote", value="LOTE EL RUBY")
        municipio = st.text_input("Municipio", value="Guatavita")
        vereda = st.text_input("Vereda", value="Centro")
    with col2:
        departamento = st.text_input("Departamento", value="Cundinamarca")
        cedula_catastral = st.text_input("Cédula Catastral", value="25-000-00-00-0000-0000-000")
        matricula_inmobiliaria = st.text_input("Matrícula Inmobiliaria", value="150-00000")
    with col3:
        propietario = st.text_input("Propietario / Demandante / Solicitante", value="CARLOS CESAR YANCE REDONDO")
        profesional = st.text_input("Topógrafo / Perito", value="DOUGLAS CHAPETON GOMEZ")
        matricula_prof = st.text_input("Matrícula Profesional CPNT", value="01-19914 CPNT")

# 2. Coordenadas
st.markdown("### 📍 2. Coordenadas del Polígono (Excel / CSV o Tabla)")
archivo_coords = st.file_uploader("Cargar archivo Excel (.xlsx, .xls) o CSV (Columnas: Punto, Norte, Este, Distancia opcional)", type=["xlsx", "xls", "csv"], key="uploader_coords")

datos_defecto = {
    "Punto": ["1", "2", "3", "4", "5", "6", "7", "8"],
    "Norte": [2119344.68, 2119347.01, 2119344.82, 2119340.35, 2119247.20, 2119236.88, 2119209.69, 2119292.31],
    "Este": [4840129.36, 4840156.42, 4840160.43, 4840162.03, 4840144.94, 4840135.55, 4840068.98, 4840086.06],
    "Distancia": [27.16, 4.57, 4.75, 94.70, 13.95, 71.91, 84.37, 67.95]
}

if archivo_coords is not None:
    try:
        if archivo_coords.name.endswith('.csv'):
            df_raw = pd.read_csv(archivo_coords)
        else:
            df_raw = pd.read_excel(archivo_coords)
        st.success(f"✅ Archivo '{archivo_coords.name}' cargado con {len(df_raw)} puntos.")
    except Exception as e:
        st.error(f"Error al leer archivo: {e}")
        df_raw = pd.DataFrame(datos_defecto)
else:
    df_raw = pd.DataFrame(datos_defecto)

cols = list(df_raw.columns)
def col_match(lista, ops):
    for c in lista:
        c_l = str(c).lower().strip()
        for o in ops:
            if o in c_l:
                return c
    return lista[0] if lista else None

c_pto = col_match(cols, ['punto', 'pto', 'id', 'name', 'vertice', 'est', 'item', 'no'])
c_nor = col_match(cols, ['norte', 'north', 'lat', 'y'])
c_est = col_match(cols, ['este', 'east', 'lon', 'x'])
c_dist = col_match(cols, ['distancia', 'dist', 'longitud_tramo', 'dist_m'])

with st.expander("⚙️ Asignación de Columnas de Coordenadas"):
    sc1, sc2, sc3, sc4 = st.columns(4)
    sel_p = sc1.selectbox("Columna de Puntos", cols, index=cols.index(c_pto) if c_pto in cols else 0)
    sel_n = sc2.selectbox("Columna Norte", cols, index=cols.index(c_nor) if c_nor in cols else 0)
    sel_e = sc3.selectbox("Columna Este", cols, index=cols.index(c_est) if c_est in cols else 0)
    opciones_dist = ["(Calcular automáticamente de coordenadas)"] + cols
    index_dist = (cols.index(c_dist) + 1) if (c_dist and c_dist in cols and c_dist not in [sel_p, sel_n, sel_e]) else 0
    sel_dist = sc4.selectbox("Columna Distancia (Opcional)", opciones_dist, index=index_dist)

df_pts = pd.DataFrame()
df_pts['Punto'] = df_raw[sel_p].astype(str)
df_pts['Norte'] = df_raw[sel_n].apply(limpiar_numero)
df_pts['Este'] = df_raw[sel_e].apply(limpiar_numero)

if sel_dist != "(Calcular automáticamente de coordenadas)":
    df_pts['Distancia'] = df_raw[sel_dist].apply(limpiar_numero)

st.markdown("#### Tabla de Coordenadas del Polígono (Editable):")
df_editor = st.data_editor(df_pts, num_rows="dynamic", use_container_width=True)

lista_puntos_actuales = [str(x).strip() for x in df_editor['Punto'].tolist() if str(x).strip()]
if len(lista_puntos_actuales) < 2:
    lista_puntos_actuales = ["1", "2", "3", "4"]

# 3. Configuración de Linderos
st.markdown("### 🧭 3. Configuración de Colindancias por Costados Cardinales")
st.caption("Selecciona los mojones iniciales y finales para cada costado y el elemento delimitador material.")

opciones_elementos = [
    "cerca de alambre al medio",
    "cerca viva al medio",
    "muro en ladrillo al medio",
    "muro en piedra / mampostería al medio",
    "vía pública / carretera pavimentada al medio",
    "camino veredal al medio",
    "servidumbre de acceso al medio",
    "quebrada aguas arriba",
    "quebrada aguas abajo",
    "río aguas arriba",
    "río aguas abajo",
    "mojones de concreto y línea imaginaria",
    "límite natural según levantamiento"
]

# NORTE
with st.container():
    st.markdown("#### 🔵 Lindero Norte")
    n1, n2, n3, n4 = st.columns(4)
    pto_ini_norte = n1.selectbox("Desde Punto", lista_puntos_actuales, index=0, key="ini_norte")
    pto_fin_norte = n2.selectbox("Hasta Punto", lista_puntos_actuales, index=min(3, len(lista_puntos_actuales)-1), key="fin_norte")
    colind_norte = n3.text_input("Colinda con:", value="Predio Lote 2 de María González", key="col_norte")
    elem_norte = n4.selectbox("Elemento Delimitador", opciones_elementos, index=0, key="elem_norte")

# ORIENTE
with st.container():
    st.markdown("#### 🟢 Lindero Oriente")
    o1, o2, o3, o4 = st.columns(4)
    pto_ini_oriente = o1.selectbox("Desde Punto", lista_puntos_actuales, index=min(3, len(lista_puntos_actuales)-1), key="ini_oriente")
    pto_fin_oriente = o2.selectbox("Hasta Punto", lista_puntos_actuales, index=min(4, len(lista_puntos_actuales)-1), key="fin_oriente")
    colind_oriente = o3.text_input("Colinda con:", value="Servidumbre de acceso veredal", key="col_oriente")
    elem_oriente = o4.selectbox("Elemento Delimitador", opciones_elementos, index=6, key="elem_oriente")

# SUR
with st.container():
    st.markdown("#### 🟡 Lindero Sur")
    s1, s2, s3, s4 = st.columns(4)
    pto_ini_sur = s1.selectbox("Desde Punto", lista_puntos_actuales, index=min(4, len(lista_puntos_actuales)-1), key="ini_sur")
    pto_fin_sur = s2.selectbox("Hasta Punto", lista_puntos_actuales, index=min(6, len(lista_puntos_actuales)-1), key="fin_sur")
    colind_sur = s3.text_input("Colinda con:", value="Hacienda El Roble de Pedro Gómez", key="col_sur")
    elem_sur = s4.selectbox("Elemento Delimitador", opciones_elementos, index=8, key="elem_sur")

# OCCIDENTE
with st.container():
    st.markdown("#### 🔴 Lindero Occidente")
    w1, w2, w3, w4 = st.columns(4)
    pto_ini_occ = w1.selectbox("Desde Punto", lista_puntos_actuales, index=min(6, len(lista_puntos_actuales)-1), key="ini_occ")
    pto_fin_occ = w2.selectbox("Hasta Punto", lista_puntos_actuales, index=0, key="fin_occ")
    colind_occ = w3.text_input("Colinda con:", value="Predio Lote 4 de la misma subdivisión", key="col_occ")
    elem_occ = w4.selectbox("Elemento Delimitador", opciones_elementos, index=2, key="elem_occ")

# 4. Anexo de Planos e Imágenes / PDF
st.markdown("### 🖼️ 4. Anexar Imágenes de Planos o Fotografías (PNG / JPG / PDF)")
archivos_planos = st.file_uploader("Adjuntar archivos de planos o fotos de linderos", type=["png", "jpg", "jpeg", "pdf"], accept_multiple_files=True, key="uploader_planos")

imagenes_para_word = []
if archivos_planos:
    st.markdown("#### Archivos adjuntos detectados:")
    cols_img = st.columns(min(len(archivos_planos), 3))
    for idx, f in enumerate(archivos_planos):
        col_act = cols_img[idx % 3]
        if f.type in ["image/png", "image/jpeg", "image/jpg"]:
            try:
                img = Image.open(f)
                col_act.image(img, caption=f.name, use_container_width=True)
                imagenes_para_word.append((f.name, f.getvalue()))
            except Exception as e:
                col_act.error(f"Error en imagen: {e}")
        elif f.type == "application/pdf":
            col_act.info(f"📄 Archivo PDF: **{f.name}** ({f.size / 1024:.1f} KB)")
            if HAY_PYPDF:
                try:
                    pdf_reader = pypdf.PdfReader(io.BytesIO(f.getvalue()))
                    col_act.caption(f"Páginas: {len(pdf_reader.pages)}")
                except:
                    pass

# 5. Botón de Generación y Cálculos
st.markdown("---")
if st.button("🚀 Generar Minuta Técnica y Documento Word", type="primary"):
    df_editor_clean = df_editor.reset_index(drop=True)
    if len(df_editor_clean) < 3:
        st.error("Se requieren al menos 3 vértices para conformar un polígono y calcular linderos.")
    else:
        n_vals = [float(limpiar_numero(x)) for x in df_editor_clean['Norte'].values]
        e_vals = [float(limpiar_numero(x)) for x in df_editor_clean['Este'].values]
        n_puntos = len(df_editor_clean)

        # Cálculos de tramos generales
        tramos_completos = []
        perimetro_total = 0.0
        for i in range(n_puntos):
            sig = (i + 1) % n_puntos
            p1 = str(df_editor_clean.iloc[i]['Punto']).strip()
            n1 = n_vals[i]
            e1 = e_vals[i]
            p2 = str(df_editor_clean.iloc[sig]['Punto']).strip()
            n2 = n_vals[sig]
            e2 = e_vals[sig]
            
            dist_calc = calcular_distancia(n1, e1, n2, e2)
            if 'Distancia' in df_editor_clean.columns and pd.notna(df_editor_clean.iloc[i]['Distancia']) and limpiar_numero(df_editor_clean.iloc[i]['Distancia']) > 0:
                dist = float(limpiar_numero(df_editor_clean.iloc[i]['Distancia']))
            else:
                dist = dist_calc
                
            perimetro_total += dist
            rumbo_cardinal = obtener_rumbo_cardinal(n1, e1, n2, e2)
            
            tramos_completos.append({
                "Punto": p1,
                "Norte": n1,
                "Este": e1,
                "Destino": p2,
                "Distancia_m": dist,
                "Sentido_Rumbo": rumbo_cardinal
            })

        df_tabla_tramos = pd.DataFrame(tramos_completos)
        area_m2 = calcular_area_gauss(n_vals, e_vals)
        hectareas = int(area_m2 // 10000)
        metros_restantes = round(area_m2 % 10000, 2)
        fanegadas = area_m2 / 6400.0

        st.success("✅ Minuta y cálculos topográficos generados con éxito.")

        m1, m2, m3 = st.columns(3)
        m1.metric("Área en Hectáreas", f"{hectareas} Has + {metros_restantes:,.2f} m²", f"{area_m2:,.2f} m² totales")
        m2.metric("Perímetro Total", f"{perimetro_total:,.2f} Metros")
        m3.metric("Fanegadas / Plazas", f"{fanegadas:.2f} Fg")

        txt_norte = redactar_costado("NORTE", pto_ini_norte, pto_fin_norte, colind_norte, elem_norte, df_editor_clean)
        txt_oriente = redactar_costado("ORIENTE", pto_ini_oriente, pto_fin_oriente, colind_oriente, elem_oriente, df_editor_clean)
        txt_sur = redactar_costado("SUR", pto_ini_sur, pto_fin_sur, colind_sur, elem_sur, df_editor_clean)
        txt_occidente = redactar_costado("OCCIDENTE", pto_ini_occ, pto_fin_occ, colind_occ, elem_occ, df_editor_clean)

        texto_minuta_completa = (
            f"DESCRIPCIÓN TÉCNICA Y DETERMINACIÓN DE LINDEROS\n\n"
            f"PREDIO: {nombre_predio.upper()}\n"
            f"UBICACIÓN: Vereda {vereda}, Municipio de {municipio}, Departamento de {departamento}.\n"
            f"PROPIETARIO / SOLICITANTE: {propietario.upper()}\n"
            f"CÉDULA CATASTRAL: {cedula_catastral}\n"
            f"MATRÍCULA INMOBILIARIA: {matricula_inmobiliaria}\n"
            f"ÁREA TOTAL: {area_m2:,.2f} M² ({hectareas} Has. + {metros_restantes:,.2f} M² / {fanegadas:.2f} Fanegadas).\n"
            f"PERÍMETRO: {perimetro_total:,.2f} Metros.\n\n"
            f"LINDEROS Y ALINDERACIÓN:\n\n"
            f"{txt_norte}\n\n"
            f"{txt_oriente}\n\n"
            f"{txt_sur}\n\n"
            f"{txt_occidente}\n\n"
            f"Punto de partida y donde encierra el polígono.\n\n"
            f"Levantamiento y memoria técnica elaborada por:\n"
            f"{profesional.upper()}\n"
            f"Matrícula Profesional No: {matricula_prof}\n"
        )

        st.subheader("📄 Minuta Técnica Generada")
        st.text_area("Texto de la Minuta listo para copiar:", value=texto_minuta_completa, height=350)

        st.subheader("📊 Cuadro Técnico de Coordenadas, Distancias y Sentido")
        st.dataframe(df_tabla_tramos[["Punto", "Norte", "Este", "Destino", "Distancia_m", "Sentido_Rumbo"]], use_container_width=True)

        # Generar Documento Word (.docx)
        doc = Document()
        
        titulo = doc.add_heading("MEMORIA TÉCNICA Y DESCRIPCIÓN DE LINDEROS", 0)
        titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        p_sub = doc.add_paragraph(f"PREDIO: {nombre_predio.upper()}\nMunicipio de {municipio} ({departamento}) - Vereda {vereda}")
        p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        doc.add_heading("1. INFORMACIÓN GENERAL DEL INMUEBLE", level=1)
        doc.add_paragraph(f"• Propietario / Demandante: {propietario}")
        doc.add_paragraph(f"• Cédula Catastral: {cedula_catastral}")
        doc.add_paragraph(f"• Matrícula Inmobiliaria: {matricula_inmobiliaria}")
        doc.add_paragraph(f"• Área Superficial Total: {area_m2:,.2f} m² ({hectareas} Has + {metros_restantes:,.2f} m² / {fanegadas:.2f} Fanegadas)")
        doc.add_paragraph(f"• Perímetro Total del Predio: {perimetro_total:,.2f} m")

        doc.add_heading("2. DETERMINACIÓN Y DESCRIPCIÓN DE LINDEROS", level=1)
        doc.add_paragraph(f"El inmueble denominado {nombre_predio.upper()} se encuentra alinderado de la siguiente manera:\n")
        doc.add_paragraph(txt_norte)
        doc.add_paragraph(txt_oriente)
        doc.add_paragraph(txt_sur)
        doc.add_paragraph(txt_occidente)
        doc.add_paragraph("Punto de partida y donde encierra el polígono.")

        doc.add_heading("3. CUADRO TÉCNICO DE COORDENADAS Y TRAMOS", level=1)
        tabla = doc.add_table(rows=1, cols=6)
        tabla.style = 'Table Grid'
        hdr = tabla.rows[0].cells
        hdr[0].text = 'Punto'
        hdr.text = 'Norte (m)'
        hdr.text = 'Este (m)'
        hdr.text = 'Destino'
        hdr.text = 'Distancia (m)'
        hdr.text = 'Sentido / Rumbo'
        
        for t in tramos_completos:
            row = tabla.add_row().cells
            row[0].text = str(t['Punto'])
            row.text = f"{t['Norte']:,.2f}"
            row.text = f"{t['Este']:,.2f}"
            row.text = str(t['Destino'])
            row.text = f"{t['Distancia_m']:.2f}"
            row.text = str(t['Sentido_Rumbo'])

        # Anexo de Imágenes en Word
        if imagenes_para_word:
            doc.add_page_break()
            doc.add_heading("4. ANEXOS CARTOGRÁFICOS Y FOTOGRÁFICOS", level=1)
            for nombre_img, bytes_img in imagenes_para_word:
                doc.add_paragraph(f"Plano / Fotografía: {nombre_img}")
                try:
                    img_stream = io.BytesIO(bytes_img)
                    doc.add_picture(img_stream, width=Inches(5.5))
                except Exception as ex_img:
                    doc.add_paragraph(f"[No se pudo incrustar la imagen: {ex_img}]")
                doc.add_paragraph("")

        doc.add_paragraph(f"\n\n____________________________________\n{profesional.upper()}\nTopógrafo / Perito\nMatrícula Profesional No: {matricula_prof}")

        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        st.download_button(
            label="📥 Descargar Informe Completo en Word (.docx)",
            data=buffer,
            file_name=f"Minuta_{nombre_predio.replace(' ', '_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
