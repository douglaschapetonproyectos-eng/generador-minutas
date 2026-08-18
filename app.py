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

def numero_a_letras(n):
    unidades = ["", "UN", "DOS", "TRES", "CUATRO", "CINCO", "SEIS", "SIETE", "OCHO", "NUEVE"]
    dieces = ["DIEZ", "ONCE", "DOCE", "TRECE", "CATORCE", "QUINCE", "DIECISÉIS", "DIECISIETE", "DIECIOCHO", "DIECINUEVE"]
    veintes = ["VEINTE", "VEINTIÚN", "VEINTIDÓS", "VEINTITRÉS", "VEINTICUATRO", "VEINTICINCO", "VEINTISÉIS", "VEINTISIETE", "VEINTIOCHO", "VEINTINUEVE"]
    decenas = ["", "DIEZ", "VEINTE", "TREINTA", "CUARENTA", "CINCUENTA", "SESENTA", "SETENTA", "OCHENTA", "NOVENTA"]
    centenas = ["", "CIENTO", "DOSCIENTOS", "TRESCIENTOS", "CUATROCIENTOS", "QUINIENTOS", "SEISCIENTOS", "SETECIENTOS", "OCHOCIENTOS", "NOVECIENTOS"]

    n_int = int(round(n))
    if n_int == 0:
        return "CERO"
    if n_int == 100:
        return "CIEN"

    def convertir_centena(num):
        if num == 0:
            return ""
        if num == 100:
            return "CIEN"
        c = num // 100
        d = (num % 100) // 10
        u = num % 10
        
        texto_c = centenas[c]
        resto = num % 100
        
        if resto == 0:
            return texto_c
        
        texto_resto = ""
        if 10 <= resto < 20:
            texto_resto = dieces[resto - 10]
        elif 20 <= resto < 30:
            texto_resto = veintes[resto - 20]
        else:
            if d > 0:
                texto_resto = decenas[d]
                if u > 0:
                    texto_resto += " Y " + unidades[u]
            else:
                texto_resto = unidades[u]
                
        if texto_c:
            return f"{texto_c} {texto_resto}".strip()
        return texto_resto.strip()

    millones = n_int // 1000000
    miles = (n_int % 1000000) // 1000
    resto = n_int % 1000

    partes = []
    if millones > 0:
        if millones == 1:
            partes.append("UN MILLÓN")
        else:
            partes.append(convertir_centena(millones) + " MILLONES")
    
    if miles > 0:
        if miles == 1:
            partes.append("MIL")
        else:
            partes.append(convertir_centena(miles) + " MIL")
            
    if resto > 0:
        partes.append(convertir_centena(resto))

    return " ".join(partes).strip()

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
    
    if 337.5 <= grados or grados < 22.5:
        return "Norte"
    elif 22.5 <= grados < 67.5:
        return "Nor-Oriental"
    elif 67.5 <= grados < 112.5:
        return "Oriental"
    elif 112.5 <= grados < 157.5:
        return "Sur-Oriental"
    elif 157.5 <= grados < 202.5:
        return "Sur"
    elif 202.5 <= grados < 247.5:
        return "Sur-Occidental"
    elif 247.5 <= grados < 292.5:
        return "Occidental"
    else:
        return "Nor-Occidental"

def calcular_area_gauss(n, e):
    num_puntos = len(n)
    if num_puntos < 3:
        return 0.0
    area = 0.0
    for i in range(num_puntos):
        j = (i + 1) % num_puntos
        area += e[i] * n[j] - e[j] * n[i]
    return abs(area) / 2.0

def obtener_tramos_entre(p_inicio, p_fin, df, es_occidente=False):
    df_clean = df.reset_index(drop=True)
    n_p = len(df_clean)
    if n_p < 2:
        return []
    
    pts_list = [str(x).strip() for x in df_clean['Punto'].tolist()]
    p_inicio_str = str(p_inicio).strip()
    p_fin_str = str(p_fin).strip()
    
    if p_inicio_str not in pts_list:
        return []
    
    idx_ini = pts_list.index(p_inicio_str)
    seq_indices = []
    curr = idx_ini
    
    if es_occidente:
        while curr < n_p:
            seq_indices.append(curr)
            curr += 1
        seq_indices.append(0)  # Cierra volviendo al mojón 1
    else:
        if p_fin_str not in pts_list:
            return []
        idx_fin = pts_list.index(p_fin_str)
        if idx_ini == idx_fin:
            return []
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

def redactar_lindero_formato(nombre_costado, num_lindero, p_ini, p_fin, colindante, elemento_lindero, df, es_occidente=False):
    tramos = obtener_tramos_entre(p_ini, p_fin, df, es_occidente=es_occidente)
    if not tramos:
        return f"Por el {nombre_costado}: lindero {num_lindero}: no se pudo delimitar el tramo."
    
    dist_total = sum(t['dist'] for t in tramos)
    p_orig = tramos[0]
    p_dest = tramos[-1]
    
    tipo_linea = "en línea semi-recta" if len(tramos) == 1 else "en línea quebrada"
    sentido_gral = p_orig['rumbo']
    colind_txt = str(colindante).strip() if str(colindante).strip() else "predio vecino"
    elem_txt = str(elemento_lindero).strip() if str(elemento_lindero).strip() else "cerca de alambre al medio"
    
    if not elem_txt.endswith("al medio") and not elem_txt.startswith("quebrada") and not elem_txt.startswith("río") and not elem_txt.startswith("camino"):
        elem_txt_completo = f"{elem_txt} al medio."
    else:
        elem_txt_completo = f"{elem_txt}."

    txt = f"Por el {nombre_costado}: lindero {num_lindero}: inicia en el mojón {p_orig['origen']} (N: {p_orig['n1']:.4f}, E: {p_orig['e1']:.4f}); "
    txt += f"en sentido {sentido_gral} {tipo_linea}, "
    
    if len(tramos) == 1:
        if es_occidente:
            txt += f"regresando hasta el mojón {p_dest['destino']} (N: {p_dest['n2']:.4f}, E: {p_dest['e2']:.4f}) punto de partida y encierra, "
        else:
            txt += f"hasta el mojón {p_dest['destino']} (N: {p_dest['n2']:.4f}, E: {p_dest['e2']:.4f}); "
    else:
        for t in tramos[:-1]:
            txt += f"hasta el mojón {t['destino']} (N: {t['n2']:.4f}, E: {t['e2']:.4f}); "
        
        if es_occidente:
            txt += f"regresando hasta el mojón {p_dest['destino']} (N: {p_dest['n2']:.4f}, E: {p_dest['e2']:.4f}) punto de partida y encierra, "
        else:
            txt += f"hasta el mojón {p_dest['destino']} (N: {p_dest['n2']:.4f}, E: {p_dest['e2']:.4f}); "
            
    txt += f"colindando con el {colind_txt} {elem_txt_completo} teniendo como distancia total para este lindero de {dist_total:.2f} mts."
    return txt

st.title("📐 Generador Automatizado de Minutas Topográficas y Linderos")
st.markdown("Herramienta técnica conforme al **Art. 2.2.2.2.17 del Decreto 148 de 2020** (Sistema MAGNA-SIRGAS Origen Único EPSG 9377).")

# 1. Datos Generales
with st.expander("📝 1. Identificación del Inmueble y Jurisdicción Registral", expanded=True):
    col1, col2, col3 = st.columns(3)
    with col1:
        nombre_predio = st.text_input("Nombre del Predio", value="SAN AGUSTIN")
        departamento = st.text_input("Departamento", value="Cundinamarca")
        municipio = st.text_input("Municipio", value="San Francisco")
    with col2:
        vereda = st.text_input("Vereda / Sector", value="El Arrayan")
        cedula_catastral = st.text_input("Número Predial / Cédula Catastral", value="256580000000000020232000000000")
        matricula_inmobiliaria = st.text_input("Matrícula Inmobiliaria", value="156-47224")
    with col3:
        circuito_registral = st.text_input("Circuito Registral / ORIP", value="Facatativá")
        profesional = st.text_input("Topógrafo / Perito", value="DOUGLAS CHAPETON GOMEZ")
        matricula_prof = st.text_input("Matrícula Profesional CPNT", value="01-19914 CPNT")

# 2. Coordenadas
st.markdown("### 📍 2. Coordenadas del Polígono")
st.info("Carga tu archivo Excel o CSV con las columnas: **Punto**, **Norte**, **Este** (y **Distancia** opcional).")

archivo_coords = st.file_uploader("Cargar archivo Excel (.xlsx, .xls) o CSV con coordenadas", type=["xlsx", "xls", "csv"], key="uploader_coords")

df_base_inicial = pd.DataFrame(columns=["Punto", "Norte", "Este", "Distancia"])

if archivo_coords is not None:
    try:
        if archivo_coords.name.endswith('.csv'):
            df_raw = pd.read_csv(archivo_coords)
        else:
            df_raw = pd.read_excel(archivo_coords)
        
        cols = list(df_raw.columns)
        def col_match(lista, ops):
            for c in lista:
                c_l = str(c).lower().strip()
                for o in ops:
                    if o in c_l:
                        return c
            return lista[0] if lista else None

        c_pto = col_match(cols, ['punto', 'pto', 'id', 'name', 'vertice', 'est', 'item', 'no', 'mojon'])
        c_nor = col_match(cols, ['norte', 'north', 'lat', 'y'])
        c_est = col_match(cols, ['este', 'east', 'lon', 'x'])
        c_dist = col_match(cols, ['distancia', 'dist', 'longitud', 'dist_m'])

        with st.expander("⚙️ Asignación de Columnas del Archivo"):
            sc1, sc2, sc3, sc4 = st.columns(4)
            sel_p = sc1.selectbox("Columna Punto / Mojón", cols, index=cols.index(c_pto) if c_pto in cols else 0)
            sel_n = sc2.selectbox("Columna Norte", cols, index=cols.index(c_nor) if c_nor in cols else 0)
            sel_e = sc3.selectbox("Columna Este", cols, index=cols.index(c_est) if c_est in cols else 0)
            opciones_dist = ["(Calcular automáticamente)"] + cols
            index_dist = (cols.index(c_dist) + 1) if (c_dist and c_dist in cols and c_dist not in [sel_p, sel_n, sel_e]) else 0
            sel_dist = sc4.selectbox("Columna Distancia (Opcional)", opciones_dist, index=index_dist)

        df_pts = pd.DataFrame()
        df_pts['Punto'] = df_raw[sel_p].astype(str)
        df_pts['Norte'] = df_raw[sel_n].apply(limpiar_numero)
        df_pts['Este'] = df_raw[sel_e].apply(limpiar_numero)
        if sel_dist != "(Calcular automáticamente)":
            df_pts['Distancia'] = df_raw[sel_dist].apply(limpiar_numero)
            
        st.success(f"✅ Archivo '{archivo_coords.name}' cargado con éxito ({len(df_pts)} puntos detectados).")
    except Exception as e:
        st.error(f"Error al leer archivo: {e}")
        df_pts = df_base_inicial
else:
    df_pts = df_base_inicial

st.markdown("#### Tabla de Coordenadas (Ingresa o edita los puntos):")
df_editor = st.data_editor(df_pts, num_rows="dynamic", use_container_width=True)

lista_puntos = [str(x).strip() for x in df_editor['Punto'].tolist() if str(x).strip()]

# 3. Configuración de Linderos
st.markdown("### 🧭 3. Configuración de Colindancias por Costados Cardinales")

opciones_elementos = [
    "cerca de alambre al medio",
    "cerca viva al medio",
    "muro en ladrillo al medio",
    "muro en piedra / mampostería al medio",
    "vía pública al medio",
    "camino veredal al medio",
    "servidumbre de acceso al medio",
    "quebrada aguas arriba",
    "quebrada aguas abajo",
    "río aguas arriba",
    "río aguas abajo",
    "mojones de concreto y línea imaginaria",
    "límite natural según levantamiento"
]

if len(lista_puntos) >= 3:
    # NORTE
    with st.container():
        st.markdown("#### 🔵 Costado Norte (Lindero 1)")
        n1, n2, n3, n4 = st.columns(4)
        pto_ini_norte = n1.selectbox("Inicia en Mojón", lista_puntos, index=0, key="ini_norte")
        pto_fin_norte = n2.selectbox("Hasta Mojón", lista_puntos, index=min(1, len(lista_puntos)-1), key="fin_norte")
        colind_norte = n3.text_input("Colinda con el predio:", value="Predio 00 00 0002 0233 000", key="col_norte")
        elem_norte = n4.selectbox("Elemento Delimitador", opciones_elementos, index=0, key="elem_norte")

    # ORIENTE
    with st.container():
        st.markdown("#### 🟢 Costado Oriente (Lindero 2)")
        o1, o2, o3, o4 = st.columns(4)
        pto_ini_oriente = o1.selectbox("Inicia en Mojón", lista_puntos, index=min(1, len(lista_puntos)-1), key="ini_oriente")
        pto_fin_oriente = o2.selectbox("Hasta Mojón", lista_puntos, index=min(2, len(lista_puntos)-1), key="fin_oriente")
        colind_oriente = o3.text_input("Colinda con el predio:", value="predio 00 00 0002 00608 000", key="col_oriente")
        elem_oriente = o4.selectbox("Elemento Delimitador", opciones_elementos, index=0, key="elem_oriente")

    # SUR
    with st.container():
        st.markdown("#### 🟡 Costado Sur (Lindero 3)")
        s1, s2, s3, s4 = st.columns(4)
        pto_ini_sur = s1.selectbox("Inicia en Mojón", lista_puntos, index=min(2, len(lista_puntos)-1), key="ini_sur")
        pto_fin_sur = s2.selectbox("Hasta Mojón", lista_puntos, index=min(len(lista_puntos)-2, len(lista_puntos)-1), key="fin_sur")
        colind_sur = s3.text_input("Colinda con el predio:", value="predio 00 00 0002 0131 000", key="col_sur")
        elem_sur = s4.selectbox("Elemento Delimitador", opciones_elementos, index=0, key="elem_sur")

    # OCCIDENTE
    with st.container():
        st.markdown("#### 🔴 Costado Occidente (Lindero 4 / Cierre al Mojón 1)")
        w1, w2, w3 = st.columns(3)
        pto_ini_occ = w1.selectbox("Inicia en Mojón", lista_puntos, index=min(len(lista_puntos)-2, len(lista_puntos)-1), key="ini_occ")
        colind_occ = w2.text_input("Colinda con el predio:", value="predio 00 01 0004 0164 000", key="col_occ")
        elem_occ = w3.selectbox("Elemento Delimitador", opciones_elementos, index=0, key="elem_occ")
else:
    st.info("Ingresa o carga al menos 3 puntos en la tabla superior para configurar los linderos por costados.")

# 4. Anexo de Planos e Imágenes / PDF
st.markdown("### 🖼️ 4. Anexar Imágenes de Planos o Fotografías (PNG / JPG / PDF)")
archivos_planos = st.file_uploader("Adjuntar planos o fotos de linderos", type=["png", "jpg", "jpeg", "pdf"], accept_multiple_files=True, key="uploader_planos")

imagenes_para_word = []
if archivos_planos:
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

# 5. Generación de la Minuta
st.markdown("---")
if st.button("🚀 Generar Minuta Técnica Oficial y Documento Word", type="primary"):
    df_editor_clean = df_editor.dropna(subset=['Punto', 'Norte', 'Este']).reset_index(drop=True)
    if len(df_editor_clean) < 3:
        st.error("Se requieren al menos 3 vértices con coordenadas válidas para conformar el polígono.")
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
        area_en_letras = numero_a_letras(area_m2)

        st.success("✅ Minuta técnica redactada y cálculos generados con éxito.")

        m1, m2, m3 = st.columns(3)
        m1.metric("Área en Hectáreas", f"{hectareas} ha + {metros_restantes:,.0f} M²", f"{area_m2:,.2f} m² totales")
        m2.metric("Perímetro Total", f"{perimetro_total:,.2f} Metros")
        m3.metric("Área en Letras", f"{area_en_letras}")

        # Redacción de linderos
        txt_norte = redactar_lindero_formato("Norte", "1", pto_ini_norte, pto_fin_norte, colind_norte, elem_norte, df_editor_clean)
        txt_oriente = redactar_lindero_formato("Oriente", "2", pto_ini_oriente, pto_fin_oriente, colind_oriente, elem_oriente, df_editor_clean)
        txt_sur = redactar_lindero_formato("SUR", "3", pto_ini_sur, pto_fin_sur, colind_sur, elem_sur, df_editor_clean)
        txt_occidente = redactar_lindero_formato("OCCIDENTE", "4", pto_ini_occ, "", colind_occ, elem_occ, df_editor_clean, es_occidente=True)

        encabezado_minuta = (
            f"MINUTA TECNICA DE LINDEROS\n\n"
            f"Corresponde a un inmueble identificado con el numero predial {cedula_catastral} ubicado en la zona Rural Vereda {vereda} del Municipio de {municipio}-{departamento}, denominado {nombre_predio.upper()} e inscrito en el circuito registral de {circuito_registral} bajo matricula inmobiliaria No {matricula_inmobiliaria}\n\n"
            f"La descripción técnica de los linderos del inmueble, en observancia del artículo 2.2.2.2.17 del decreto 148 de 2020, estableciendo la forma, orientación y extensión de los linderos en los siguientes términos:\n\n"
            f"“el bien inmueble identificado catastralmente con el numero predial {cedula_catastral} y folio de matrícula inmobiliaria {matricula_inmobiliaria}, denominado “{nombre_predio.upper()}”, presenta los siguientes linderos referidos al sistema de coordenadas proyectadas magna sirgas origen único nacional con epsg 9377:\n\n"
        )

        cuerpo_linderos = f"{txt_norte}\n\n{txt_oriente}\n\n{txt_sur}\n\n{txt_occidente}\n\n"
        
        cierre_area = f"De acuerdo con los anteriores linderos, el área del citado bien inmueble es de {area_m2:,.0f} metros cuadrados o {area_en_letras} metros cuadrados ({hectareas} ha + {metros_restantes:,.0f} M2)."

        texto_minuta_oficial = encabezado_minuta + cuerpo_linderos + cierre_area

        st.subheader("📄 Minuta Técnica Oficial Generada")
        st.text_area("Texto oficial listo para copiar:", value=texto_minuta_oficial, height=380)

        st.subheader("📊 Cuadro Técnico de Coordenadas, Distancias y Sentido")
        st.dataframe(df_tabla_tramos[["Punto", "Norte", "Este", "Destino", "Distancia_m", "Sentido_Rumbo"]], use_container_width=True)

        # Generar Documento Word (.docx)
        doc = Document()
        
        titulo = doc.add_heading("MINUTA TECNICA DE LINDEROS", 0)
        titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        p1_doc = doc.add_paragraph(
            f"Corresponde a un inmueble identificado con el numero predial {cedula_catastral} ubicado en la zona Rural Vereda {vereda} del Municipio de {municipio}-{departamento}, denominado {nombre_predio.upper()} e inscrito en el circuito registral de {circuito_registral} bajo matricula inmobiliaria No {matricula_inmobiliaria}\n"
        )
        
        p2_doc = doc.add_paragraph(
            "La descripción técnica de los linderos del inmueble, en observancia del artículo 2.2.2.2.17 del decreto 148 de 2020, estableciendo la forma, orientación y extensión de los linderos en los siguientes términos:\n"
        )

        p3_doc = doc.add_paragraph(
            f"“el bien inmueble identificado catastralmente con el numero predial {cedula_catastral} y folio de matrícula inmobiliaria {matricula_inmobiliaria}, denominado “{nombre_predio.upper()}”, presenta los siguientes linderos referidos al sistema de coordenadas proyectadas magna sirgas origen único nacional con epsg 9377:\n"
        )

        doc.add_paragraph(txt_norte)
        doc.add_paragraph(txt_oriente)
        doc.add_paragraph(txt_sur)
        doc.add_paragraph(txt_occidente)
        
        p_cierre = doc.add_paragraph(f"\n{cierre_area}")
        p_cierre.runs[0].bold = True

        doc.add_heading("CUADRO TÉCNICO DE COORDENADAS Y DISTANCIAS", level=1)
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
            row.text = f"{t['Norte']:,.4f}"
            row.text = f"{t['Este']:,.4f}"
            row.text = str(t['Destino'])
            row.text = f"{t['Distancia_m']:.2f}"
            row.text = str(t['Sentido_Rumbo'])

        # Anexos
        if imagenes_para_word:
            doc.add_page_break()
            doc.add_heading("ANEXOS CARTOGRÁFICOS Y FOTOGRÁFICOS", level=1)
            for nombre_img, bytes_img in imagenes_para_word:
                doc.add_paragraph(f"Plano / Fotografía: {nombre_img}")
                try:
                    img_stream = io.BytesIO(bytes_img)
                    doc.add_picture(img_stream, width=Inches(5.8))
                except Exception as ex_img:
                    doc.add_paragraph(f"[No se pudo incrustar imagen: {ex_img}]")
                doc.add_paragraph("")

        doc.add_paragraph(f"\n\n____________________________________\n{profesional.upper()}\nTopógrafo / Perito\nMatrícula Profesional No: {matricula_prof}")

        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        st.download_button(
            label="📥 Descargar Minuta Oficial en Word (.docx)",
            data=buffer,
            file_name=f"Minuta_{nombre_predio.replace(' ', '_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
