# -*- coding: utf-8 -*-
import pypdf
import pandas as pd
import re
import unicodedata
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
import io
import datetime
import streamlit as st

# ==========================================
# CONFIGURACIÓN DE PÁGINA
# ==========================================
st.set_page_config(
    page_title="Sistema FEFO - Hospital Vicente Corral Moscoso",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Estilos CSS Profesionales e Institucionales (Diseño Ejecutivo Moderno)
st.markdown("""
    <style>
    /* Ocultar barra lateral y header nativo de Streamlit */
    [data-testid="stSidebar"] {
        display: none !important;
    }
    header, .stApp > header {
        display: none !important;
        visibility: hidden !important;
        height: 0px !important;
    }
    /* Eliminar paddings y márgenes superiores de la página principal */
    .main .block-container {
        padding-top: 0rem !important;
        padding-bottom: 2rem !important;
        padding-left: 0rem !important;
        padding-right: 0rem !important;
        max-width: 100% !important;
        background-color: transparent !important;
    }
    /* Posicionamiento de las pestañas */
    .stTabs {
        position: fixed !important;
        top: 12px !important;
        left: 20px !important;
        z-index: 99999 !important;
        width: auto !important;
        background: transparent !important;
    }
    /* Barra contenedora con diseño glassmorphism y color institucional */
    .stTabs [data-baseweb="tab-list"],
    div[data-baseweb="tab-list"] {
        background: rgba(255, 255, 255, 0.90) !important;
        background-color: rgba(255, 255, 255, 0.90) !important;
        border-radius: 10px !important;
        padding: 4px 8px !important;
        border-bottom: none !important;
        gap: 8px !important;
        backdrop-filter: blur(12px) !important;
        box-shadow: 0 4px 20px rgba(31, 78, 120, 0.12) !important;
        border: 1px solid rgba(227, 242, 253, 0.8) !important;
    }
    div[data-baseweb="tab-highlight"] {
        background-color: transparent !important;
    }
    /* Estilos generales de las pestañas */
    .stTabs [data-baseweb="tab"] {
        height: 36px !important;
        background-color: transparent !important;
        border-radius: 8px !important;
        padding: 0px 14px !important;
        font-size: 0.85rem !important;
        font-weight: 600 !important;
        color: #1F4E78 !important;
        border: none !important;
        box-shadow: none !important;
        transition: all 0.2s ease-in-out;
    }
    
    .stTabs [data-baseweb="tab"]:hover {
        background-color: rgba(31, 78, 120, 0.08) !important;
        color: #0d47a1 !important;
    }
    
    /* Primera pestaña convertida en el Logo del MSP */
    .stTabs [data-baseweb="tab"]:nth-child(1) {
        width: 55px !important;
        min-width: 55px !important;
        background-image: url('https://images.seeklogo.com/logo-png/23/2/ministerio-de-salud-publica-logo-png_seeklogo-231476.png') !important;
        background-size: contain !important;
        background-repeat: no-repeat !important;
        background-position: center !important;
        color: transparent !important;
        font-size: 0px !important;
        padding: 0 !important;
    }
    
    .stTabs [aria-selected="true"] {
        background-color: #1F4E78 !important;
        color: white !important;
        box-shadow: 0 2px 8px rgba(31, 78, 120, 0.3) !important;
    }
    
    .stTabs [aria-selected="true"]:nth-child(1) {
        background-color: rgba(31, 78, 120, 0.15) !important;
        background-image: url('https://images.seeklogo.com/logo-png/23/2/ministerio-de-salud-publica-logo-png_seeklogo-231476.png') !important;
        background-size: contain !important;
        background-repeat: no-repeat !important;
        background-position: center !important;
    }

    /* Fondo de pantalla fijo cubriendo toda la pantalla */
    .hero-fullscreen {
        position: fixed;
        top: 0;
        left: 0;
        width: 100vw;
        height: 100vh;
        background: linear-gradient(135deg, #f0f7ff 0%, #e3f2fd 50%, #ffffff 100%);
        display: flex;
        align-items: center;
        justify-content: space-between;
        overflow: hidden;
        margin: 0 !important;
        z-index: 0;
    }
    .hero-left-content {
        width: 48%;
        padding-left: 5%;
        padding-right: 3%;
        z-index: 1;
        display: flex;
        flex-direction: column;
        justify-content: center;
        transform: translateY(-15px);
    }
    .welcome-text {
        font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;
        font-size: 0.9rem;
        letter-spacing: 3px;
        color: #0d47a1;
        font-weight: 700;
        text-transform: uppercase;
        margin-bottom: 10px;
        display: inline-flex;
        align-items: center;
        gap: 6px;
    }
    .hero-title {
        font-size: 2.6rem;
        font-weight: 800;
        color: #1F4E78;
        line-height: 1.18;
        margin-bottom: 15px;
        text-shadow: 0 2px 10px rgba(0,0,0,0.03);
    }
    .hero-subtitle {
        font-size: 1.05rem;
        color: #555555;
        margin-bottom: 25px;
        line-height: 1.5;
        font-weight: 400;
    }
    .hero-right-image {
        width: 52%;
        height: 100%;
        position: absolute;
        right: 0;
        top: 0;
        background-image: url('https://www.metroecuador.com.ec/resizer/v2/JPWRHUMSIRGCDIHJTUBWDHGKJU.jpg?auth=5e0819a9b2036e60b75f706117b2939a852dfd313c77b0cc4feb0e1e423b24c0&width=1440&height=752');
        background-size: cover;
        background-position: center;
        -webkit-mask-image: linear-gradient(to right, rgba(0,0,0,0) 0%, rgba(0,0,0,0.85) 20%, rgba(0,0,0,1) 100%);
        mask-image: linear-gradient(to right, rgba(0,0,0,0) 0%, rgba(0,0,0,0.85) 20%, rgba(0,0,0,1) 100%);
        z-index: 1;
    }
    
    /* Contenedor flotante executive-ready para las pestañas */
    .content-overlay {
        position: relative;
        margin-top: 75px;
        margin-left: 25px;
        margin-right: 25px;
        padding: 35px 45px;
        z-index: 10;
        background: rgba(255, 255, 255, 0.88);
        backdrop-filter: blur(16px);
        border-radius: 16px;
        border: 1px solid rgba(255, 255, 255, 0.95);
        box-shadow: 0 12px 40px rgba(31, 78, 120, 0.08);
    }
    
    .stButton>button {
        width: 100%;
        background-color: #1F4E78;
        color: white;
        font-weight: 700;
        border-radius: 8px;
        padding: 0.65rem 1.2rem;
        border: none;
        box-shadow: 0 4px 12px rgba(31, 78, 120, 0.2);
        transition: all 0.2s ease;
    }
    .stButton>button:hover {
        background-color: #153450;
        color: white;
        box-shadow: 0 6px 16px rgba(31, 78, 120, 0.3);
        transform: translateY(-1px);
    }
    .stDownloadButton>button {
        background-color: #2e7d32;
        color: white;
        font-weight: 700;
        border-radius: 8px;
        padding: 0.65rem 1.2rem;
        border: none;
        box-shadow: 0 4px 12px rgba(46, 125, 50, 0.2);
    }
    .stDownloadButton>button:hover {
        background-color: #1b5e20;
    }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# MOTOR LÓGICO FEFO (INTACTO Y PROTEGIDO)
# ==========================================

def cargar_excel_seguro(file_bytes):
    try:
        return pd.read_excel(io.BytesIO(file_bytes), engine='openpyxl')
    except Exception:
        try:
            return pd.read_excel(io.BytesIO(file_bytes), engine='xlrd')
        except Exception:
            return pd.read_excel(io.BytesIO(file_bytes))

def clean_exact(s):
    if pd.isna(s):
        return ""
    if not isinstance(s, str):
        s = str(s)
    try:
        if isinstance(s, bytes):
            s = s.decode('utf-8', errors='ignore')
    except Exception:
        pass
    
    s = unicodedata.normalize('NFD', s).encode('ascii', 'ignore').decode("utf-8", errors='ignore')
    s = s.upper().strip()
    s = s.replace('P.V.', 'PV').replace('P.V', 'PV').replace('N°', 'N').replace('°', '')
    s = re.sub(r'[^A-Z0-9\s]', ' ', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s

def buscar_producto_inteligente(p_name_limpio, df_unificado, df_lotes_grouped):
    match = df_unificado[df_unificado['clean_name'] == p_name_limpio]
    if not match.empty:
        return match, p_name_limpio

    palabras = [w for w in p_name_limpio.split() if len(w) > 1 and w not in ['CON', 'DEL', 'LAS', 'LOS', 'POR']]
    if palabras:
        for _, row in df_unificado.iterrows():
            c_name = row['clean_name']
            if all(pal in c_name for pal in palabras):
                matched_df = df_unificado[df_unificado['clean_name'] == c_name]
                return matched_df, c_name

    return pd.DataFrame(), p_name_limpio

def consolidar_lotes_texto(texto_lotes_crudo):
    if not isinstance(texto_lotes_crudo, str) or not texto_lotes_crudo or 'SIN LOTE' in texto_lotes_crudo or 'FALTAN' in texto_lotes_crudo:
        return str(texto_lotes_crudo)
    
    segmentos = re.split(r'[/|]', texto_lotes_crudo)
    acumulador_lotes = {}
    
    for seg in segmentos:
        seg = seg.strip()
        if not seg:
            continue
        match = re.search(r'^(.*?)\s*\((\d+)\s*unids?\)', seg, re.IGNORECASE)
        if match:
            lote_nombre = match.group(1).strip()
            cantidad = int(match.group(2))
            acumulador_lotes[lote_nombre] = acumulador_lotes.get(lote_nombre, 0) + cantidad
        else:
            acumulador_lotes[seg] = acumulador_lotes.get(seg, 0)
            
    resultado_final = []
    for lote, cant in acumulador_lotes.items():
        if cant > 0:
            resultado_final.append(f"{lote} ({cant} unids)")
        else:
            resultado_final.append(lote)
            
    return " / ".join(resultado_final)

def procesar_cruce_hospitalario_lote(lista_pdfs, excel_lotes_bytes, excel_unificado_bytes, firmantes_custom=None):
    try:
        df_lotes = cargar_excel_seguro(excel_lotes_bytes)
        df_unificado = cargar_excel_seguro(excel_unificado_bytes)

        df_lotes['Producto'] = df_lotes['Producto'].astype(str)
        df_unificado['Producto'] = df_unificado['Producto'].astype(str)

        df_lotes['clean_name'] = df_lotes['Producto'].apply(clean_exact)
        df_unificado['clean_name'] = df_unificado['Producto'].apply(clean_exact)

        df_lotes_grouped = df_lotes.groupby(['clean_name', 'Lote', 'F. Cad'], as_index=False)['Existen'].sum()
        df_lotes_grouped['F. Cad DT'] = pd.to_datetime(df_lotes_grouped['F. Cad'], errors='coerce')
        df_lotes_grouped = df_lotes_grouped.sort_values(by=['clean_name', 'F. Cad DT'], ascending=[True, True])

        stock_disponible_dict = {}
        for _, row in df_lotes_grouped.iterrows():
            key = (row['clean_name'], row['Lote'], row['F. Cad'])
            stock_disponible_dict[key] = float(row['Existen']) if pd.notna(row['Existen']) else 0.0

        nombres_servicios_raw = []
        for idx_pdf, (nombre_pdf, pdf_bytes) in enumerate(lista_pdfs):
            pdf_text = ""
            try:
                reader = pypdf.PdfReader(io.BytesIO(pdf_bytes))
                for page in reader.pages:
                    text_page = page.extract_text()
                    if text_page:
                        pdf_text += text_page + "\n"
            except Exception:
                pdf_text = ""

            base_servicio = f"SERVICIO {idx_pdf + 1}"
            match_serv = re.search(r'Servicio\s*[:]\s*([^\r\n]+)', pdf_text, re.IGNORECASE)
            if match_serv:
                raw_serv = match_serv.group(1).strip()
                for stop_word in ['FECHA', 'NRO', 'VERSIÓN', 'VERSION', 'CÓDIGO', 'CODIGO']:
                    if stop_word in raw_serv.upper():
                        raw_serv = raw_serv.upper().split(stop_word)[0].strip()
                if len(raw_serv) > 1:
                    base_servicio = raw_serv.upper().strip()
            nombres_servicios_raw.append((idx_pdf, nombre_pdf, pdf_bytes, base_servicio, pdf_text))

        conteo_servicios = {}
        asignaciones_servicios = []
        for idx_pdf, nombre_pdf, pdf_bytes, base_serv, pdf_text in nombres_servicios_raw:
            conteo_servicios[base_serv] = conteo_servicios.get(base_serv, 0) + 1
            freq = conteo_servicios[base_serv]
            if freq == 1:
                serv_final = base_serv
            else:
                serv_final = f"{base_serv} ({freq})"
            asignaciones_servicios.append((idx_pdf, nombre_pdf, pdf_bytes, serv_final, pdf_text))

        todos_los_resultados = []
        metadata_actas = [] 

        for idx_pdf, nombre_pdf, pdf_bytes, servicio_nombre, pdf_text in asignaciones_servicios:
            firmantes_encontrados = []
            texto_plano_continuo = re.sub(r'[\r\n]+', ' ', pdf_text)
            matches_firma = re.finditer(r'(?:Firmado\s+electr[oó]nicamente\s+por|por\s*[:])\s*([A-ZÁÉÍÓÚÑ\s]{5,50})', texto_plano_continuo, re.IGNORECASE)
            for m in matches_firma:
                candidato = m.group(1).strip()
                palabras = [w for w in candidato.split() if w.isalpha() and len(w) > 1 and w.upper() not in ['VALIDAR', 'UNICAMENTE', 'FIRMAEC', 'TIME', 'STAMPING', 'SECURITY', 'DATA', 'QR', 'POR', 'DE', 'DEL', 'LA', 'LAS', 'LOS', 'EL', 'Y']]
                if len(palabras) >= 2:
                    nombre_valido = " ".join(palabras[:4]).upper()
                    if nombre_valido not in firmantes_encontrados:
                        firmantes_encontrados.append(nombre_valido)

            if len(firmantes_encontrados) < 2:
                for linea in pdf_text.split('\n'):
                    l_up = linea.strip().upper()
                    if any(n in l_up for n in ['CANTOS ESPINOZA', 'DELEG SARI', 'YUQUILIMA PUMA', 'CABRERA ZENTENO', 'ROJAS BURI']):
                        if l_up not in firmantes_encontrados:
                            firmantes_encontrados.append(l_up)

            elaborado_por = firmantes_encontrados[0] if len(firmantes_encontrados) > 0 else "RESPONSABLE DE SERVICIO"
            validado_por = firmantes_encontrados[1] if len(firmantes_encontrados) > 1 else ("DANIEL GUILLERMO DELEG SARI" if "EMERGENCIA" in servicio_nombre.upper() else "XIMENA KARINA CABRERA ZENTENO")
            entregado_por = "GUARDALMACÉN / DESPACHADOR"

            if firmantes_custom:
                if isinstance(firmantes_custom, dict):
                    if nombre_pdf in firmantes_custom and isinstance(firmantes_custom[nombre_pdf], dict):
                        cfg = firmantes_custom[nombre_pdf]
                        elaborado_por = cfg.get('elaborado_por', elaborado_por)
                        validado_por = cfg.get('validado_por', validado_por)
                        entregado_por = cfg.get('entregado_por', entregado_por)
                    elif any(k in firmantes_custom for k in ['elaborado_por', 'validado_por', 'entregado_por']):
                        elaborado_por = firmantes_custom.get('elaborado_por', elaborado_por)
                        validado_por = firmantes_custom.get('validado_por', validado_por)
                        entregado_por = firmantes_custom.get('entregado_por', entregado_por)

            idx_tabla = pdf_text.find("Nro.")
            texto_items = pdf_text[idx_tabla:] if idx_tabla != -1 else pdf_text
            clean_blob = re.sub(r'\s+', ' ', texto_items)
            items_extraidos = []
            
            for i in range(1, 100):
                pattern_item = rf'(?:^|\s){i}\s+(.+?)\s*\(\s*AUTORIZADA[^\)]+\)\s*(\d+)\s+(\d+)'
                m = re.search(pattern_item, clean_blob)
                if m:
                    desc = m.group(1).strip()
                    sol = int(m.group(2))
                    aprob = int(m.group(3))
                    items_extraidos.append({
                        'Nro.': i,
                        'Dispositivo Médico': desc,
                        'Cantidades Solicitadas': sol,
                        'Cantidades Aprobadas': aprob
                    })
                    clean_blob = clean_blob.replace(m.group(0), "", 1)
                else:
                    if i > 3:
                        break

            if not items_extraidos:
                items_extraidos.append({
                    'Nro.': 1,
                    'Dispositivo Médico': 'REVISAR ARCHIVO PDF (SIN ITEMS DETECTADOS)',
                    'Cantidades Solicitadas': 0,
                    'Cantidades Aprobadas': 0
                })

            df_pdf = pd.DataFrame(items_extraidos)
            df_pdf['Nro.'] = range(1, len(df_pdf) + 1)
            df_pdf['clean_name'] = df_pdf['Dispositivo Médico'].apply(clean_exact)

            for _, pdf_row in df_pdf.iterrows():
                p_name_raw = pdf_row['clean_name']
                cant_pendiente = pdf_row['Cantidades Aprobadas']
                
                match_unif, _ = buscar_producto_inteligente(p_name_raw, df_unificado, df_lotes_grouped)

                if not match_unif.empty:
                    codigo = match_unif.iloc[0].get('Código', match_unif.iloc[0].get('Codigo', 'SIN-CODIGO'))
                    sku = match_unif.iloc[0].get('SKU', match_unif.iloc[0].get('Sku', match_unif.iloc[0].get('SKU.', 'SIN-SKU')))
                    
                    presentacion = "N/A"
                    for col_p in ['Presentacion', 'PRESENTACION', 'Presentación', 'PRESENTACIÓN']:
                        if col_p in match_unif.columns:
                            val_p = match_unif.iloc[0][col_p]
                            if pd.notna(val_p):
                                presentacion = str(val_p)
                                break
                    if presentacion == "N/A":
                        match_lote_temp = df_lotes[df_lotes['clean_name'] == p_name_raw]
                        for col_p in ['Presentacion', 'PRESENTACION', 'Presentación', 'PRESENTACIÓN']:
                            if not match_lote_temp.empty and col_p in match_lote_temp.columns:
                                val_p = match_lote_temp.iloc[0][col_p]
                                if pd.notna(val_p):
                                    presentacion = str(val_p)
                                    break

                    cols_posibles = [c for c in match_unif.columns if any(w in str(c).upper() for w in ['BOD', 'FARM', 'STOCK', 'EXIST', 'TOTAL'])]
                    col_bodega = cols_posibles[0] if cols_posibles else match_unif.columns[-1]
                    stock_bodega = match_unif.iloc[0][col_bodega]
                    if pd.isna(stock_bodega):
                        stock_bodega = 0.0
                else:
                    codigo = 'SIN-CODIGO'
                    sku = 'SIN-SKU'
                    presentacion = 'N/A'
                    stock_bodega = 0.0

                _, matched_key_lotes = buscar_producto_inteligente(p_name_raw, df_lotes_grouped.drop_duplicates(subset=['clean_name']), df_lotes_grouped)
                lotes_prod = df_lotes_grouped[df_lotes_grouped['clean_name'] == matched_key_lotes]

                if lotes_prod.empty or float(stock_bodega) <= 0:
                    item_data = {
                        'PDF Origen': nombre_pdf,
                        'Servicio': servicio_nombre,
                        'Nro.': pdf_row['Nro.'],
                        'Código': str(codigo),
                        'SKU': str(sku),
                        'Presentación': str(presentacion),
                        'Dispositivo Médico': pdf_row['Dispositivo Médico'],
                        'Cant. Solicitada': pdf_row['Cantidades Solicitadas'],
                        'Cant. Aprobada': pdf_row['Cantidades Aprobadas'],
                        'Lotes y Cantidades a Despachar': 'SIN LOTE DISPONIBLE (0 unids)',
                        'Fechas de Caducidad': 'S/F',
                        'Stock Bodega': float(stock_bodega),
                        'Alerta_Rojo': True
                    }
                else:
                    lotes_info = []
                    fechas_info = []
                    for _, l_row in lotes_prod.iterrows():
                        if cant_pendiente <= 0:
                            break
                        key = (matched_key_lotes, l_row['Lote'], l_row['F. Cad'])
                        disp_lote = stock_disponible_dict.get(key, 0.0)
                        if disp_lote <= 0:
                            continue
                        if disp_lote >= cant_pendiente:
                            cant_a_usar = cant_pendiente
                            stock_disponible_dict[key] -= cant_pendiente  
                            cant_pendiente = 0
                        else:
                            cant_a_usar = disp_lote
                            stock_disponible_dict[key] = 0  
                            cant_pendiente -= disp_lote
                        
                        lotes_info.append(f"{l_row['Lote']} ({int(cant_a_usar)} unids)")
                        fecha_fmt = l_row['F. Cad DT'].strftime('%d/%m/%Y') if pd.notna(l_row['F. Cad DT']) else str(l_row['F. Cad'])
                        fechas_info.append(f"{l_row['Lote']}: {fecha_fmt}")
                    
                    if cant_pendiente > 0:
                        lotes_info.append(f"FALTAN ({int(cant_pendiente)} unids)")
                        fechas_info.append("STOCK INSUFICIENTE")
                        
                    lotes_crudos_str = " / ".join(lotes_info)
                    lotes_consolidados_str = consolidar_lotes_texto(lotes_crudos_str)
                        
                    item_data = {
                        'PDF Origen': nombre_pdf,
                        'Servicio': servicio_nombre,
                        'Nro.': pdf_row['Nro.'],
                        'Código': str(codigo),
                        'SKU': str(sku),
                        'Presentación': str(presentacion),
                        'Dispositivo Médico': pdf_row['Dispositivo Médico'],
                        'Cant. Solicitada': pdf_row['Cantidades Solicitadas'],
                        'Cant. Aprobada': pdf_row['Cantidades Aprobadas'],
                        'Lotes y Cantidades a Despachar': lotes_consolidados_str,
                        'Fechas de Caducidad': " / ".join(fechas_info),
                        'Stock Bodega': float(stock_bodega),
                        'Alerta_Rojo': False
                    }
                
                todos_los_resultados.append(item_data)

            metadata_actas.append({
                'nombre_pdf': nombre_pdf,
                'servicio': servicio_nombre,
                'elaborado_por': elaborado_por,
                'validado_por': validado_por,
                'entregado_por': entregado_por
            })

        df_final = pd.DataFrame(todos_los_resultados)
        if df_final.empty:
            return None, None

        output_excel = io.BytesIO()
        wb = Workbook()
        ws = wb.active
        ws.title = "ConsolidadoPedidos"

        servicios_unicos = sorted(df_final['Servicio'].unique())
        
        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        red_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
        red_font = Font(name="Calibri", size=11, bold=True, color="C00000")
        normal_font = Font(name="Calibri", size=11)
        border_thin = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'), top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))

        ws.append([f"DETALLE DE PEDIDOS Y LOTES ASIGNADOS POR SERVICIO (Generado el {datetime.datetime.now().strftime('%Y-%m-%d')})"])
        ws.cell(row=1, column=1).font = Font(name="Calibri", size=12, bold=True, color="1F4E78")
        ws.append([]) 

        headers_matriz = ['DISPOSITIVO / INSUMO', 'CÓDIGO', 'SKU', 'PRESENTACIÓN', 'STOCK BODEGA', 'LOTE Y CADUCIDAD ASIGNADA (FEFO)'] + servicios_unicos + ['TOTAL APROBADO']
        ws.append(headers_matriz)

        for col_num, _ in enumerate(headers_matriz, 1):
            cell = ws.cell(row=3, column=col_num)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

        productos_unicos = df_final[['Código', 'SKU', 'Presentación', 'Dispositivo Médico', 'Stock Bodega']].drop_duplicates(subset=['Dispositivo Médico'])
        productos_unicos = productos_unicos.sort_values(by='Dispositivo Médico', ascending=True)

        for idx, p_row in productos_unicos.iterrows():
            prod_nombre = p_row['Dispositivo Médico']
            prod_codigo = p_row['Código']
            prod_sku = p_row['SKU']
            prod_pres = p_row['Presentación']
            prod_stock = p_row['Stock Bodega']
            
            df_prod_all = df_final[df_final['Dispositivo Médico'] == prod_nombre]
            
            lotes_brutos_totales = " / ".join(df_prod_all['Lotes y Cantidades a Despachar'].dropna().astype(str))
            lotes_str = consolidar_lotes_texto(lotes_brutos_totales)

            fila_datos = [prod_nombre, prod_codigo, prod_sku, prod_pres, prod_stock, lotes_str]
            
            total_aprob_prod = 0
            for serv in servicios_unicos:
                match_serv = df_prod_all[df_prod_all['Servicio'] == serv]
                if not match_serv.empty:
                    cant = match_serv['Cant. Aprobada'].sum()
                    fila_datos.append(cant)
                    total_aprob_prod += cant
                else:
                    fila_datos.append("")
            
            fila_datos.append(total_aprob_prod)
            ws.append(fila_datos)

            current_row = ws.max_row
            is_alert = (prod_stock <= 0) or ('FALTAN' in lotes_str) or ('SIN LOTE' in lotes_str) or ('REVISAR' in prod_nombre)

            for col_num in range(1, len(headers_matriz) + 1):
                cell = ws.cell(row=current_row, column=col_num)
                cell.border = border_thin
                if is_alert:
                    cell.fill = red_fill
                    cell.font = red_font
                else:
                    cell.font = normal_font

        wb.save(output_excel)
        output_excel.seek(0)

        output_pdf = io.BytesIO()
        doc = SimpleDocTemplate(output_pdf, pagesize=letter, rightMargin=18, leftMargin=18, topMargin=20, bottomMargin=20)
        elements = []
        
        styles = getSampleStyleSheet()
        title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontName='Helvetica-Bold', fontSize=10, alignment=1, textColor=colors.HexColor('#1F4E78'))
        subtitle_style = ParagraphStyle('SubStyle', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=7.5, alignment=1, textColor=colors.darkgray)
        meta_style = ParagraphStyle('MetaStyle', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8, alignment=0, textColor=colors.black)
        cell_style = ParagraphStyle('CellStyle', parent=styles['Normal'], fontName='Helvetica', fontSize=7)
        cell_center = ParagraphStyle('CellCenter', parent=styles['Normal'], fontName='Helvetica', fontSize=7, alignment=1)
        cell_red = ParagraphStyle('CellRed', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=7, textColor=colors.HexColor('#C00000'))
        cell_red_center = ParagraphStyle('CellRedCenter', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=7, textColor=colors.HexColor('#C00000'), alignment=1)
        sig_style = ParagraphStyle('SigStyle', parent=styles['Normal'], fontName='Helvetica', fontSize=6.5, alignment=1)

        for i, meta in enumerate(metadata_actas):
            p_name_actual = meta['nombre_pdf']
            df_acta_actual = df_final[df_final['PDF Origen'] == p_name_actual]
            if df_acta_actual.empty:
                continue

            elements.append(Paragraph("HOSPITAL PROVINCIAL GENERAL DOCENTE VICENTE CORRAL MOSCOSO", title_style))
            elements.append(Paragraph(f"ACTA OFICIAL DE DISPENSACIÓN Y CRUCE DE STOCK (FEFO) - Archivo: {p_name_actual}", subtitle_style))
            elements.append(Spacer(1, 4))

            elements.append(Paragraph(f"<b>SERVICIO:</b> {meta['servicio']} &nbsp;&nbsp;|&nbsp;&nbsp; <b>EMISIÓN:</b> {datetime.datetime.now().strftime('%d/%m/%Y %H:%M')}", meta_style))
            elements.append(Spacer(1, 4))

            table_data = [['Nro.', 'Código', 'Dispositivo Médico', 'Cant. Ap.', 'Lotes y Cantidades a Despachar', 'Fechas Caducidad', 'Stock']]
            
            for _, row in df_acta_actual.iterrows():
                is_red = row['Alerta_Rojo'] or (row['Stock Bodega'] <= 0)
                c_style = cell_red if is_red else cell_style
                c_style_ctr = cell_red_center if is_red else cell_center
                
                table_data.append([
                    str(row['Nro.']),
                    Paragraph(str(row['Código']), c_style_ctr),
                    Paragraph(str(row['Dispositivo Médico']), c_style),
                    str(int(row['Cant. Aprobada'])),
                    Paragraph(str(row['Lotes y Cantidades a Despachar']), c_style),
                    Paragraph(str(row['Fechas de Caducidad']), c_style),
                    Paragraph(f"{row['Stock Bodega']:,.0f}", c_style_ctr)
                ])

            t = Table(table_data, colWidths=[22, 50, 180, 38, 150, 75, 45])
            t.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1F4E78')),
                ('TEXTCOLOR', (0,0), (-1,0), colors.white),
                ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                ('FONTSIZE', (0,0), (-1,0), 7.5),
                ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('BOTTOMPADDING', (0,0), (-1,0), 3),
                ('TOPPADDING', (0,0), (-1,0), 3),
                ('GRID', (0,0), (-1,-1), 0.5, colors.lightgrey),
            ]))
            
            elements.append(t)
            elements.append(Spacer(1, 10))

            firma_data = [
                [
                    Paragraph(f"<b>ELABORÓ / RECIBIDO:</b><br/><br/><br/>_________________________________<br/><b>{meta['elaborado_por']}</b>", sig_style),
                    Paragraph(f"<b>ENTREGADO BODEGA:</b><br/><br/><br/>_________________________________<br/><b>{meta['entregado_por']}</b>", sig_style),
                    Paragraph(f"<b>VALIDÓ / AUTORIZÓ:</b><br/><br/><br/>_________________________________<br/><b>{meta['validado_por']}</b>", sig_style)
                ]
            ]
            
            t_firmas = Table(firma_data, colWidths=[185, 185, 185])
            t_firmas.setStyle(TableStyle([
                ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                ('VALIGN', (0,0), (-1,-1), 'TOP'),
                ('LEFTPADDING', (0,0), (-1,-1), 4),
                ('RIGHTPADDING', (0,0), (-1,-1), 4),
            ]))

            elements.append(t_firmas)

            if i < len(metadata_actas) - 1:
                elements.append(PageBreak())

        doc.build(elements)
        output_pdf.seek(0)

        return output_excel.getvalue(), output_pdf.getvalue()
    except Exception as e:
        st.error(f"Error interno en el procesamiento: {e}")
        return None, None

# ==========================================
# FONDO DE PANTALLA CONSTANTE
# ==========================================
st.markdown("""
    <div class="hero-fullscreen">
        <div class="hero-left-content">
            <div class="welcome-text">✨ SISTEMA INSTITUCIONAL</div>
            <div class="hero-title">Hospital Provincial General Docente Vicente Corral Moscoso</div>
            <div class="hero-subtitle">Gestión automatizada de cruce FEFO, lotes y generación de actas oficiales con trazabilidad estricta.</div>
        </div>
        <div class="hero-right-image"></div>
    </div>
""", unsafe_allow_html=True)

# ==========================================
# PESTAÑAS PRINCIPALES DE NAVEGACIÓN
# ==========================================
tab_inicio, tab_cruce, tab_bodega, tab_actas, tab_normativa = st.tabs([
    "msp", 
    "📊 Cruce de Lote", 
    "📦 Lotes & Bodega", 
    "📄 Actas Oficiales", 
    "⚖️ Normativa FEFO"
])

# ------------------------------------------
# PESTAÑA 0: INICIO
# ------------------------------------------
with tab_inicio:
    st.markdown('<div class="content-overlay">', unsafe_allow_html=True)
    st.markdown("### 🏥 Bienvenido al Sistema de Cruce FEFO y Lotes")
    st.markdown("Utilice las pestañas superiores para navegar entre el motor de cruce de pedidos, la consulta de stock en bodega, la generación de actas oficiales y la revisión de normativas.")
    st.markdown('</div>', unsafe_allow_html=True)

# ------------------------------------------
# PESTAÑA 1: CRUCE DE LOTE
# ------------------------------------------
with tab_cruce:
    st.markdown('<div class="content-overlay">', unsafe_allow_html=True)
    st.markdown("### 📊 Cruce de Lotes y Pedidos (FEFO)")
    st.write("Cargue los archivos PDF de pedidos de los servicios y los archivos Excel de control de inventario y lotes.")
    
    col1, col2 = st.columns(2)
    with col1:
        uploaded_pdfs = st.file_uploader("📂 Archivos PDF de Pedidos de Servicios", type=['pdf'], accept_multiple_files=True)
        excel_lotes = st.file_uploader("📊 Excel de Lotes y Caducidades", type=['xlsx', 'xls'])
    with col2:
        excel_unificado = st.file_uploader("📈 Excel Catálogo / Stock General", type=['xlsx', 'xls'])
        
    if st.button("🚀 Ejecutar Cruce FEFO"):
        if uploaded_pdfs and excel_lotes and excel_unificado:
            pdf_tuples = [(p.name, p.getvalue()) for p in uploaded_pdfs]
            excel_lotes_bytes = excel_lotes.getvalue()
            excel_unificado_bytes = excel_unificado.getvalue()
            
            with st.spinner("Procesando cruce de lotes bajo normativa FEFO..."):
                ex_bytes, pdf_bytes = procesar_cruce_hospitalario_lote(pdf_tuples, excel_lotes_bytes, excel_unificado_bytes)
                if ex_bytes and pdf_bytes:
                    st.success("¡Cruce completado con éxito!")
                    st.download_button("📥 Descargar Consolidado Excel", data=ex_bytes, file_name="Consolidado_FEFO.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    st.download_button("📥 Descargar Actas en PDF", data=pdf_bytes, file_name="Actas_Oficiales_FEFO.pdf", mime="application/pdf")
        else:
            st.warning("Por favor cargue todos los archivos requeridos (PDFs de pedidos, Excel de lotes y Excel unificado).")
    st.markdown('</div>', unsafe_allow_html=True)

# ------------------------------------------
# PESTAÑA 2: LOTES & BODEGA
# ------------------------------------------
with tab_bodega:
    st.markdown('<div class="content-overlay">', unsafe_allow_html=True)
    st.markdown("### 📦 Consulta de Inventario y Lotes en Bodega")
    st.write("Visualice y filtre el estado actual de las existencias, fechas de caducidad y trazabilidad por lote.")
    archivo_bodega = st.file_uploader("Cargar Archivo Excel de Bodega para Consulta", type=['xlsx', 'xls'], key="bodega_file")
    if archivo_bodega:
        df_b = cargar_excel_seguro(archivo_bodega.getvalue())
        st.dataframe(df_b, use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

# ------------------------------------------
# PESTAÑA 3: ACTAS OFICIALES
# ------------------------------------------
with tab_actas:
    st.markdown('<div class="content-overlay">', unsafe_allow_html=True)
    st.markdown("### 📄 Gestión de Actas Oficiales y Firmantes")
    st.write("Configure los firmantes responsables y genere los reportes formales listos para firma electrónica.")
    st.info("El sistema asigna automáticamente los firmantes detectados en los metadatos de los PDFs institucionales.")
    st.markdown('</div>', unsafe_allow_html=True)

# ------------------------------------------
# PESTAÑA 4: NORMATIVA FEFO
# ------------------------------------------
with tab_normativa:
    st.markdown('<div class="content-overlay">', unsafe_allow_html=True)
    st.markdown("### ⚖️ Normativa y Lineamientos FEFO")
    st.markdown("""
    * **Principio FEFO (First Expire, First Out):** Los productos con fecha de caducidad más cercana deben ser despachados y consumidos primero para minimizar la merma por caducidad.
    * **Trazabilidad Institucional:** Cada despacho y cruce de lote debe quedar respaldado mediante acta oficial firmada por los responsables de servicio, bodega y validación.
    * **Control de Stock Crítico:** Los ítems con stock insuficiente o sin lotes disponibles se resaltan automáticamente en alertas operativas.
    """)
    st.markdown('</div>', unsafe_allow_html=True)
