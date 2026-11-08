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
# CONFIGURACIÓN DE PÁGINA STREAMLIT
# ==========================================
st.set_page_config(
    page_title="Sistema FEFO | Hospital Vicente Corral Moscoso",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ==========================================
# DISEÑO Y ESTILOS CSS (ESTÉTICA INSTITUCIONAL & NAVBAR FIJA)
# ==========================================
st.markdown("""
<style>
    /* Fondo: Difuminado suave (~50%) en la izquierda y nítido a todo color en la derecha */
    .stApp {
        background: linear-gradient(to right, rgba(255, 255, 255, 0.55) 0%, rgba(255, 255, 255, 0.40) 45%, rgba(255, 255, 255, 0.05) 75%, rgba(255, 255, 255, 0) 100%), 
                    url("https://www.metroecuador.com.ec/resizer/v2/JPWRHUMSIRGCDIHJTUBWDHGKJU.jpg?auth=5e0819a9b2036e60b75f706117b2939a852dfd313c77b0cc4feb0e1e423b24c0&width=1440&height=752");
        background-size: cover;
        background-position: center;
        background-attachment: fixed;
    }

    /* Ocultar margen superior predeterminado y dar espacio para la barra fija */
    .block-container {
        padding-top: 4.2rem !important;
    }

    /* Ocultar elementos predeterminados de Streamlit */
    .st-emotion-cache-1wmy9hl, header {
        visibility: hidden;
    }

    /* Estilos del texto institucional derecho */
    .header-right-text {
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        color: #3f5973;
        font-size: 0.9rem;
        font-weight: 500;
        display: inline-block;
        margin-top: 8px;
        text-align: right;
    }

    /* Cabecera Principal */
    .hospital-header {
        background: linear-gradient(135deg, #163a5f 0%, #1f4e78 100%);
        padding: 24px 28px;
        border-radius: 12px;
        color: white;
        margin-top: 10px;
        margin-bottom: 25px;
        box-shadow: 0 6px 18px rgba(0, 0, 0, 0.12);
        border-left: 6px solid #007bff;
    }
    .hospital-header h1 {
        color: #ffffff !important;
        font-weight: 700;
        font-size: 1.85rem;
        margin: 0;
    }
    .hospital-header p {
        color: #e0effe;
        font-size: 0.95rem;
        margin-top: 6px;
        margin-bottom: 0;
    }

    /* Botón Ejecutar Estilizado */
    .stButton > button {
        background: linear-gradient(135deg, #1f4e78 0%, #163a5f 100%);
        color: white !important;
        border: none;
        border-radius: 10px;
        font-weight: 600;
        padding: 12px 24px;
        transition: all 0.3s ease;
        width: 100%;
        box-shadow: 0 4px 10px rgba(0,0,0,0.15);
    }
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 15px rgba(0,0,0,0.22);
    }
</style>
""", unsafe_allow_html=True)


# ==========================================
# BARRA SUPERIOR FIJA & NAVEGACIÓN EN LA MISMA LÍNEA
# ==========================================
col_nav, col_header_right = st.columns([6.8, 3.2])

with col_nav:
    selected_tab = st.radio(
        "Navegación",
        ["🏠 Inicio", "📚 Cruce de Lote", "📊 Lotes & Bodega", "📄 Actas Oficiales", "🛡️ Normativa FEFO"],
        horizontal=True,
        label_visibility="collapsed",
        key="main_navigation_menu"
    )

with col_header_right:
    if selected_tab == "📚 Cruce de Lote":
        c_btn, c_txt = st.columns([1.2, 1.8])
        with c_btn:
            if st.button("🔄 Reiniciar", key="btn_reset_page", use_container_width=True):
                keys_to_reset = [
                    'unique_pdfs_key', 'unique_lotes_key', 'unique_unificado_key', 
                    'excel_res', 'pdf_res', 'df_resultado', 'show_firmantes_config',
                    'df_lotes_raw', 'df_unif_raw', 'df_lotes_fefo_view', 'productos_pedidos_set', 'metadata_actas_store', 'firmantes_individuales_store', 'vista_acta_activa'
                ]
                for k in keys_to_reset:
                    if k in st.session_state:
                        del st.session_state[k]
                st.rerun()
        with c_txt:
            st.markdown('<div class="header-right-text">MSP Ecuador · Cuenca</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div style="text-align: right;" class="header-right-text">MSP Ecuador · Cuenca</div>', unsafe_allow_html=True)

st.markdown("<hr style='margin-top: 2px; margin-bottom: 15px; border-color: #e1e8f0;'>", unsafe_allow_html=True)


# ==========================================
# FUNCIONES LÓGICAS (MOTOR FEFO ORIGINAL INTACTO)
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

def obtener_firmantes_acta(meta):
    pdf_key = meta['nombre_pdf']
    if 'firmantes_individuales_store' in st.session_state and pdf_key in st.session_state['firmantes_individuales_store']:
        return st.session_state['firmantes_individuales_store'][pdf_key]
    return {
        'elaborado_por': meta['elaborado_por'],
        'entregado_por': meta['entregado_por'],
        'validado_por': meta['validado_por']
    }

def generar_pdf_individual_acta(meta, df_acta_actual):
    firmantes = obtener_firmantes_acta(meta)
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

    elements.append(Paragraph("HOSPITAL PROVINCIAL GENERAL DOCENTE VICENTE CORRAL MOSCOSO", title_style))
    elements.append(Paragraph(f"ACTA OFICIAL DE DISPENSACIÓN Y CRUCE DE STOCK (FEFO) - Archivo: {meta['nombre_pdf']}", subtitle_style))
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
            Paragraph(f"<b>ELABORÓ / RECIBIDO:</b><br/><br/><br/>_________________________________<br/><b>{firmantes['elaborado_por']}</b>", sig_style),
            Paragraph(f"<b>ENTREGADO BODEGA:</b><br/><br/><br/>_________________________________<br/><b>{firmantes['entregado_por']}</b>", sig_style),
            Paragraph(f"<b>VALIDÓ / AUTORIZÓ:</b><br/><br/><br/>_________________________________<br/><b>{firmantes['validado_por']}</b>", sig_style)
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
    doc.build(elements)
    output_pdf.seek(0)
    return output_pdf.getvalue()

def generar_pdf_consolidado_actualizado(metadata_list, df_res):
    pdf_gen_bytes = io.BytesIO()
    doc_gen = SimpleDocTemplate(pdf_gen_bytes, pagesize=letter, rightMargin=18, leftMargin=18, topMargin=20, bottomMargin=20)
    elements_gen = []
    
    styles_g = getSampleStyleSheet()
    t_style_g = ParagraphStyle('TSG', parent=styles_g['Heading1'], fontName='Helvetica-Bold', fontSize=10, alignment=1, textColor=colors.HexColor('#1F4E78'))
    sub_style_g = ParagraphStyle('SSG', parent=styles_g['Normal'], fontName='Helvetica-Bold', fontSize=7.5, alignment=1, textColor=colors.darkgray)
    meta_style_g = ParagraphStyle('MSG', parent=styles_g['Normal'], fontName='Helvetica-Bold', fontSize=8, alignment=0, textColor=colors.black)
    cell_style_g = ParagraphStyle('CSG', parent=styles_g['Normal'], fontName='Helvetica', fontSize=7)
    cell_center_g = ParagraphStyle('CCG', parent=styles_g['Normal'], fontName='Helvetica', fontSize=7, alignment=1)
    cell_red_g = ParagraphStyle('CRG', parent=styles_g['Normal'], fontName='Helvetica-Bold', fontSize=7, textColor=colors.HexColor('#C00000'))
    cell_red_center_g = ParagraphStyle('CRCG', parent=styles_g['Normal'], fontName='Helvetica-Bold', fontSize=7, textColor=colors.HexColor('#C00000'), alignment=1)
    sig_style_g = ParagraphStyle('SGG', parent=styles_g['Normal'], fontName='Helvetica', fontSize=6.5, alignment=1)

    for i, meta_g in enumerate(metadata_list):
        p_name_g = meta_g['nombre_pdf']
        df_acta_g = df_res[df_res['PDF Origen'] == p_name_g]
        if df_acta_g.empty:
            continue

        firmantes_g = obtener_firmantes_acta(meta_g)

        elements_gen.append(Paragraph("HOSPITAL PROVINCIAL GENERAL DOCENTE VICENTE CORRAL MOSCOSO", t_style_g))
        elements_gen.append(Paragraph(f"ACTA OFICIAL DE DISPENSACIÓN Y CRUCE DE STOCK (FEFO) - Archivo: {p_name_g}", sub_style_g))
        elements_gen.append(Spacer(1, 4))
        elements_gen.append(Paragraph(f"<b>SERVICIO:</b> {meta_g['servicio']} &nbsp;&nbsp;|&nbsp;&nbsp; <b>EMISIÓN:</b> {datetime.datetime.now().strftime('%d/%m/%Y %H:%M')}", meta_style_g))
        elements_gen.append(Spacer(1, 4))

        table_data_g = [['Nro.', 'Código', 'Dispositivo Médico', 'Cant. Ap.', 'Lotes y Cantidades a Despachar', 'Fechas Caducidad', 'Stock']]
        for _, row_g in df_acta_g.iterrows():
            is_red_g = row_g['Alerta_Rojo'] or (row_g['Stock Bodega'] <= 0)
            c_s = cell_red_g if is_red_g else cell_style_g
            c_s_c = cell_red_center_g if is_red_g else cell_center_g
            
            table_data_g.append([
                str(row_g['Nro.']),
                Paragraph(str(row_g['Código']), c_s_c),
                Paragraph(str(row_g['Dispositivo Médico']), c_s),
                str(int(row_g['Cant. Aprobada'])),
                Paragraph(str(row_g['Lotes y Cantidades a Despachar']), c_s),
                Paragraph(str(row_g['Fechas de Caducidad']), c_s),
                Paragraph(f"{row_g['Stock Bodega']:,.0f}", c_s_c)
            ])

        t_g = Table(table_data_g, colWidths=[22, 50, 180, 38, 150, 75, 45])
        t_g.setStyle(TableStyle([
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
        elements_gen.append(t_g)
        elements_gen.append(Spacer(1, 10))

        firma_data_g = [
            [
                Paragraph(f"<b>ELABORÓ / RECIBIDO:</b><br/><br/><br/>_________________________________<br/><b>{firmantes_g['elaborado_por']}</b>", sig_style_g),
                Paragraph(f"<b>ENTREGADO BODEGA:</b><br/><br/><br/>_________________________________<br/><b>{firmantes_g['entregado_por']}</b>", sig_style_g),
                Paragraph(f"<b>VALIDÓ / AUTORIZÓ:</b><br/><br/><br/>_________________________________<br/><b>{firmantes_g['validado_por']}</b>", sig_style_g)
            ]
        ]
        t_firmas_g = Table(firma_data_g, colWidths=[185, 185, 185])
        t_firmas_g.setStyle(TableStyle([
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'TOP'),
            ('LEFTPADDING', (0,0), (-1,-1), 4),
            ('RIGHTPADDING', (0,0), (-1,-1), 4),
        ]))
        elements_gen.append(t_firmas_g)

        if i < len(metadata_list) - 1:
            elements_gen.append(PageBreak())

    doc_gen.build(elements_gen)
    pdf_gen_bytes.seek(0)
    return pdf_gen_bytes.getvalue()

def procesar_cruce_hospitalario_lote(lista_pdfs, excel_lotes_bytes, excel_unificado_bytes, firmantes_custom=None):
    try:
        df_lotes = cargar_excel_seguro(excel_lotes_bytes)
        df_unificado = cargar_excel_seguro(excel_unificado_bytes)

        st.session_state['df_lotes_raw'] = df_lotes
        st.session_state['df_unif_raw'] = df_unificado

        df_lotes['Producto'] = df_lotes['Producto'].astype(str)
        df_unificado['Producto'] = df_unificado['Producto'].astype(str)

        df_lotes['clean_name'] = df_lotes['Producto'].apply(clean_exact)
        df_unificado['clean_name'] = df_unificado['Producto'].apply(clean_exact)

        df_lotes_grouped = df_lotes.groupby(['clean_name', 'Lote', 'F. Cad'], as_index=False)['Existen'].sum()
        df_lotes_grouped['F. Cad DT'] = pd.to_datetime(df_lotes_grouped['F. Cad'], errors='coerce')
        df_lotes_grouped = df_lotes_grouped.sort_values(by=['clean_name', 'F. Cad DT'], ascending=[True, True])

        map_nombres = df_lotes[['clean_name', 'Producto']].drop_duplicates(subset=['clean_name']).set_index('clean_name')['Producto'].to_dict()
        df_lotes_grouped['Producto'] = df_lotes_grouped['clean_name'].map(map_nombres)

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
        productos_pedidos_set = set()

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
            if idx_tabla == -1:
                idx_tabla = pdf_text.find("N°")
            if idx_tabla == -1:
                idx_tabla = 0

            texto_items = pdf_text[idx_tabla:] if idx_tabla != -1 else pdf_text
            clean_blob = re.sub(r'[\r\n]+', ' ', texto_items)
            items_extraidos = []
            
            for i in range(1, 100):
                pattern_item = rf'(?:^|\s){i}\s+(.+?)\s*\(\s*AUTORIZAD[AO][^\)]*\)\s*(\d+)\s+(\d+)'
                m = re.search(pattern_item, clean_blob, re.IGNORECASE)
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
                    pattern_alt = rf'(?:^|\s){i}\s+(.+?)\s+(\d+)\s+(\d+)'
                    m_alt = re.search(pattern_alt, clean_blob)
                    if m_alt and i > 2:
                        desc = m_alt.group(1).strip()
                        sol = int(m_alt.group(2))
                        aprob = int(m_alt.group(3))
                        if not any(w in desc.upper() for w in ['FIRMADO', 'ELECTRONICAMENTE', 'VALIDAR', 'STAMPING']):
                            items_extraidos.append({
                                'Nro.': i,
                                'Dispositivo Médico': desc,
                                'Cantidades Solicitadas': sol,
                                'Cantidades Aprobadas': aprob
                            })
                            clean_blob = clean_blob.replace(m_alt.group(0), "", 1)
                            continue

                    if i > 5 and not items_extraidos:
                        break
                    elif i > 40 and len(items_extraidos) > 0:
                        break

            if not items_extraidos:
                items_extraidos.append({
                    'Nro.': 1,
                    'Dispositivo Médico': f'ADVERTENCIA: REVISAR ARCHIVO {nombre_pdf}',
                    'Cantidades Solicitadas': 0,
                    'Cantidades Aprobadas': 0
                })

            df_pdf = pd.DataFrame(items_extraidos)
            df_pdf['Nro.'] = range(1, len(df_pdf) + 1)
            df_pdf['clean_name'] = df_pdf['Dispositivo Médico'].apply(clean_exact)

            for _, pdf_row in df_pdf.iterrows():
                p_name_raw = pdf_row['clean_name']
                productos_pedidos_set.add(p_name_raw)
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

        st.session_state['productos_pedidos_set'] = productos_pedidos_set
        st.session_state['metadata_actas_store'] = metadata_actas

        # CONSTRUCCIÓN DE LA TABLA LOTES FEFO FILTRADA
        lista_lotes_fefo_view = []
        df_lotes_filtrados = df_lotes_grouped[df_lotes_grouped['clean_name'].isin(productos_pedidos_set)]

        for idx_l, row_l in df_lotes_filtrados.iterrows():
            key_l = (row_l['clean_name'], row_l['Lote'], row_l['F. Cad'])
            inicial_val = float(row_l['Existen']) if pd.notna(row_l['Existen']) else 0.0
            restante_val = stock_disponible_dict.get(key_l, 0.0)
            
            f_cad_str = row_l['F. Cad DT'].strftime('%d/%m/%Y') if pd.notna(row_l['F. Cad DT']) else str(row_l['F. Cad'])
            
            if restante_val <= 0:
                estado_din = f"Consumido en Lote ({int(inicial_val)})"
            else:
                estado_din = "Disponible"

            lista_lotes_fefo_view.append({
                '#': len(lista_lotes_fefo_view) + 1,
                'PRODUCTO (NORMALIZADO)': str(row_l['Producto']).upper(),
                'LOTE': str(row_l['Lote']),
                'FECHA DE CADUCIDAD': f_cad_str,
                'EXISTENCIAS INICIALES': int(inicial_val),
                'STOCK RESTANTE POST-LOTE': int(restante_val),
                'ESTADO DINÁMICO': estado_din,
                'clean_name': row_l['clean_name']
            })

        df_lotes_fefo_final = pd.DataFrame(lista_lotes_fefo_view)
        st.session_state['df_lotes_fefo_view'] = df_lotes_fefo_final

        df_final = pd.DataFrame(todos_los_resultados)
        if df_final.empty:
            return None, None, None

        # MATRIZ EXCEL
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

        # PDF ACTAS COMPLETO (CONSULTANDO FIRMANTES ACTUALIZADOS)
        output_pdf_bytes = generar_pdf_consolidado_actualizado(metadata_actas, df_final)

        return output_excel.getvalue(), output_pdf_bytes, df_final
    except Exception as e:
        st.error(f"Error interno en el procesamiento: {e}")
        return None, None, None


# ==========================================
# RENDERIZADO DE CONTENIDO SEGÚN LA PESTAÑA ACTIVA
# ==========================================

# 1. INICIO
if selected_tab == "🏠 Inicio":
    st.write("")
    st.markdown("<br>", unsafe_allow_html=True)

    col_izq, col_der = st.columns([6, 4])

    with col_izq:
        st.markdown("""
        <div style="text-align: left; padding: 10px 0;">
            <h1 style="color: #1f4e78; font-weight: 700; font-size: 3rem; margin-bottom: 0;">¡Bienvenido!</h1>
            <h1 style="color: #163a5f; font-weight: 800; font-size: 3.8rem; margin-top: 10px; line-height: 1.15;">Hospital Provincial General Docente Vicente Corral Moscoso</h1>
        </div>
        """, unsafe_allow_html=True)

# 2. CRUCE DE LOTE
elif selected_tab == "📚 Cruce de Lote":
    st.markdown("""
    <div class="hospital-header">
        <h1>🏥 Sistema de Gestión FEFO & Cruce de Pedidos</h1>
        <p>Hospital Provincial General Docente Vicente Corral Moscoso | Control Automatizado de Inventarios y Dispensación</p>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("### 📂 Carga de Archivos Institucionales")
    
    col_card1, col_card2, col_card3 = st.columns(3)

    with col_card1:
        with st.container(border=True):
            st.markdown("📄 **1. Pedidos en PDF (Hasta 30+)**")
            st.markdown("<span style='font-size:0.82rem; color:#555;'>Solicitudes de despacho con servicios y firmas electrónicas.</span>", unsafe_allow_html=True)
            uploaded_pdfs = st.file_uploader(
                "Cargar PDFs de Pedidos", 
                type=["pdf"], 
                accept_multiple_files=True, 
                key="unique_pdfs_key"
            )

    with col_card2:
        with st.container(border=True):
            st.markdown("📊 **2. Excel de Lotes (FEFO)**")
            st.markdown("<span style='font-size:0.82rem; color:#555;'>Columnas: Producto, Lote, F. Cad, Existen.</span>", unsafe_allow_html=True)
            uploaded_lotes = st.file_uploader(
                "Cargar Excel de Lotes", 
                type=["xlsx", "xls"], 
                key="unique_lotes_key"
            )

    with col_card3:
        with st.container(border=True):
            st.markdown("📈 **3. Excel Unificado / Stock**")
            st.markdown("<span style='font-size:0.82rem; color:#555;'>Columnas: Producto, Código, SKU, Stock Bodega.</span>", unsafe_allow_html=True)
            uploaded_unificado = st.file_uploader(
                "Cargar Excel Unificado", 
                type=["xlsx", "xls"], 
                key="unique_unificado_key"
            )

    st.markdown("<br>", unsafe_allow_html=True)

    with st.container(border=True):
        col_chk, col_btn = st.columns([2.2, 1.8])
        with col_chk:
            st.markdown("<br>", unsafe_allow_html=True)
            chk_descuento = st.checkbox("Descuento acumulativo secuencial de lotes entre pedidos", value=True, key="chk_desc_secuencial")
        with col_btn:
            btn_procesar = st.button("🔄 Ejecutar Cruce Hospitalario en Lote ➔", key="unique_process_btn_key")

    st.markdown("<br>", unsafe_allow_html=True)

    if 'show_firmantes_config' not in st.session_state:
        st.session_state['show_firmantes_config'] = False

    with st.container(border=True):
        col_f_title, col_f_btn = st.columns([8, 2])
        
        def_elab = st.session_state.get('unique_elab_key', "RESPONSABLE DE SERVICIO")
        def_entre = st.session_state.get('unique_entre_key', "GUARDALMACÉN / DESPACHADOR")
        def_vali = st.session_state.get('unique_vali_key', "XIMENA KARINA CABRERA ZENTENO")

        with col_f_title:
            st.markdown("🖋️ **Personalizar Firmantes Globales** &nbsp;&nbsp; <span style='background-color:#e2ebf4; color:#1f4e78; padding:3px 10px; border-radius:12px; font-size:0.75rem; font-weight:600;'>Plantilla base para nuevas cargas</span>", unsafe_allow_html=True)
            if not st.session_state['show_firmantes_config']:
                st.markdown(f"<span style='color:#555; font-size:0.85rem;'>Elaboró: <b>{def_elab}</b> &nbsp;·&nbsp; Entregó: <b>{def_entre}</b> &nbsp;·&nbsp; Validó: <b>{def_vali}</b></span>", unsafe_allow_html=True)

        with col_f_btn:
            st.markdown("<div style='text-align: right;'>", unsafe_allow_html=True)
            if st.session_state['show_firmantes_config']:
                if st.button("Ocultar ajustes ▴", key="btn_toggle_firmantes"):
                    st.session_state['show_firmantes_config'] = False
                    st.rerun()
            else:
                if st.button("Configurar ▾", key="btn_toggle_firmantes"):
                    st.session_state['show_firmantes_config'] = True
                    st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)

        if st.session_state['show_firmantes_config']:
            st.markdown("<hr style='margin: 10px 0; border-color: #eaeaea;'>", unsafe_allow_html=True)
            
            col_desc, col_rest = st.columns([7, 3])
            with col_desc:
                st.markdown("<span style='color:#555; font-size:0.85rem;'>Define los firmantes predeterminados al procesar los PDFs.</span>", unsafe_allow_html=True)
            with col_rest:
                if st.button("🔄 Restaurar Predeterminados", key="btn_restaurar_firmas"):
                    st.session_state['unique_elab_key'] = "RESPONSABLE DE SERVICIO"
                    st.session_state['unique_entre_key'] = "GUARDALMACÉN / DESPACHADOR"
                    st.session_state['unique_vali_key'] = "XIMENA KARINA CABRERA ZENTENO"
                    st.rerun()

            st.markdown("<br>", unsafe_allow_html=True)

            c_f1, c_f2, c_f3 = st.columns(3)
            with c_f1:
                custom_elab = st.text_input("1. Elaboró / Recibido", def_elab, key="unique_elab_key")
            with c_f2:
                custom_entre = st.text_input("2. Entregado Bodega", def_entre, key="unique_entre_key")
            with c_f3:
                custom_vali = st.text_input("3. Validó / Autorizó", def_vali, key="unique_vali_key")

    custom_elab = st.session_state.get('unique_elab_key', def_elab)
    custom_entre = st.session_state.get('unique_entre_key', def_entre)
    custom_vali = st.session_state.get('unique_vali_key', def_vali)

    st.markdown("<br>", unsafe_allow_html=True)

    if btn_procesar:
        if uploaded_lotes and uploaded_unificado and uploaded_pdfs:
            with st.spinner("⏳ Procesando cruce de lotes, asignando algoritmos FEFO y generando actas..."):
                
                lotes_bytes = uploaded_lotes.getvalue()
                unificado_bytes = uploaded_unificado.getvalue()
                
                lista_pdfs = []
                for pdf in uploaded_pdfs:
                    pdf_bytes = pdf.getvalue() if hasattr(pdf, 'getvalue') else pdf.read()
                    lista_pdfs.append((pdf.name, pdf_bytes))

                firmantes_config = {
                    'elaborado_por': custom_elab,
                    'entregado_por': custom_entre,
                    'validado_por': custom_vali
                }

                excel_res, pdf_res, df_resultado = procesar_cruce_hospitalario_lote(
                    lista_pdfs, lotes_bytes, unificado_bytes, firmantes_custom=firmantes_config
                )

                if excel_res and pdf_res and df_resultado is not None:
                    st.session_state['excel_res'] = excel_res
                    st.session_state['pdf_res'] = pdf_res
                    st.session_state['df_resultado'] = df_resultado
                    st.success("🎉 **¡Procesamiento finalizado con éxito!** Se han asignado los lotes óptimos FEFO y generado las actas.")
                else:
                    st.error("🚨 No se pudieron generar los resultados. Asegúrese de que las estructuras internas sean correctas.")
        else:
            st.warning("⚠️ **Requisito incompleto:** Por favor, cargue los tres tipos de archivo en sus respectivas tarjetas antes de ejecutar el proceso.")

    if 'df_resultado' in st.session_state and st.session_state['df_resultado'] is not None:
        df_res = st.session_state['df_resultado']
        num_pdfs_val = df_res['PDF Origen'].nunique()
        num_servicios_val = df_res['Servicio'].nunique()
        num_items_val = len(df_res)
        
        sin_quiebres = df_res[~df_res['Lotes y Cantidades a Despachar'].str.contains('FALTAN|SIN LOTE', case=False, na=False)]
        lotes_asignados_val = len(sin_quiebres)
        porcentaje_asignados = int((lotes_asignados_val / num_items_val) * 100) if num_items_val > 0 else 0
        quiebres_val = num_items_val - lotes_asignados_val
    else:
        num_pdfs_val = len(uploaded_pdfs) if uploaded_pdfs else 0
        num_servicios_val = num_pdfs_val
        num_items_val = 0
        lotes_asignados_val = 0
        porcentaje_asignados = 0
        quiebres_val = 0

    col_kpi1, col_kpi2, col_kpi3, col_kpi4 = st.columns(4)

    with col_kpi1:
        with st.container(border=True):
            st.markdown("Pedidos en Lote &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 📄", unsafe_allow_html=True)
            st.markdown(f"<h2 style='color:#163a5f; margin: 0; padding: 5px 0;'>{num_pdfs_val}</h2>", unsafe_allow_html=True)
            st.markdown(f"<span style='color:#666; font-size:0.82rem;'>en {num_servicios_val} servicios</span>", unsafe_allow_html=True)

    with col_kpi2:
        with st.container(border=True):
            st.markdown("Líneas de Pedido &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 📦", unsafe_allow_html=True)
            st.markdown(f"<h2 style='color:#163a5f; margin: 0; padding: 5px 0;'>{num_items_val}</h2>", unsafe_allow_html=True)
            st.markdown("<span style='color:#666; font-size:0.82rem;'>ítems cruzados</span>", unsafe_allow_html=True)

    with col_kpi3:
        with st.container(border=True):
            st.markdown("Lotes 100% Asignados &nbsp;&nbsp;&nbsp; ✅", unsafe_allow_html=True)
            st.markdown(f"<h2 style='color:#008000; margin: 0; padding: 5px 0;'>{lotes_asignados_val} <span style='font-size:1rem; color:#666;'>({porcentaje_asignados}%)</span></h2>", unsafe_allow_html=True)
            st.markdown("<span style='color:transparent; font-size:0.82rem;'>.</span>", unsafe_allow_html=True)

    with col_kpi4:
        with st.container(border=True):
            st.markdown("Quiebres / Sin Stock &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ⚠️", unsafe_allow_html=True)
            st.markdown(f"<h2 style='color:#C00000; margin: 0; padding: 5px 0;'>{quiebres_val}</h2>", unsafe_allow_html=True)
            st.markdown("<span style='color:#666; font-size:0.82rem;'>requieren gestión</span>", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    col_search, col_alert_btn, col_dl_excel, col_dl_pdf = st.columns([3.2, 1.8, 2.5, 2.5])

    with col_search:
        search_query = st.text_input("Búsqueda", placeholder="🔍 Buscar por dispositivo, código, lote o PDF...", label_visibility="collapsed")

    with col_alert_btn:
        solo_faltantes = st.toggle("⚠️ Solo Faltantes / Alertas", value=False)

    with col_dl_excel:
        if 'excel_res' in st.session_state:
            st.download_button(
                label="📊 Descargar Excel (.xlsx)",
                data=st.session_state['excel_res'],
                file_name="Consolidado_Pedidos_Matriz_FEFO.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="btn_dl_excel_top",
                use_container_width=True
            )
        else:
            st.button("📊 Descargar Excel (.xlsx)", disabled=True, use_container_width=True, key="btn_dl_excel_dis")

    with col_dl_pdf:
        if 'pdf_res' in st.session_state:
            st.download_button(
                label="📄 Descargar Actas Oficiales PDF",
                data=st.session_state['pdf_res'],
                file_name="Actas_Oficiales_Consolidas.pdf",
                mime="application/pdf",
                key="btn_dl_pdf_top",
                use_container_width=True
            )
        else:
            st.button("📄 Descargar Actas Oficiales PDF", disabled=True, use_container_width=True, key="btn_dl_pdf_dis")

    st.markdown("<br>", unsafe_allow_html=True)

    if 'df_resultado' in st.session_state and st.session_state['df_resultado'] is not None:
        df_display = st.session_state['df_resultado'].copy()

        if search_query:
            q = search_query.upper()
            df_display = df_display[
                df_display['Dispositivo Médico'].str.upper().str.contains(q, na=False) |
                df_display['Código'].str.upper().str.contains(q, na=False) |
                df_display['Lotes y Cantidades a Despachar'].str.upper().str.contains(q, na=False) |
                df_display['PDF Origen'].str.upper().str.contains(q, na=False)
            ]

        if solo_faltantes:
            df_display = df_display[df_display['Lotes y Cantidades a Despachar'].str.contains('FALTAN|SIN LOTE', case=False, na=False)]

        servicios_disponibles = ['TODOS'] + sorted(df_display['Servicio'].unique().tolist())
        conteo_serv = df_display['Servicio'].value_counts()
        
        selected_serv_tab = st.radio(
            "Servicios",
            options=servicios_disponibles,
            format_func=lambda x: f"TODOS ({len(df_display)})" if x == 'TODOS' else f"{x} ({conteo_serv.get(x, 0)})",
            horizontal=True,
            label_visibility="collapsed",
            key="service_tabs_selector"
        )

        if selected_serv_tab != 'TODOS':
            df_display = df_display[df_display['Servicio'] == selected_serv_tab]

        st.markdown("<br>", unsafe_allow_html=True)

        st.markdown(f"""
        <div style="background-color: #1f4e78; color: white; padding: 10px 14px; border-top-left-radius: 8px; border-top-right-radius: 8px; font-weight: 600; font-size: 0.85rem;">
            TABLA DE CRUCE FEFO Y ASIGNACIÓN DE LOTES POR SERVICIO
        </div>
        """, unsafe_allow_html=True)

        st.dataframe(
            df_display[['Nro.', 'PDF Origen', 'Servicio', 'Código', 'SKU', 'Dispositivo Médico', 'Cant. Solicitada', 'Cant. Aprobada', 'Lotes y Cantidades a Despachar', 'Fechas de Caducidad', 'Stock Bodega']],
            use_container_width=True,
            hide_index=True
        )

# 3. LOTES & BODEGA
elif selected_tab == "📊 Lotes & Bodega":
    st.markdown("""
    <div class="hospital-header">
        <h1>📊 Control General de Lotes & Bodega</h1>
        <p>Visualización Interactiva de Lotes FEFO y Stock Maestro del Hospital</p>
    </div>
    """, unsafe_allow_html=True)

    sub_tab_lotes = st.radio(
        "Subpestañas Lotes",
        ["📦 Lotes FEFO", "📈 Stock Maestro Bodega"],
        horizontal=True,
        label_visibility="collapsed",
        key="subtab_lotes_bodega_selector"
    )

    st.markdown("<br>", unsafe_allow_html=True)

    if sub_tab_lotes == "📦 Lotes FEFO":
        col_space, col_search_lotes = st.columns([6, 4])
        with col_search_lotes:
            search_lotes_query = st.text_input("Buscar lote o producto", placeholder="🔍 Buscar lote o producto...", label_visibility="collapsed", key="search_lotes_fefo_input")

        if 'df_lotes_fefo_view' in st.session_state and st.session_state['df_lotes_fefo_view'] is not None:
            df_lotes_view = st.session_state['df_lotes_fefo_view'].copy()
            
            if search_lotes_query:
                sq = search_lotes_query.upper()
                df_lotes_view = df_lotes_view[
                    df_lotes_view['PRODUCTO (NORMALIZADO)'].str.contains(sq, na=False) |
                    df_lotes_view['LOTE'].str.contains(sq, na=False) |
                    df_lotes_view['FECHA DE CADUCIDAD'].str.contains(sq, na=False) |
                    df_lotes_view['ESTADO DINÁMICO'].str.contains(sq, na=False)
                ]

            st.dataframe(
                df_lotes_view[['#', 'PRODUCTO (NORMALIZADO)', 'LOTE', 'FECHA DE CADUCIDAD', 'EXISTENCIAS INICIALES', 'STOCK RESTANTE POST-LOTE', 'ESTADO DINÁMICO']],
                use_container_width=True,
                hide_index=True
            )
        else:
            st.info("ℹ️ Para visualizar el detalle de **Lotes FEFO** de sus pedidos, cargue los archivos y ejecute el cruce en la pestaña **'Cruce de Lote'**.")

    elif sub_tab_lotes == "📈 Stock Maestro Bodega":
        col_space, col_search_unif = st.columns([6, 4])
        with col_search_unif:
            search_unif_query = st.text_input("Buscar producto maestro", placeholder="🔍 Buscar código, SKU o producto...", label_visibility="collapsed", key="search_unif_input")

        if 'df_unif_raw' in st.session_state and st.session_state['df_unif_raw'] is not None and 'productos_pedidos_set' in st.session_state:
            df_u_view = st.session_state['df_unif_raw'].copy()
            df_u_view['clean_name'] = df_u_view['Producto'].apply(clean_exact)
            df_u_view = df_u_view[df_u_view['clean_name'].isin(st.session_state['productos_pedidos_set'])]
            df_u_view = df_u_view.drop(columns=['clean_name'])

            if search_unif_query:
                uq = search_unif_query.upper()
                cols_str = [c for c in df_u_view.columns if df_u_view[c].dtype == object]
                mask = False
                for c in cols_str:
                    mask = mask | df_u_view[c].str.upper().str.contains(uq, na=False)
                df_u_view = df_u_view[mask]

            st.dataframe(df_u_view, use_container_width=True, hide_index=True)
        else:
            st.info("ℹ️ Para visualizar el **Stock Maestro de Bodega** filtrado por sus pedidos, ejecute primero el cruce en la pestaña **'Cruce de Lote'**.")

# 4. ACTAS OFICIALES
elif selected_tab == "📄 Actas Oficiales":
    if 'vista_acta_activa' not in st.session_state:
        st.session_state['vista_acta_activa'] = None

    if st.session_state['vista_acta_activa'] is not None:
        meta_v = st.session_state['vista_acta_activa']
        pdf_key = meta_v['nombre_pdf']
        df_res_actas = st.session_state['df_resultado']
        df_acta_item = df_res_actas[df_res_actas['PDF Origen'] == pdf_key]

        edit_state_key = f"editing_firmantes_{pdf_key}"
        if edit_state_key not in st.session_state:
            st.session_state[edit_state_key] = False

        if st.button("⬅️ Volver a la lista de Actas Oficiales", key="btn_volver_lista"):
            st.session_state['vista_acta_activa'] = None
            st.rerun()

        st.markdown("<br>", unsafe_allow_html=True)

        with st.container(border=True):
            col_info_v, col_edit_v, col_dl_v, col_print_v, col_close_v = st.columns([3.5, 2.2, 2.2, 1.6, 1.5])
            
            firmantes_actuales = obtener_firmantes_acta(meta_v)
            pdf_bytes_acta = generar_pdf_individual_acta(meta_v, df_acta_item)

            with col_info_v:
                st.markdown(f"**Vista Previa:** <span style='font-size:0.82rem; font-family:monospace;'>{pdf_key}</span>", unsafe_allow_html=True)
            with col_edit_v:
                if st.button("✏️ Editar Firmantes", key=f"panel_edit_{pdf_key}", use_container_width=True):
                    st.session_state[edit_state_key] = not st.session_state[edit_state_key]
            with col_dl_v:
                st.download_button(
                    label="📥 Descargar PDF",
                    data=pdf_bytes_acta,
                    file_name=f"Acta_{pdf_key.replace('.pdf', '')}.pdf",
                    mime="application/pdf",
                    key=f"panel_dl_{pdf_key}",
                    use_container_width=True
                )
            with col_print_v:
                if st.button("🖨️ Imprimir", key=f"panel_print_{pdf_key}", use_container_width=True):
                    st.toast("Abriendo diálogo de impresión...")
            with col_close_v:
                if st.button("❌ Cerrar", key=f"panel_close_{pdf_key}", use_container_width=True):
                    st.session_state['vista_acta_activa'] = None
                    st.rerun()

            if st.session_state[edit_state_key]:
                st.markdown("<br>", unsafe_allow_html=True)
                with st.container(border=True):
                    st.markdown(f"**Editando Firmantes para este documento:** <span style='font-family:monospace; color:#1f4e78;'>{pdf_key}</span>", unsafe_allow_html=True)
                    st.markdown("<span style='font-size:0.8rem; color:#555;'>Los cambios aplicarán <b>únicamente</b> a esta acta y se reflejarán tanto aquí como en la descarga consolidada.</span>", unsafe_allow_html=True)
                    st.markdown("<br>", unsafe_allow_html=True)

                    nuevo_elab = st.text_input("1. Elaboró / Recibido", firmantes_actuales['elaborado_por'], key=f"in_elab_{pdf_key}")
                    nuevo_entre = st.text_input("2. Entregado Bodega", firmantes_actuales['entregado_por'], key=f"in_entre_{pdf_key}")
                    nuevo_vali = st.text_input("3. Validó / Autorizó", firmantes_actuales['validado_por'], key=f"in_vali_{pdf_key}")

                    st.markdown("<br>", unsafe_allow_html=True)
                    col_save_btn, col_cancel_btn = st.columns(2)
                    with col_save_btn:
                        if st.button("💾 Guardar Cambios de esta Acta", key=f"save_ind_{pdf_key}", use_container_width=True):
                            if 'firmantes_individuales_store' not in st.session_state:
                                st.session_state['firmantes_individuales_store'] = {}
                            st.session_state['firmantes_individuales_store'][pdf_key] = {
                                'elaborado_por': nuevo_elab,
                                'entregado_por': nuevo_entre,
                                'validado_por': nuevo_vali
                            }
                            st.session_state[edit_state_key] = False
                            st.success("✅ ¡Firmantes actualizados con éxito para esta acta y para el consolidado general!")
                    with col_cancel_btn:
                        if st.button("Cancelar", key=f"cancel_ind_{pdf_key}", use_container_width=True):
                            st.session_state[edit_state_key] = False

            st.markdown("<hr style='margin: 20px 0; border-color: #ddd;'>", unsafe_allow_html=True)

            st.markdown("<div style='text-align: center; font-family: Helvetica, sans-serif;'>", unsafe_allow_html=True)
            st.markdown("<span style='font-size:0.8rem; color:#555; font-weight:600;'>MINISTERIO DE SALUD PÚBLICA DEL ECUADOR</span>", unsafe_allow_html=True)
            st.markdown("<h3 style='color:#163a5f; margin: 4px 0;'>HOSPITAL PROVINCIAL GENERAL DOCENTE VICENTE CORRAL MOSCOSO</h3>", unsafe_allow_html=True)
            st.markdown("<span style='font-size:0.85rem; font-weight:bold; color:#333;'>ACTA OFICIAL DE DISPENSACIÓN Y CRUCE DE STOCK (FEFO)</span>", unsafe_allow_html=True)
            st.markdown(f"<br><span style='font-size:0.75rem; color:#777;'>Archivo de Origen: {pdf_key}</span>", unsafe_allow_html=True)
            st.markdown("</div>", unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)

            col_m_serv, col_m_emi = st.columns([6, 4])
            with col_m_serv:
                st.markdown(f"**SERVICIO:** <span style='color:#1f4e78;'>{meta_v['servicio']}</span>", unsafe_allow_html=True)
            with col_m_emi:
                st.markdown(f"**EMISIÓN:** {datetime.datetime.now().strftime('%d/%m/%Y, %H:%M:%S')}", unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)

            df_tabla_modal = df_acta_item[['Nro.', 'Código', 'Dispositivo Médico', 'Cant. Aprobada', 'Lotes y Cantidades a Despachar', 'Fechas de Caducidad', 'Stock Bodega']].copy()
            df_tabla_modal.columns = ['Nro.', 'Código', 'Dispositivo Médico', 'Cant. Ap.', 'Lotes & Cantidades a Despachar', 'Fechas Caducidad', 'Stock']

            st.dataframe(df_tabla_modal, use_container_width=True, hide_index=True)

            st.markdown("<br>", unsafe_allow_html=True)

            firmantes_actuales = obtener_firmantes_acta(meta_v)
            col_s1, col_s2, col_s3 = st.columns(3)
            with col_s1:
                st.markdown(f"<div style='font-size:0.75rem; border-top:1px solid #333; padding-top:6px; text-align:center;'><b>ELABORÓ / RECIBIDO:</b><br>{firmantes_actuales['elaborado_por']}</div>", unsafe_allow_html=True)
            with col_s2:
                st.markdown(f"<div style='font-size:0.75rem; border-top:1px solid #333; padding-top:6px; text-align:center;'><b>ENTREGADO BODEGA:</b><br>{firmantes_actuales['entregado_por']}</div>", unsafe_allow_html=True)
            with col_s3:
                st.markdown(f"<div style='font-size:0.75rem; border-top:1px solid #333; padding-top:6px; text-align:center;'><b>VALIDÓ / AUTORIZÓ:</b><br>{firmantes_actuales['validado_por']}</div>", unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown(f"<span style='font-size:0.8rem; color:#666;'>{len(df_acta_item)} ítems en este pedido</span>", unsafe_allow_html=True)

    else:
        col_t_left, col_t_right = st.columns([7, 3])
        with col_t_left:
            st.markdown("### Actas Oficiales de Dispensación y Cruce de Stock")
            st.markdown("<span style='color:#555; font-size:0.9rem;'>Cada pedido procesado cuenta con su acta institucional con firmas de responsabilidad listas para imprimir o archivar.</span>", unsafe_allow_html=True)
        with col_t_right:
            if 'pdf_res' in st.session_state:
                st.markdown("<div style='text-align: right;'>", unsafe_allow_html=True)
                
                if 'metadata_actas_store' in st.session_state and 'df_resultado' in st.session_state:
                    pdf_bytes_masivo = generar_pdf_consolidado_actualizado(
                        st.session_state['metadata_actas_store'], 
                        st.session_state['df_resultado']
                    )
                else:
                    pdf_bytes_masivo = st.session_state['pdf_res']

                st.download_button(
                    label="⬇️ Descargar Todas las Actas (.pdf)",
                    data=pdf_bytes_masivo,
                    file_name="Actas_Oficiales_Consolidas.pdf",
                    mime="application/pdf",
                    key="btn_dl_all_actas_top",
                    use_container_width=True
                )
                st.markdown("</div>", unsafe_allow_html=True)
            else:
                st.button("⬇️ Descargar Todas las Actas (.pdf)", disabled=True, use_container_width=True, key="btn_dl_all_dis")

        st.markdown("<br>", unsafe_allow_html=True)

        if 'df_resultado' in st.session_state and st.session_state['df_resultado'] is not None and 'metadata_actas_store' in st.session_state:
            df_res_actas = st.session_state['df_resultado']
            metadata_list = st.session_state['metadata_actas_store']

            for i in range(0, len(metadata_list), 3):
                cols = st.columns(3)
                for j in range(3):
                    if i + j < len(metadata_list):
                        meta = metadata_list[i + j]
                        pdf_orig = meta['nombre_pdf']
                        serv_name = meta['servicio']
                        df_acta_item = df_res_actas[df_res_actas['PDF Origen'] == pdf_orig]
                        
                        num_items_acta = len(df_acta_item)
                        has_quiebres = df_acta_item['Lotes y Cantidades a Despachar'].str.contains('FALTAN|SIN LOTE', case=False, na=False).any()
                        firmantes_card = obtener_firmantes_acta(meta)

                        with cols[j]:
                            with st.container(border=True):
                                col_badge, col_status = st.columns([7, 3])
                                with col_badge:
                                    st.markdown(f"<span style='background-color:#f1f3f5; color:#495057; padding:2px 8px; border-radius:6px; font-size:0.7rem; font-weight:600; border:1px solid #dee2e6;'>{num_items_acta} ítems</span>", unsafe_allow_html=True)
                                with col_status:
                                    if has_quiebres:
                                        st.markdown("<div style='text-align: right;'><span style='background-color:#f8d7da; color:#842029; padding:2px 8px; border-radius:6px; font-size:0.7rem; font-weight:600;'>Con Quiebres</span></div>", unsafe_allow_html=True)
                                    else:
                                        st.markdown("<div style='text-align: right;'><span style='background-color:#d1e7dd; color:#0f5132; padding:2px 8px; border-radius:6px; font-size:0.7rem; font-weight:600;'>Conforme</span></div>", unsafe_allow_html=True)

                                st.markdown(f"<h5 style='color:#163a5f; margin: 8px 0 2px 0; font-size:0.95rem;'>{serv_name}</h5>", unsafe_allow_html=True)
                                st.markdown(f"<span style='color:#6c757d; font-size:0.75rem; font-family:monospace; word-break:break-all;'>{pdf_orig}</span>", unsafe_allow_html=True)

                                st.markdown("<hr style='margin: 8px 0; border-color: #eaeaea;'>", unsafe_allow_html=True)

                                st.markdown(f"""
                                <div style='font-size:0.78rem; color:#333; line-height: 1.5;'>
                                    <b>Elaboró:</b> {firmantes_card['elaborado_por']}<br>
                                    <b>Entregó:</b> {firmantes_card['entregado_por']}<br>
                                    <b>Autorizó:</b> {firmantes_card['validado_por']}
                                </div>
                                """, unsafe_allow_html=True)

                                st.markdown("<br>", unsafe_allow_html=True)

                                col_btn_view, col_btn_dl = st.columns([7, 3])
                                with col_btn_view:
                                    if st.button("👁️ Ver Acta", key=f"btn_view_panel_{pdf_orig}_{i+j}", use_container_width=True):
                                        st.session_state['vista_acta_activa'] = meta
                                        st.session_state[f"editing_firmantes_{pdf_orig}"] = False
                                        st.rerun()
                                with col_btn_dl:
                                    pdf_bytes_individual = generar_pdf_individual_acta(meta, df_acta_item)
                                    st.download_button(
                                        label="📥",
                                        data=pdf_bytes_individual,
                                        file_name=f"Acta_{pdf_orig.replace('.pdf', '')}.pdf",
                                        mime="application/pdf",
                                        key=f"btn_dl_{pdf_orig}_{i+j}",
                                        use_container_width=True
                                    )
        else:
            st.info("📂 Cargue sus archivos y ejecute el cruce hospitalario en la pestaña **'Cruce de Lote'** para generar y visualizar las tarjetas de las Actas Oficiales.")

# 5. NORMATIVA FEFO
elif selected_tab == "🛡️ Normativa FEFO":
    st.markdown("""
    <div class="hospital-header">
        <h1>🛡️ Repositorio Normativo y Legal - Gestión FEFO & Bodega MSP</h1>
        <p>Acceso Directo a los Documentos Oficiales en PDF Publicados para Contratación, BPADT, Canjes y Red Pública RPIS</p>
    </div>
    """, unsafe_allow_html=True)

    # BLOQUE 1: CONTRATACIÓN PÚBLICA
    with st.container(border=True):
        st.markdown("### ⚖️ 1. Ley Orgánica del Sistema Nacional de Contratación Pública (LOSNCP)")
        st.markdown("""
        * **Reglamento General de la LOSNCP:** Documento oficial publicado por el SERCOP con las reglas de contratación pública, ejecución contractual y recepción de bienes.
        """, unsafe_allow_html=True)
        
        st.link_button(
            "📄 Abrir PDF Oficial: Reglamento General de la LOSNCP (SERCOP)", 
            "https://portal.compraspublicas.gob.ec/sercop/wp-content/uploads/2026/03/RGLOSNCP.pdf",
            use_container_width=True
        )

    st.markdown("<br>", unsafe_allow_html=True)

    # BLOQUE 2: ALMACENAMIENTO, DISTRIBUCIÓN Y TRANSPORTE (BPADT)
    with st.container(border=True):
        st.markdown("### 📦 2. Buenas Prácticas de Almacenamiento, Distribución y Transporte (BPADT)")
        st.markdown("""
        * **Reglamento de BPADT (Acuerdo 4872 - ARCSA):** Normativa oficial completa dictada por la ARCSA para la conservación, almacenamiento y rotación FEFO en establecimientos farmacéuticos.
        """, unsafe_allow_html=True)
        
        st.link_button(
            "📄 Abrir PDF Oficial: Reglamento de BPADT para Establecimientos Farmacéuticos (ARCSA)", 
            "https://www.controlsanitario.gob.ec/wp-content/uploads/downloads/2014/09/A-4872-Reglamento-de-BPADT-para-Establecimientos-Farmac%C3%A9uticos.pdf",
            use_container_width=True
        )

    st.markdown("<br>", unsafe_allow_html=True)

    # BLOQUE 3: REGLAMENTO DE CANJES DE MEDICAMENTOS
    with st.container(border=True):
        st.markdown("### 🔄 3. Reglamento de Canjes de Medicamentos (Acuerdo Ministerial No. 00015-2019)")
        st.markdown("""
        * **Acuerdo Ministerial No. 00015-2019:** Documento oficial emitido por el Ministerio de Salud Pública regulando los plazos y procedimientos formales para el canje de medicamentos próximos a caducar.
        """, unsafe_allow_html=True)
        
        st.link_button(
            "📄 Abrir PDF Oficial: Reglamento de Canjes (Acuerdo Ministerial 00015-2019 MSP)", 
            "https://www.salud.gob.ec/wp-content/uploads/2026/03/AC-00015-2019-AGOSTO-16_Reglamento-de-canje_compressed.pdf",
            use_container_width=True
        )

    st.markdown("<br>", unsafe_allow_html=True)

    # BLOQUE 4: DISTRIBUCIÓN Y RECEPCIÓN EN LA RED PÚBLICA (RPIS)
    with st.container(border=True):
        st.markdown("### 🏥 4. Manual de Recepción, Almacenamiento, Distribución y Transporte (Acuerdo 00050-2022)")
        st.markdown("""
        * **Manual Técnico Unificado (3er Suplemento R.O. 29):** Manual sustitutivo oficial del MSP para la recepción, almacenamiento, distribución y transporte en las entidades del Sistema Nacional de Salud.
        """, unsafe_allow_html=True)
        
        st.link_button(
            "📄 Abrir PDF Oficial: Manual de Recepción, Almacenamiento, Distribución y Transporte (MSP)", 
            "https://www.salud.gob.ec/wp-content/uploads/2022/03/Manual-recep-almacenam-distribuc-y-transp.3er.-supl.-R.O.-29-25-03-2022.-publicacion-web.pdf",
            use_container_width=True
        )
