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
    """
    Toma cadenas de lotes (ej: '20250510 (1 unids) / 20251220 (4 unids) / 20251220 (5 unids)')
    y suma las cantidades de los lotes repetidos en una sola cantidad consolidada.
    """
    if not isinstance(texto_lotes_crudo, str) or not texto_lotes_crudo or 'SIN LOTE' in texto_lotes_crudo or 'FALTAN' in texto_lotes_crudo:
        return str(texto_lotes_crudo)
    
    # Separar por '/' o '|'
    segmentos = re.split(r'[/|]', texto_lotes_crudo)
    acumulador_lotes = {}
    
    for seg in segmentos:
        seg = seg.strip()
        if not seg:
            continue
        # Extraer el nombre del lote y la cantidad entre paréntesis
        match = re.search(r'^(.*?)\s*\((\d+)\s*unids?\)', seg, re.IGNORECASE)
        if match:
            lote_nombre = match.group(1).strip()
            cantidad = int(match.group(2))
            acumulador_lotes[lote_nombre] = acumulador_lotes.get(lote_nombre, 0) + cantidad
        else:
            # Si no tiene formato de cantidad, se deja tal cual
            acumulador_lotes[seg] = acumulador_lotes.get(seg, 0)
            
    # Reconstruir la cadena unificada
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

        todos_los_resultados = []
        metadata_actas = [] 

        for idx_pdf, (nombre_pdf, pdf_bytes) in enumerate(lista_pdfs):
            try:
                reader = pypdf.PdfReader(io.BytesIO(pdf_bytes))
                pdf_text = ""
                for page in reader.pages:
                    text_page = page.extract_text()
                    if text_page:
                        pdf_text += text_page + "\n"
            except Exception:
                continue

            servicio_nombre = "GENERAL"
            match_serv = re.search(r'Servicio\s*[:]\s*([^\r\n]+)', pdf_text, re.IGNORECASE)
            if match_serv:
                raw_serv = match_serv.group(1).strip()
                for stop_word in ['FECHA', 'NRO', 'VERSIÓN', 'VERSION', 'CÓDIGO', 'CODIGO']:
                    if stop_word in raw_serv.upper():
                        raw_serv = raw_serv.upper().split(stop_word)[0].strip()
                if len(raw_serv) > 1:
                    servicio_nombre = raw_serv

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

            df_pdf = pd.DataFrame(items_extraidos)
            if not df_pdf.empty:
                df_pdf['Nro.'] = range(1, len(df_pdf) + 1)
                df_pdf['clean_name'] = df_pdf['Dispositivo Médico'].apply(clean_exact)

                for _, pdf_row in df_pdf.iterrows():
                    p_name_raw = pdf_row['clean_name']
                    cant_pendiente = pdf_row['Cantidades Aprobadas']
                    
                    match_unif, _ = buscar_producto_inteligente(p_name_raw, df_unificado, df_lotes_grouped)

                    if not match_unif.empty:
                        codigo = match_unif.iloc[0].get('Código', match_unif.iloc[0].get('Codigo', 'SIN-CODIGO'))
                        sku = match_unif.iloc[0].get('SKU', 'SIN-SKU')
                        cols_posibles = [c for c in match_unif.columns if any(w in str(c).upper() for w in ['BOD', 'FARM', 'STOCK', 'EXIST', 'TOTAL'])]
                        col_bodega = cols_posibles[0] if cols_posibles else match_unif.columns[-1]
                        stock_bodega = match_unif.iloc[0][col_bodega]
                        if pd.isna(stock_bodega):
                            stock_bodega = 0.0
                    else:
                        codigo = 'SIN-CODIGO'
                        sku = 'SIN-SKU'
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
                            
                        # Consolidar lotes repetidos de forma limpia
                        lotes_crudos_str = " / ".join(lotes_info)
                        lotes_consolidados_str = consolidar_lotes_texto(lotes_crudos_str)
                            
                        item_data = {
                            'PDF Origen': nombre_pdf,
                            'Servicio': servicio_nombre,
                            'Nro.': pdf_row['Nro.'],
                            'Código': str(codigo),
                            'SKU': str(sku),
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

        # --- CONSOLIDADO EN FORMATO MATRIZ ---
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

        headers_matriz = ['DISPOSITIVO / INSUMO', 'CÓDIGO', 'STOCK BODEGA', 'LOTE Y CADUCIDAD ASIGNADA (FEFO)'] + servicios_unicos + ['TOTAL APROBADO']
        ws.append(headers_matriz)

        for col_num, _ in enumerate(headers_matriz, 1):
            cell = ws.cell(row=3, column=col_num)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

        productos_unicos = df_final[['Código', 'Dispositivo Médico', 'Stock Bodega']].drop_duplicates(subset=['Dispositivo Médico'])

        for idx, p_row in productos_unicos.iterrows():
            prod_nombre = p_row['Dispositivo Médico']
            prod_codigo = p_row['Código']
            prod_stock = p_row['Stock Bodega']
            
            df_prod_all = df_final[df_final['Dispositivo Médico'] == prod_nombre]
            
            # Unir y consolidar todos los lotes de todas las solicitudes de este producto
            lotes_brutos_totales = " / ".join(df_prod_all['Lotes y Cantidades a Despachar'].dropna().astype(str))
            lotes_str = consolidar_lotes_texto(lotes_brutos_totales)

            fila_datos = [prod_nombre, prod_codigo, prod_stock, lotes_str]
            
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
            is_alert = (prod_stock <= 0) or ('FALTAN' in lotes_str) or ('SIN LOTE' in lotes_str)

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

        # --- PDF DE ACTAS OFICIALES ---
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

st.title("Sistema FEFO - Hospital Vicente Corral Moscoso")

st.sidebar.header("📂 Carga de Archivos")
uploaded_lotes = st.sidebar.file_uploader("Excel de Lotes", type=["xlsx", "xls"], key="unique_lotes_key")
uploaded_unificado = st.sidebar.file_uploader("Excel de Stock / Unificado", type=["xlsx", "xls"], key="unique_unificado_key")
uploaded_pdfs = st.sidebar.file_uploader("Pedidos en PDF (Múltiples)", type=["pdf"], accept_multiple_files=True, key="unique_pdfs_key")

st.sidebar.header("✍️ Personalización de Firmantes")
with st.sidebar.expander("Modificar Firmantes de Actas"):
    custom_elab = st.text_input("Elaboró / Recibido", "RESPONSABLE DE SERVICIO", key="unique_elab_key")
    custom_entre = st.text_input("Entregado Bodega", "GUARDALMACÉN / DESPACHADOR", key="unique_entre_key")
    custom_vali = st.text_input("Validó / Autorizó", "XIMENA KARINA CABRERA ZENTENO", key="unique_vali_key")

if st.sidebar.button("🚀 Procesar Lote FEFO", key="unique_process_btn_key"):
    if uploaded_lotes and uploaded_unificado and uploaded_pdfs:
        with st.spinner("Procesando lotes y aplicando FEFO acumulativo..."):
            
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

            excel_res, pdf_res = procesar_cruce_hospitalario_lote(
                lista_pdfs, lotes_bytes, unificado_bytes, firmantes_custom=firmantes_config
            )

            if excel_res and pdf_res:
                st.success("¡Procesamiento completado con éxito!")
                st.download_button("📥 Descargar Excel Consolidado", excel_res, "Consolidado_Pedidos_Matriz_FEFO.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="unique_dl_excel_key")
                st.download_button("📥 Descargar Actas en PDF", pdf_res, "Actas_Oficiales_Consolidas.pdf", "application/pdf", key="unique_dl_pdf_key")
            else:
                st.error("No se pudieron generar resultados. Verifique que los archivos sean correctos.")
    else:
        st.warning("⚠️ Por favor, cargue todos los archivos necesarios (Lotes, Stock y al menos un PDF).")
