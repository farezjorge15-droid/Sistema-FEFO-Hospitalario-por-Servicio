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
                            
                        lotes_cr
