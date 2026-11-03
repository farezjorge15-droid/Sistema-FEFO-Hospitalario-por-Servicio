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

        # PASO 1: Pre-detectar los nombres base de servicio para cada PDF y asegurar nombres únicos limpios
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

        # Asignar sufijos numéricos (1, 2, 3...) si se repite el nombre del servicio
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
                    'Nro.': 1
