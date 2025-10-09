#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Sistema Avançado de Extração de Dados de Notas Fiscais
Versão Aprimorada - Combina técnicas de OCR, processamento de imagem e validação
"""

import os
import re
import json
import glob
import cv2
import pandas as pd
import numpy as np
import hashlib
import logging
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Tuple, Any
import pytesseract
from concurrent.futures import ProcessPoolExecutor, as_completed
from unidecode import unidecode
from dataclasses import dataclass
import fitz  # PyMuPDF para PDFs

# ------------------------------ Configuração ------------------------------
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('nf_processor.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# ------------------------------ Constantes/Regex ------------------------------
UF_RE = r'\b(AC|AL|AP|AM|BA|CE|DF|ES|GO|MA|MT|MS|MG|PA|PB|PR|PE|PI|RJ|RN|RS|RO|RR|SC|SP|SE|TO)\b'
CNPJ_RE = r'\b\d{2}\.?\d{3}\.?\d{3}\/?\d{4}-?\d{2}\b'
CPF_RE = r'\b\d{3}\.?\d{3}\.?\d{3}-\d{2}\b'
IE_RE = r'\b(ISENTO|\d{2,3}\.?\d{3}\.?\d{3}\.?\d{1,3})\b'

VALOR_PATS = [
    r'VALOR\s+TOTAL\s+DA\s+NOTA[:\s]*R?\$?\s*([0-9][0-9\.\,]*)',
    r'VALOR\s+TOTAL\s+DOS\s+PRODUTOS[:\s]*R?\$?\s*([0-9][0-9\.\,]*)',
    r'VALOR\s+TOTAL\s+DO\s+SERVI[CÇ]O\S*[:=\s]*R?\$?\s*([0-9][0-9\.\,]*)',
    r'VALOR\s+DA\s+NOTA[:\s]*R?\$?\s*([0-9][0-9\.\,]*)',
    r'TOTAL\s+DA\s+NOTA[:\s]*R?\$?\s*([0-9][0-9\.\,]*)',
    r'VALOR\s+TOTAL\s*[:\s]*R?\$?\s*([0-9][0-9\.\,]*)',
    r'TOTAL\s*R?\$?\s*([0-9][0-9\.\,]*)',
    r'VALOR\s+TOTAL\s+DA\s+NFS-?E[:\s]*R?\$?\s*([0-9][0-9\.\,]*)',
    r'VALOR\s+TOTAL\s+DOS\s+SERVI[CÇ]OS[:\s]*R?\$?\s*([0-9][0-9\.\,]*)',
]

ITEMS_START = [
    'DADOS DOS PRODUTOS/SERVI', 'DADOS DO PRODUTO/SERVI',
    'DADOS DOS PRODUTOS / SERVI', 'DISCRIMINAÇÃO DOS SERVIÇOS',
    'DISCRIMINACAO DOS SERVICOS', 'ITENS DA NOTA', 'DADOS DOS PRODUTOS',
    'DISCRIMINACAO DA NOTA', 'DESCRICAO DOS SERVICOS'
]

ITEMS_END = [
    'DADOS ADICIONAIS', 'CALCULO DO ISSQN', 'CÁLCULO DO ISSQN',
    'RESERVADO AO FISCO', 'INFORMAÇÕES COMPLEMENTARES', 'INFORMACOES COMPLEMENTARES',
    'CÁLCULO DO IMPOSTO', 'CALCULO DO IMPOSTO', 'TOTAL DA NOTA',
    'VALOR TOTAL', 'OBSERVACOES'
]

# ------------------------------ Classes de Dados ------------------------------
@dataclass
class ItemNota:
    """Representa um item da nota fiscal"""
    chave_acesso: str
    arquivo: str
    descricao: str
    ncm: Optional[str]
    cfop: Optional[str]
    qtd: Optional[float]
    unidade: Optional[str]
    vl_unit: Optional[float]
    vl_total: Optional[float]
    linha_ocr: str

@dataclass
class NotaFiscal:
    """Representa uma nota fiscal completa"""
    arquivo: str
    sha256: str
    tipo: str
    chave_acesso: Optional[str]
    numero_nf: Optional[str]
    serie: Optional[str]
    data_emissao: Optional[str]
    cnpj_emitente: Optional[str]
    razao_emitente: Optional[str]
    cnpj_destinatario: Optional[str]
    razao_destinatario: Optional[str]
    uf: Optional[str]
    valor_total: Optional[float]
    itens_raw: Optional[str]
    itens: List[ItemNota]
    endereco_emitente: Optional[str]
    municipio_emitente: Optional[str]
    ie_emitente: Optional[str]

# ------------------------------ Utilitários ------------------------------
def setup_tesseract():
    """Configura caminho do Tesseract se necessário"""
    try:
        pytesseract.get_tesseract_version()
    except pytesseract.TesseractNotFoundError:
        # Tenta encontrar automaticamente em paths comuns
        possible_paths = [
            r'C:\Program Files\Tesseract-OCR\tesseract.exe',
            r'C:\Users\*\AppData\Local\Programs\Tesseract-OCR\tesseract.exe',
            '/usr/bin/tesseract',
            '/usr/local/bin/tesseract'
        ]
        for path in possible_paths:
            if glob.glob(path):
                pytesseract.pytesseract.tesseract_cmd = glob.glob(path)[0]
                break

setup_tesseract()

def sha256_file(path):
    """Calcula hash SHA256 do arquivo"""
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        for chunk in iter(lambda: f.read(1<<20), b''):
            h.update(chunk)
    return h.hexdigest()

def pdf_to_images(pdf_path, dpi=200):
    """Converte PDF para lista de imagens"""
    try:
        doc = fitz.open(pdf_path)
        images = []
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            pix = page.get_pixmap(matrix=fitz.Matrix(dpi/72, dpi/72))
            img_data = pix.tobytes("png")
            nparr = np.frombuffer(img_data, np.uint8)
            img = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
            if img is not None:
                images.append(img)
        doc.close()
        return images
    except Exception as e:
        logger.error(f"Erro ao converter PDF {pdf_path}: {e}")
        return []

def advanced_preprocess(img):
    """Pré-processamento avançado da imagem"""
    # Redimensionamento inteligente
    h, w = img.shape[:2]
    if w > 2000:
        scale = 2000.0 / w
        new_w, new_h = int(w * scale), int(h * scale)
        img = cv2.resize(img, (new_w, new_h), interpolation=cv2.INTER_AREA)
    
    # Conversão para escala de cinza
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    
    # Múltiplas técnicas de melhoria
    # 1. Equalização de histograma adaptativa
    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8,8))
    gray = clahe.apply(gray)
    
    # 2. Filtro bilateral para reduzir ruído preservando bordas
    gray = cv2.bilateralFilter(gray, 9, 75, 75)
    
    # 3. Threshold adaptativo
    thresh = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, 
                                  cv2.THRESH_BINARY, 11, 2)
    
    # 4. Operações morfológicas para limpeza
    kernel = np.ones((1,1), np.uint8)
    thresh = cv2.morphologyEx(thresh, cv2.MORPH_CLOSE, kernel)
    thresh = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, kernel)
    
    return thresh

def smart_ocr(img):
    """OCR inteligente com múltiplas estratégias"""
    # Configurações para diferentes tipos de conteúdo
    configs = [
        "--oem 3 --psm 6 -l por+eng",  # Padrão
        "--oem 3 --psm 4 -l por+eng",  # Coluna única
        "--oem 3 --psm 8 -l por+eng",  # Palavra única
        "--oem 3 --psm 13 -l por+eng", # Linha bruta
    ]
    
    results = []
    for config in configs:
        try:
            text = pytesseract.image_to_string(img, config=config)
            results.append((config, text))
        except Exception as e:
            logger.warning(f"OCR com config {config} falhou: {e}")
    
    # Escolhe o resultado com mais conteúdo válido
    best_text = ""
    best_score = 0
    
    for config, text in results:
        # Pontua baseado em características de NF
        score = 0
        if re.search(r'\d{44}', text.replace(' ', '')):
            score += 100  # Chave de acesso
        if re.search(CNPJ_RE, text):
            score += 50   # CNPJ
        if re.search(r'NOTA FISCAL', text.upper()):
            score += 30   # Menção a nota fiscal
        if len(re.findall(r'\d+\.\d+\.\d+', text)) > 0:
            score += 20   # Números com formato de valor
            
        if score > best_score:
            best_score = score
            best_text = text
    
    return best_text if best_text else results[0][1] if results else ""

def normalize_text(text):
    """Normaliza texto removendo acentos e padronizando"""
    if not text:
        return ""
    text = unidecode(text)
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

def clean_number(s):
    """Remove caracteres não numéricos"""
    return re.sub(r'\D+', '', s or '')

def parse_money(value_str):
    """Converte string para valor monetário com tratamento robusto"""
    if not value_str:
        return None
    
    # Limpa e padroniza
    value_str = value_str.upper().replace('R$', '').replace('RS', '').strip()
    value_str = re.sub(r'[^\d,\.]', '', value_str)
    
    # Encontra o padrão de decimal
    if ',' in value_str and '.' in value_str:
        # Decide qual é o separador decimal baseado na posição
        if value_str.rfind(',') > value_str.rfind('.'):
            value_str = value_str.replace('.', '').replace(',', '.')
        else:
            value_str = value_str.replace(',', '')
    elif ',' in value_str:
        # Assume que vírgula é decimal se tiver 2 ou 3 dígitos após
        parts = value_str.split(',')
        if len(parts) > 1 and len(parts[-1]) in [2, 3]:
            value_str = value_str.replace('.', '').replace(',', '.')
        else:
            value_str = value_str.replace(',', '')
    
    try:
        return float(value_str)
    except (ValueError, TypeError):
        return None

def parse_date(text):
    """Extrai e valida datas com múltiplos formatos"""
    date_patterns = [
        r'(\d{2})[/\-](\d{2})[/\-](\d{2,4})',
        r'(\d{4})[/\-](\d{2})[/\-](\d{2})',
        r'(\d{2})\.(\d{2})\.(\d{4})',
    ]
    
    for pattern in date_patterns:
        matches = re.finditer(pattern, text)
        for match in matches:
            groups = match.groups()
            if len(groups[2]) == 2:  # Ano com 2 dígitos
                year = int(groups[2]) + (2000 if int(groups[2]) < 50 else 1900)
            else:
                year = int(groups[2])
            
            # Tenta diferentes ordenações (DD/MM/YYYY vs MM/DD/YYYY)
            for day, month in [(int(groups[0]), int(groups[1])), (int(groups[1]), int(groups[0]))]:
                try:
                    if 1 <= month <= 12 and 1 <= day <= 31:
                        date_obj = datetime(year, month, day)
                        # Verifica se a data é razoável (não muito no futuro/passado)
                        if datetime(2000, 1, 1) <= date_obj <= datetime.now() + timedelta(days=365):
                            return date_obj.strftime("%Y-%m-%d")
                except ValueError:
                    continue
    return None

def extract_text_block(text, start_patterns, end_patterns):
    """Extrai bloco de texto entre padrões de início e fim"""
    upper_text = text.upper()
    start_pos = -1
    
    for pattern in start_patterns:
        pos = upper_text.find(pattern)
        if pos != -1:
            start_pos = pos
            break
    
    if start_pos == -1:
        return None
    
    end_pos = len(text)
    for pattern in end_patterns:
        pos = upper_text.find(pattern, start_pos + 1)
        if pos != -1:
            end_pos = min(end_pos, pos)
    
    return text[start_pos:end_pos].strip()

# ------------------------------ Parsing Principal ------------------------------
def parse_invoice_data(text, filename):
    """Extrai dados principais da nota fiscal"""
    normalized = normalize_text(text)
    upper_text = normalized.upper()
    no_spaces = upper_text.replace(' ', '')
    
    # Chave de acesso
    chave_match = re.search(r'(\d{44})', no_spaces)
    chave_acesso = chave_match.group(1) if chave_match else None
    
    # Tipo de documento
    doc_type = "NF-e"
    if any(x in upper_text for x in ['NFS-E', 'NOTA FISCAL DE SERVI']):
        doc_type = "NFS-e"
    elif any(x in upper_text for x in ['CT-E', 'CONHECIMENTO']):
        doc_type = "CT-e"
    elif any(x in upper_text for x in ['CCE', 'CARTA DE CORRECAO']):
        doc_type = "CC-e"
    
    # CNPJs
    cnpjs = re.findall(CNPJ_RE, upper_text)
    cnpj_emit = clean_number(cnpjs[0]) if cnpjs else None
    cnpj_dest = clean_number(cnpjs[1]) if len(cnpjs) > 1 else None
    
    # Número e série
    num_match = re.search(r'N[ºO]?\s*[:\-]?\s*(\d{1,12})', upper_text)
    numero = num_match.group(1) if num_match else None
    
    serie_match = re.search(r'S[ÉE]RIE?\s*[:\-]?\s*(\d{1,5}|\w{1,3})', upper_text)
    serie = serie_match.group(1) if serie_match else None
    
    # Data de emissão
    data_emissao = parse_date(text)
    
    # Valor total
    valor_total = None
    for pattern in VALOR_PATS:
        match = re.search(pattern, upper_text)
        if match:
            valor_total = parse_money(match.group(1))
            if valor_total:
                break
    
    # Razões sociais (tentativa inteligente)
    lines = [line.strip() for line in normalized.split('\n') if line.strip()]
    razao_emit, razao_dest = find_company_names(lines, cnpjs)
    
    # UF
    uf_match = re.search(UF_RE, upper_text)
    uf = uf_match.group(1) if uf_match else None
    
    # Bloco de itens
    itens_block = extract_text_block(normalized, ITEMS_START, ITEMS_END)
    
    # Endereço e município
    endereco, municipio = extract_address(lines)
    
    # IE
    ie_emit = extract_ie(lines)
    
    return {
        'arquivo': filename,
        'tipo': doc_type,
        'chave_acesso': chave_acesso,
        'numero_nf': numero,
        'serie': serie,
        'data_emissao': data_emissao,
        'cnpj_emitente': cnpj_emit,
        'razao_emitente': razao_emit,
        'cnpj_destinatario': cnpj_dest,
        'razao_destinatario': razao_dest,
        'uf': uf,
        'valor_total': valor_total,
        'itens_raw': itens_block,
        'endereco_emitente': endereco,
        'municipio_emitente': municipio,
        'ie_emitente': ie_emit
    }

def find_company_names(lines, cnpjs_found):
    """Encontra razões sociais próximas aos CNPJs"""
    razao_emit, razao_dest = None, None
    cnpj_lines = []
    
    # Encontra linhas com CNPJ
    for i, line in enumerate(lines):
        if re.search(CNPJ_RE, line):
            cnpj_lines.append(i)
    
    for i, cnpj_line_idx in enumerate(cnpj_lines):
        # Procura até 3 linhas acima do CNPJ
        for j in range(1, 4):
            candidate_line = cnpj_line_idx - j
            if candidate_line >= 0:
                candidate = lines[candidate_line]
                # Verifica se parece ser um nome de empresa
                if (len(candidate) > 5 and len(candidate) < 100 and
                    not re.search(CNPJ_RE, candidate) and
                    not re.search(CPF_RE, candidate) and
                    not re.search(r'\d{4,}', candidate)):
                    if i == 0 and not razao_emit:
                        razao_emit = candidate
                    elif i == 1 and not razao_dest:
                        razao_dest = candidate
                    break
    
    return razao_emit, razao_dest

def extract_address(lines):
    """Extrai endereço e município"""
    endereco, municipio = None, None
    
    for i, line in enumerate(lines):
        upper_line = line.upper()
        # Procura por padrões de endereço
        if any(x in upper_line for x in ['RUA ', 'AV ', 'AVENIDA ', 'RODOVIA ', 'TRAVESSA ']):
            # Combina com linha seguinte se necessário
            address_parts = [line]
            if i + 1 < len(lines) and not re.search(CNPJ_RE, lines[i + 1]):
                address_parts.append(lines[i + 1])
            endereco = ' '.join(address_parts)
        
        # Procura município
        if not municipio and re.search(r',\s*[A-Z][A-Z]\s*$', line):
            municipio = re.sub(r',\s*[A-Z][A-Z]\s*$', '', line).strip()
    
    return endereco, municipio

def extract_ie(lines):
    """Extrai Inscrição Estadual"""
    for line in lines:
        ie_match = re.search(IE_RE, line, re.IGNORECASE)
        if ie_match and ie_match.group(1).upper() != 'ISENTO':
            return ie_match.group(1)
    return None

def parse_items_detailed(block_text, chave_acesso, filename):
    """Analisa detalhadamente os itens da nota"""
    if not block_text:
        return []
    
    items = []
    lines = [line.strip() for line in block_text.split('\n') if line.strip()]
    
    current_item = {}
    item_lines = []
    
    for line in lines:
        upper_line = line.upper()
        
        # Ignora cabeçalhos e totais
        if any(x in upper_line for x in ['COD', 'NCM', 'CFOP', 'QTD', 'UNIT', 'TOTAL', 'UNID']):
            continue
        
        # Tenta identificar fim do item atual
        if current_item and (re.search(r'\d+[,.]\d{2}$', line) or len(item_lines) >= 3):
            # Processa item completo
            item_data = parse_single_item(item_lines, chave_acesso, filename)
            if item_data:
                items.append(item_data)
            current_item = {}
            item_lines = []
        
        item_lines.append(line)
    
    # Processa último item
    if item_lines:
        item_data = parse_single_item(item_lines, chave_acesso, filename)
        if item_data:
            items.append(item_data)
    
    return items

def parse_single_item(lines, chave_acesso, filename):
    """Analisa um único item da nota"""
    full_text = ' '.join(lines)
    
    # NCM
    ncm_match = re.search(r'\b(\d{8})\b', full_text)
    ncm = ncm_match.group(1) if ncm_match else None
    
    # CFOP
    cfop_match = re.search(r'\b([123456]\d{3})\b', full_text)
    cfop = cfop_match.group(1) if cfop_match else None
    
    # Quantidade
    qtd = None
    qtd_match = re.search(r'(\d+[,.]?\d*)\s*(KG|UN|M|M2|M3|LT|CX|PC)', full_text.upper())
    if qtd_match:
        qtd = parse_money(qtd_match.group(1))
    
    # Valores
    vl_unit, vl_total = extract_values_from_text(full_text)
    
    return ItemNota(
        chave_acesso=chave_acesso,
        arquivo=filename,
        descricao=lines[0] if lines else "",
        ncm=ncm,
        cfop=cfop,
        qtd=qtd,
        unidade=None,  # Pode ser extraído com regex mais específico
        vl_unit=vl_unit,
        vl_total=vl_total,
        linha_ocr=full_text
    )

def extract_values_from_text(text):
    """Extrai valores unitário e total do texto"""
    vl_unit = vl_total = None
    
    # Procura padrões de valores
    value_patterns = [
        r'(\d+[.,]\d{2})\s*$',  # Valor no final da linha
        r'R\$\s*(\d+[.,]\d+)',  # Formato R$
        r'(\d{1,3}(?:[.,]\d{3})*[.,]\d{2})',  # Formato com milhares
    ]
    
    values = []
    for pattern in value_patterns:
        matches = re.findall(pattern, text)
        for match in matches:
            value = parse_money(match)
            if value and value > 0:
                values.append(value)
    
    # Ordena e tenta identificar unitário e total
    if len(values) >= 2:
        values.sort()
        vl_unit = values[0]
        vl_total = values[-1]
    elif len(values) == 1:
        vl_total = values[0]
    
    return vl_unit, vl_total

# ------------------------------ Processamento Principal ------------------------------
def process_single_file(filepath):
    """Processa um único arquivo"""
    try:
        logger.info(f"Processando: {filepath}")
        
        # Verifica tipo de arquivo
        if filepath.lower().endswith('.pdf'):
            images = pdf_to_images(filepath)
            if not images:
                return None, []
            # Usa primeira página do PDF para análise principal
            img = images[0]
        else:
            img = cv2.imread(filepath)
            if img is None:
                logger.warning(f"Não foi possível ler imagem: {filepath}")
                return None, []
        
        # Pré-processamento e OCR
        processed_img = advanced_preprocess(img)
        text = smart_ocr(processed_img)
        
        if not text.strip():
            logger.warning(f"OCR não retornou texto para: {filepath}")
            return None, []
        
        # Extração de dados
        invoice_data = parse_invoice_data(text, os.path.basename(filepath))
        invoice_data['sha256'] = sha256_file(filepath)
        
        # Extração de itens
        items = parse_items_detailed(
            invoice_data.get('itens_raw'), 
            invoice_data.get('chave_acesso'),
            invoice_data['arquivo']
        )
        
        logger.info(f"Sucesso: {filepath} - {len(items)} itens encontrados")
        return invoice_data, items
        
    except Exception as e:
        logger.error(f"Erro processando {filepath}: {str(e)}")
        return None, []

def process_files(file_paths, max_workers=None):
    """Processa múltiplos arquivos em paralelo"""
    index_data = []
    items_data = []
    
    if max_workers is None:
        max_workers = max(1, os.cpu_count() // 2)
    
    successful = 0
    total = len(file_paths)
    
    with ProcessPoolExecutor(max_workers=max_workers) as executor:
        future_to_file = {executor.submit(process_single_file, fp): fp for fp in file_paths}
        
        for future in as_completed(future_to_file):
            filepath = future_to_file[future]
            try:
                invoice_data, items = future.result()
                if invoice_data:
                    index_data.append(invoice_data)
                    items_data.extend(items)
                    successful += 1
            except Exception as e:
                logger.error(f"Erro no processamento de {filepath}: {e}")
    
    logger.info(f"Processamento concluído: {successful}/{total} arquivos processados com sucesso")
    return index_data, items_data

# ------------------------------ Validação e Exportação ------------------------------
def validate_invoice_data(invoice_data):
    """Valida dados da nota fiscal e retorna flags"""
    flags = []
    
    if not invoice_data.get('chave_acesso'):
        flags.append('SEM_CHAVE_ACESSO')
    
    if not invoice_data.get('cnpj_emitente'):
        flags.append('SEM_CNPJ_EMITENTE')
    
    if not invoice_data.get('data_emissao'):
        flags.append('SEM_DATA_EMISSAO')
    
    if invoice_data.get('valor_total') is None:
        flags.append('SEM_VALOR_TOTAL')
    elif invoice_data.get('valor_total') == 0:
        flags.append('VALOR_ZERO')
    
    if not invoice_data.get('itens_raw'):
        flags.append('SEM_ITENS')
    
    # Validação de data
    try:
        if invoice_data.get('data_emissao'):
            emissao = datetime.strptime(invoice_data['data_emissao'], '%Y-%m-%d')
            hoje = datetime.now()
            if emissao > hoje:
                flags.append('DATA_FUTURA')
            elif emissao < datetime(2000, 1, 1):
                flags.append('DATA_MUITO_ANTIGA')
    except:
        flags.append('DATA_INVALIDA')
    
    return ';'.join(flags) if flags else 'OK'

def export_results(index_data, items_data, output_dir="saida_nf_avancada"):
    """Exporta resultados em múltiplos formatos"""
    os.makedirs(output_dir, exist_ok=True)
    
    # DataFrame principal
    df_index = pd.DataFrame(index_data)
    
    # Adiciona flags de validação
    df_index['flags_validacao'] = df_index.apply(validate_invoice_data, axis=1)
    
    # DataFrame de itens
    df_items = pd.DataFrame([item.__dict__ for item in items_data])
    
    # Exporta Excel
    with pd.ExcelWriter(os.path.join(output_dir, 'notas_fiscais.xlsx')) as writer:
        df_index.to_excel(writer, sheet_name='Notas', index=False)
        df_items.to_excel(writer, sheet_name='Itens', index=False)
        
        # Sheet de validação
        df_validation = df_index[['arquivo', 'chave_acesso', 'flags_validacao']].copy()
        df_validation.to_excel(writer, sheet_name='Validacao', index=False)
    
    # Exporta CSV
    df_index.to_csv(os.path.join(output_dir, 'notas_fiscais.csv'), index=False, encoding='utf-8-sig')
    df_items.to_csv(os.path.join(output_dir, 'itens.csv'), index=False, encoding='utf-8-sig')
    
    # JSON individual
    json_dir = os.path.join(output_dir, 'json')
    os.makedirs(json_dir, exist_ok=True)
    
    for invoice in index_data:
        filename = f"{invoice['arquivo']}_{invoice.get('chave_acesso', 'sem_chave')}.json"
        with open(os.path.join(json_dir, filename), 'w', encoding='utf-8') as f:
            json.dump(invoice, f, ensure_ascii=False, indent=2)
    
    # Estatísticas
    stats = {
        'total_notas': len(df_index),
        'notas_validas': len(df_index[df_index['flags_validacao'] == 'OK']),
        'total_itens': len(df_items),
        'tipos_documento': df_index['tipo'].value_counts().to_dict(),
        'ufs': df_index['uf'].value_counts().to_dict()
    }
    
    with open(os.path.join(output_dir, 'estatisticas.json'), 'w', encoding='utf-8') as f:
        json.dump(stats, f, ensure_ascii=False, indent=2)
    
    return df_index, df_items, stats

# ------------------------------ Interface Principal ------------------------------
def main():
    """Função principal"""
    import sys
    import argparse
    
    parser = argparse.ArgumentParser(description='Processador Avançado de Notas Fiscais')
    parser.add_argument('input', nargs='+', help='Arquivos ou diretórios para processar')
    parser.add_argument('-o', '--output', default='saida_nf_avancada', help='Diretório de saída')
    parser.add_argument('-w', '--workers', type=int, help='Número de workers paralelos')
    parser.add_argument('-v', '--verbose', action='store_true', help='Log verboso')
    
    args = parser.parse_args()
    
    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)
    
    # Coleta arquivos
    files = []
    for input_path in args.input:
        if os.path.isdir(input_path):
            for ext in ('*.jpg', '*.jpeg', '*.png', '*.tif', '*.tiff', '*.bmp', '*.pdf'):
                files.extend(glob.glob(os.path.join(input_path, ext)))
        else:
            files.append(input_path)
    
    files = sorted(set(files))
    
    if not files:
        logger.error("Nenhum arquivo encontrado para processar")
        return
    
    logger.info(f"Encontrados {len(files)} arquivos para processar")
    
    # Processamento
    index_data, items_data = process_files(files, max_workers=args.workers)
    
    if not index_data:
        logger.error("Nenhum arquivo foi processado com sucesso")
        return
    
    # Exportação
    df_index, df_items, stats = export_results(index_data, items_data, args.output)
    
    # Relatório final
    logger.info("\n" + "="*50)
    logger.info("RELATÓRIO FINAL")
    logger.info("="*50)
    logger.info(f"Notas processadas: {stats['total_notas']}")
    logger.info(f"Notas válidas: {stats['notas_validas']}")
    logger.info(f"Total de itens: {stats['total_itens']}")
    logger.info(f"Tipos de documento: {stats['tipos_documento']}")
    logger.info(f"Arquivos de saída salvos em: {args.output}")

if __name__ == "__main__":
    main()