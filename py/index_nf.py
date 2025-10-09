# -*- coding: utf-8 -*-
import os, re, json, glob, cv2, pandas as pd, numpy as np, hashlib
from datetime import datetime
import pytesseract
from concurrent.futures import ProcessPoolExecutor, as_completed
from unidecode import unidecode

# ------------------------------ Constantes/Regex ------------------------------
UF_RE = r'\b(AC|AL|AP|AM|BA|CE|DF|ES|GO|MA|MT|MS|MG|PA|PB|PR|PE|PI|RJ|RN|RS|RO|RR|SC|SP|SE|TO)\b'
CNPJ_RE = r'\b\d{2}\.?\d{3}\.?\d{3}\/?\d{4}-?\d{2}\b'
ONLY_NUM = re.compile(r'\D+')

VALOR_PATS = [
    r'VALOR\s+TOTAL\s+DA\s+NOTA[:\s]*R?\$?\s*([0-9\.\,]+)',
    r'VALOR\s+TOTAL\s+DOS\s+PRODUTOS[:\s]*R?\$?\s*([0-9\.\,]+)',
    r'VALOR\s+TOTAL\s+DO\s+SERVI[CÇ]O\S*[:=\s]*R?\$?\s*([0-9\.\,]+)',
    r'VALOR\s+DA\s+NOTA[:\s]*R?\$?\s*([0-9\.\,]+)',
    r'TOTAL\s+DA\s+NOTA[:\s]*R?\$?\s*([0-9\.\,]+)',
    r'VALOR\s+TOTAL\s*[:\s]*R?\$?\s*([0-9\.\,]+)',
    r'VALOR\s+TOTAL\s+DA\s+NFS-?E[:\s]*R?\$?\s*([0-9\.\,]+)',
    r'VALOR\s+TOTAL\s+DOS\s+SERVI[CÇ]OS[:\s]*R?\$?\s*([0-9\.\,]+)',
    r'VALOR\s+TOTAL\s+DO\s+CTE[:\s]*R?\$?\s*([0-9\.\,]+)'
]

ITEMS_START = [
    'DADOS DOS PRODUTOS/SERVI', 'DADOS DO PRODUTO/SERVI',
    'DADOS DOS PRODUTOS / SERVI', 'DISCRIMINAÇÃO DOS SERVIÇOS',
    'DISCRIMINACAO DOS SERVICOS', 'ITENS DA NOTA', 'DADOS DOS PRODUTOS'
]
ITEMS_END = [
    'DADOS ADICIONAIS', 'CALCULO DO ISSQN', 'CÁLCULO DO ISSQN',
    'RESERVADO AO FISCO', 'INFORMAÇÕES COMPLEMENTARES', 'INFORMACOES COMPLEMENTARES',
    'CÁLCULO DO IMPOSTO', 'CALCULO DO IMPOSTO'
]

# ------------------------------ Utilitários ------------------------------
def sha256_file(path):
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        for chunk in iter(lambda: f.read(1<<20), b''):
            h.update(chunk)
    return h.hexdigest()

def _deskew(gray):
    # binário p/ skew
    th = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY+cv2.THRESH_OTSU)[1]
    th = 255 - th
    coords = np.column_stack(np.where(th > 0))
    if coords.size == 0: return gray
    angle = cv2.minAreaRect(coords)[-1]
    if angle < -45: angle = 90 + angle
    M = cv2.getRotationMatrix2D((gray.shape[1]//2, gray.shape[0]//2), angle, 1.0)
    return cv2.warpAffine(gray, M, (gray.shape[1], gray.shape[0]), flags=cv2.INTER_CUBIC, borderMode=cv2.BORDER_REPLICATE)

def quick_preprocess(img):
    h, w = img.shape[:2]
    if w > 1800:
        s = 1800.0 / w
        img = cv2.resize(img, (int(w*s), int(h*s)), interpolation=cv2.INTER_AREA)

    g = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    g = cv2.fastNlMeansDenoising(g, h=10)                    # remove ruído fino
    g = _deskew(g)
    g = cv2.equalizeHist(g)                                  # normaliza contraste
    # adaptive threshold ajuda com variações de fundo
    th = cv2.adaptiveThreshold(g, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                               cv2.THRESH_BINARY, 31, 11)
    return th

def _fix_ocr_numbers(s: str) -> str:
    # corrige confusões comuns quando só deveria haver dígitos
    return (s.replace('O','0').replace('o','0')
             .replace('I','1').replace('l','1')
             .replace('|','1').replace('S','5')
             .replace('B','8').replace(' ‘','')
             .replace('—','-'))

def ocr_text(img):
    # 2 passadas: geral + números
    base_cfg = "--oem 1 --psm 6 -l por+eng"
    txt = pytesseract.image_to_string(img, config=base_cfg)

    num_cfg = "--oem 1 --psm 6 -l por+eng tessedit_char_whitelist=0123456789,./:- R$"
    num = pytesseract.image_to_string(img, config=num_cfg)
    return txt if len(txt) >= len(num) else txt + "\n" + num

def normalize_text(t: str) -> str:
    t = t or ''
    # padroniza: remove acentos e normaliza espaços
    t = unidecode(t)
    t = re.sub(r'[ \t]+', ' ', t)
    return t

def cleannum(s): return ONLY_NUM.sub('', s or '')

def money(s):
    if not s: return None
    s = s.strip()
    s = s.replace('R$','').replace('RS','').replace(' ', '')
    s = _fix_ocr_numbers(s)
    s = s.replace('.', '').replace(',', '.')
    try:
        v = float(s)
        return v if v >= 0 else None
    except:
        return None

def date_guess(t):
    # aceita 01/02/2021, 01-02-21, 2021-02-01 etc.
    t = t.replace('\\n', ' ')
    pats = [
        r'(\d{2}[\/\-]\d{2}[\/\-]\d{2,4})', r'(\d{4}[\/\-]\d{2}[\/\-]\d{2})'
    ]
    for pat in pats:
        m = re.search(pat, t)
        if not m: continue
        raw = re.split(r'[\/\-]', m.group(1))
        if len(raw[0]) == 4:  # YYYY-MM-DD
            yy, mn, dd = raw
        else:
            dd, mn, yy = raw
        yy = int(yy) if len(str(yy)) == 4 else (2000 + int(yy) if int(yy) < 50 else 1900 + int(yy))
        try:
            return f"{yy:04d}-{int(mn):02d}-{int(dd):02d}"
        except:
            continue
    return None

def extract_block(text, start_keys, end_keys):
    U = text.upper()
    start = -1
    for k in start_keys:
        p = U.find(k.upper())
        if p != -1:
            start = p; break
    if start == -1: return None
    end = len(text)
    for k in end_keys:
        p = U.find(k.upper(), start+1)
        if p != -1: end = min(end, p)
    return text[start:end].strip()

# ------------------------------ Parsing ------------------------------
def find_chave_acesso(U_no_space):
    # tenta achar 44 dígitos; corrige confusões (O->0 etc.)
    cand = _fix_ocr_numbers(U_no_space)
    m = re.search(r'(\d{44})', cand)
    return m.group(1) if m else None

def guess_tipo(U):
    if 'NFS-E' in U or 'NOTA FISCAL DE SERVI' in U or 'NFS E' in U: return 'NFS-e'
    if 'CT-E' in U or 'DACTE' in U or 'CONHECIMENTO DE TRANSPORTE' in U: return 'CT-e'
    if 'CARTA DE CORRECAO' in U or 'CCE' in U or 'CC-E' in U: return 'CC-e'
    # DANFE / NF-e
    return 'NF-e' if ('DANFE' in U or 'NF-E' in U or 'NOTA FISCAL ELETRONICA' in U) else 'NF-e'

def neighbor_name(lines, idx):
    # pega até 2 linhas para cima como razão social
    for up in (1,2):
        j = idx - up
        if j >= 0:
            name = lines[j].strip()
            if 3 < len(name) < 120 and not re.search(CNPJ_RE, name):
                return name
    return None

def parse_invoice(text):
    raw = text or ''
    norm = normalize_text(raw)
    U = norm.upper()
    U_no_space = U.replace(' ', '')

    chave = find_chave_acesso(U_no_space)
    tipo = guess_tipo(U)

    cnpjs = list(re.finditer(CNPJ_RE, U))
    cnpj_emit = cleannum(cnpjs[0].group(0)) if cnpjs else None
    cnpj_dest = cleannum(cnpjs[1].group(0)) if len(cnpjs) > 1 else None

    # razão social a partir da linha acima do CNPJ
    lines = [ln.strip() for ln in norm.splitlines() if ln.strip()]
    razao_emit = razao_dest = None
    if cnpjs:
        # achar índices das linhas que contêm o cnpj
        for i, ln in enumerate(lines):
            if re.search(CNPJ_RE, ln, flags=re.I):
                if not razao_emit:
                    razao_emit = neighbor_name(lines, i)
                elif not razao_dest:
                    razao_dest = neighbor_name(lines, i)
                    break

    # número/serie
    mnum = re.search(r'\bN[ºO]?\s*(?:DA\s*NOTA)?\s*[:\-]?\s*([A-Z0-9]{1,12})\b', U)
    numero = mnum.group(1) if mnum else None
    mser = re.search(r'SERIE\s*[:\-]?\s*([A-Z0-9]{1,5})', U)
    serie = mser.group(1) if mser else None

    data_emissao = date_guess(U)

    valor_total = None
    for pat in VALOR_PATS:
        m = re.search(pat, U)
        if m:
            valor_total = money(m.group(1))
            break

    # UF (preferir em linhas com endereço)
    uf = None
    for ln in lines:
        m = re.search(UF_RE, ln.upper())
        if m: uf = m.group(1); break

    # bloco de itens
    itens_block = extract_block(norm, ITEMS_START, ITEMS_END)

    return dict(
        tipo=tipo, chave_acesso=chave, numero_nf=numero, serie=serie, data_emissao=data_emissao,
        cnpj_emitente=cnpj_emit, razao_emitente_guess=razao_emit,
        cnpj_destinatario=cnpj_dest, razao_destinatario_guess=razao_dest,
        uf=uf, valor_total=valor_total, itens_raw=itens_block
    )

def parse_items_block(block, chave, arquivo):
    if not block: return []

    rows = []
    # primeiro tenta por linhas “colunadas” (muitos espaços)
    for ln in block.splitlines():
        line = ln.strip()
        if not line: continue
        u = line.upper()

        # Split por 3+ espaços como colunas
        cols = [c for c in re.split(r'\s{3,}', line) if c.strip()]
        ncm = None; cfop = None; qtd = None; vunit = None; vtot = None

        # NCM/CFOP
        m = re.search(r'\b(\d{8})\b', u) or re.search(r'\b(\d{4}\.\d{2}\.\d{2})\b', u)
        if m: ncm = m.group(1).replace('.', '')
        m = re.search(r'\b(5\d{3}|6\d{3}|1\d{3}|2\d{3})\b', u) or re.search(r'\bCFOP\s*[:\-]?\s*(\d{4})\b', u)
        if m: cfop = m.group(1)

        # QTD / UNIT / TOTAL
        # tenta por colunas
        def _to_float(s):
            try: return float(s.replace('.', '').replace(',', '.'))
            except: return None

        for c in cols:
            cu = c.upper()
            if qtd is None and re.search(r'\bQTD[E]?\b', cu): 
                qtd = _to_float(re.sub(r'[^\d,\.]','', c))
            if vunit is None and ('UNIT' in cu or 'UNITARIO' in cu):
                vunit = money(re.sub(r'[^\d,\,\.R\$S]','', c))
            if vtot is None and ('TOTAL' in cu or 'V.TOT' in cu):
                vtot = money(re.sub(r'[^\d,\,\.R\$S]','', c))

        # fallback linha inteira
        if qtd is None:
            m = re.search(r'\bQTD[E]?\s*[:\-]?\s*([0-9]+[\,\.]?\d*)\b', u)
            if m: qtd = _to_float(m.group(1))
        if vunit is None:
            m = re.search(r'(VL?\.?\s*UNIT[AÁ]R?IO?)\s*[:\-]?\s*([0-9\.\,]+)', u)
            if m: vunit = money(m.group(2))
        if vtot is None:
            m = re.search(r'(VL?\.?\s*TOTAL|V\.?\s*TOTAL)\s*[:\-]?\s*([0-9\.\,]+)', u)
            if m: vtot = money(m.group(2))

        rows.append(dict(
            chave_acesso=chave, arquivo=arquivo, linha_ocr=line,
            ncm=ncm, cfop=cfop, qtd=qtd, vl_unit=vunit, vl_total=vtot
        ))
    return rows

# ------------------------------ Pipeline ------------------------------
def process_one(p):
    img = cv2.imread(p)
    if img is None:
        return None, []
    pre = quick_preprocess(img)
    text = ocr_text(pre)
    parsed = parse_invoice(text)
    parsed["arquivo"] = os.path.basename(p)
    parsed["sha256"] = sha256_file(p)
    items = parse_items_block(parsed.get("itens_raw"), parsed.get("chave_acesso"), parsed["arquivo"])
    return parsed, items

def process(paths, workers=0):
    index_records, item_records = [], []
    if workers and workers > 1:
        with ProcessPoolExecutor(max_workers=workers) as ex:
            futs = {ex.submit(process_one, p): p for p in paths}
            for f in as_completed(futs):
                rec, items = f.result()
                if rec: index_records.append(rec); item_records.extend(items)
    else:
        for p in paths:
            rec, items = process_one(p)
            if rec: index_records.append(rec); item_records.extend(items)
    return index_records, item_records

# ------------------------------ Execução ------------------------------
if __name__ == "__main__":
    import sys
    args = sys.argv[1:] or ["."]
    files = []
    for a in args:
        if os.path.isdir(a):
            for ext in ("*.jpg","*.jpeg","*.png","*.tif","*.tiff","*.bmp"):
                files.extend(glob.glob(os.path.join(a, ext)))
        else:
            files.append(a)
    files = sorted(files)

    idx, itens = process(files, workers=max(1, os.cpu_count()//2 - 1))

    df_index = pd.DataFrame(idx, columns=[
        "arquivo","tipo","chave_acesso","numero_nf","serie","data_emissao","cnpj_emitente",
        "razao_emitente_guess","cnpj_destinatario","razao_destinatario_guess","uf","valor_total","sha256","itens_raw"
    ])
    df_items = pd.DataFrame(itens, columns=["chave_acesso","arquivo","ncm","cfop","qtd","vl_unit","vl_total","linha_ocr"])

    # ------------------------------ Flags ------------------------------
    def flag(row):
        flags = []
        t = (row.get("tipo") or "").upper()
        U = ""
        if not row.get("chave_acesso"): flags.append("SEM_CHAVE")
        if not row.get("cnpj_emitente"): flags.append("SEM_CNPJ_EMITENTE")
        if not row.get("data_emissao"): flags.append("SEM_DATA")
        if row.get("valor_total") is None: flags.append("SEM_VALOR")
        if row.get("valor_total") == 0: flags.append("VALOR_ZERO")
        # heurísticas de tipo
        if "NFS" in t: flags.append("POSSIVEL_NFSE")
        if "CC" in t: flags.append("POSSIVEL_CCE")
        if "CT" in t or "DACTE" in t: flags.append("POSSIVEL_CTE")
        # datas anômalas
        try:
            if row.get("data_emissao"):
                d = datetime.strptime(row["data_emissao"], "%Y-%m-%d").date()
                if d > datetime.today().date(): flags.append("DATA_FUTURA")
                if d.year < 2006: flags.append("DATA_ANTIGA")
        except: pass
        return ";".join(flags)

    df_valid = df_index.copy()
    df_valid["flags"] = df_valid.apply(flag, axis=1)

    # ------------------------------ Saídas ------------------------------
    out_dir = "saida_nf"; os.makedirs(out_dir, exist_ok=True)
    df_index.to_excel(os.path.join(out_dir,"nf_index.xlsx"), index=False)
    df_items.to_excel(os.path.join(out_dir,"nf_itens.xlsx"), index=False)
    df_valid.to_excel(os.path.join(out_dir,"nf_validacao.xlsx"), index=False)
    df_index.to_csv(os.path.join(out_dir,"nf_index.csv"), index=False, encoding="utf-8-sig")
    df_items.to_csv(os.path.join(out_dir,"nf_itens.csv"), index=False, encoding="utf-8-sig")

    # JSON por NF
    jdir = os.path.join(out_dir,"json"); os.makedirs(jdir, exist_ok=True)
    for rec in idx:
        with open(os.path.join(jdir, rec["arquivo"]+".json"),"w",encoding="utf-8") as f:
            json.dump(rec, f, ensure_ascii=False, indent=2)

    # Resumo: total por emissor/mês/tipo
    try:
        df_index["emissor"] = df_index["razao_emitente_guess"].fillna("(desconhecido)")
        df_index["mes"] = pd.to_datetime(df_index["data_emissao"], errors="coerce").dt.to_period("M").astype(str)
        piv1 = df_index.groupby(["emissor","tipo"], dropna=False)["valor_total"].sum().reset_index()
        piv2 = df_index.groupby(["mes","tipo"], dropna=False)["valor_total"].sum().reset_index()
        piv1.to_csv(os.path.join(out_dir,"resumo_por_emissor_tipo.csv"), index=False, encoding="utf-8-sig")
        piv2.to_csv(os.path.join(out_dir,"resumo_por_mes_tipo.csv"), index=False, encoding="utf-8-sig")
    except Exception as e:
        pass

    print(f"[OK] Processados {len(df_index)} arquivos. Saída em: {out_dir}/")
