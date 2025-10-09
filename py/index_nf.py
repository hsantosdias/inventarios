import os, re, json, hashlib, glob, cv2, pandas as pd, pytesseract
from datetime import datetime

UF_RE = r'(AC|AL|AP|AM|BA|CE|DF|ES|GO|MA|MT|MS|MG|PA|PB|PR|PE|PI|RJ|RN|RS|RO|RR|SC|SP|SE|TO)'

def sha256_file(path):
    h=hashlib.sha256()
    with open(path,'rb') as f:
        for chunk in iter(lambda: f.read(1<<20), b''): h.update(chunk)
    return h.hexdigest()

def quick_preprocess(img):
    h,w = img.shape[:2]
    target_w = 1300
    if w>target_w:
        s = target_w/w
        img = cv2.resize(img,(int(w*s), int(h*s)), interpolation=cv2.INTER_AREA)
    g = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    _, th = cv2.threshold(g, 0, 255, cv2.THRESH_BINARY+cv2.THRESH_OTSU)
    return th

def ocr_text(img):
    return pytesseract.image_to_string(img, config="--oem 1 --psm 6 -l por+eng")

def cleannum(s): return re.sub(r'\D+','', s or '')

def money(s):
    if not s: return None
    s = s.replace('R$','').replace('RS','').replace(' ','').replace('.','').replace(',', '.')
    try: return float(s)
    except: return None

def date_guess(t):
    m = re.search(r'(\d{2}[\/\-]\d{2}[\/\-]\d{2,4})', t)
    if not m: return None
    d,mn,yy = re.split(r'[\/\-]', m.group(1))
    if len(yy)==2:
        yy = int(yy) + (2000 if int(yy)<50 else 1900)
    else:
        yy = int(yy)
    return f"{yy:04d}-{int(mn):02d}-{int(d):02d}"

def extract_block(text, start_keys, end_keys):
    U=text.upper(); start=-1
    for k in start_keys:
        p=U.find(k)
        if p!=-1: start=p; break
    if start==-1: return None
    end=len(text)
    for k in end_keys:
        p=U.find(k, start+1)
        if p!=-1: end=min(end,p)
    return text[start:end].strip()

def parse_invoice(text):
    U=text.upper()
    no_space=U.replace(' ','')
    key = re.search(r'(\d{44})', no_space)
    chave = key.group(1) if key else None

    tipo = 'NFS-e' if ('NFS-E' in U or 'NOTA FISCAL DE SERVI' in U) else 'NF-e'

    cnpjs = re.findall(r'\b\d{2}\.?\d{3}\.?\d{3}\/?\d{4}\-?\d{2}\b', U)
    cnpj_emit = cleannum(cnpjs[0]) if cnpjs else None
    cnpj_dest = cleannum(cnpjs[1]) if len(cnpjs)>1 else None

    mnum = re.search(r'\bN[ºO]\s*[:\-]?\s*(\d{1,9})', U)
    numero = int(mnum.group(1)) if mnum else None
    mser = re.search(r'SERIE\s*[:\-]?\s*([A-Z0-9]{1,5})', U)
    serie = mser.group(1) if mser else None

    data_emissao = date_guess(U)

    valor_total=None
    for pat in [
        r'VALOR\s+TOTAL\s+DA\s+NOTA[:\s]*R?\$?\s*([0-9\.\,]+)',
        r'VALOR\s+TOTAL\s+DOS\s+PRODUTOS[:\s]*([0-9\.\,]+)',
        r'VALOR\s+TOTAL\s+DO\s+SERVI[CÇ]O\s*=?\s*R?\$?\s*([0-9\.\,]+)',
        r'TOTAL\s+DA\s+NOTA[:\s]*([0-9\.\,]+)',
        r'VALOR\s+TOTAL\s*[:\s]*R?\$?\s*([0-9\.\,]+)'
    ]:
        m = re.search(pat, U)
        if m:
            valor_total = money(m.group(1)); break

    muf = re.search(rf'\b{UF_RE}\b', U)
    uf = muf.group(1) if muf else None

    lines=[ln.strip() for ln in text.splitlines() if ln.strip()]
    razoes=[]
    for i,ln in enumerate(lines):
        if re.search(r'\d{2}\.?\d{3}\.?\d{3}\/?\d{4}\-?\d{2}', ln):
            if i>0: razoes.append(lines[i-1])
    razao_emit = razoes[0] if razoes else None
    razao_dest = razoes[1] if len(razoes)>1 else None

    itens_block = extract_block(text,
        ['DADOS DOS PRODUTOS/SERVI','DADOS DO PRODUTO/SERVI','DISCRIMINAÇÃO DOS SERVIÇOS','DADOS DOS PRODUTOS / SERVI'],
        ['DADOS ADICIONAIS','CALCULO DO ISSQN','CÁLCULO DO ISSQN','OUTRAS INFORMA','RESERVADO AO FISCO','INFORMAÇÕES COMPLEMENTARES'])

    return dict(
        tipo=tipo, chave_acesso=chave, numero_nf=numero, serie=serie, data_emissao=data_emissao,
        cnpj_emitente=cnpj_emit, razao_emitente_guess=razao_emit,
        cnpj_destinatario=cnpj_dest, razao_destinatario_guess=razao_dest,
        uf=uf, valor_total=valor_total, itens_raw=itens_block
    )

def parse_items_block(block, chave, arquivo):
    if not block: return []
    rows=[]
    for ln in block.splitlines():
        if not ln.strip(): continue
        u=ln.upper()
        if len(u)<15: continue
        ncm = (re.search(r'\b(\d{8})\b', u) or re.search(r'\b(\d{4}\.\d{2}\.\d{2})\b', u))
        ncm = ncm.group(1).replace('.','') if ncm else None
        cfop = (re.search(r'\b(5\d{3}|6\d{3}|1\d{3}|2\d{3})\b', u) or re.search(r'\bCFOP\s*[:\-]?\s*(\d{4})\b', u))
        cfop = cfop.group(1) if cfop else None
        # valores (melhor esforço)
        vtot=None
        m = re.search(r'(VL?\.?\s*TOTAL|V\.?\s*TOTAL)\s*[:\-]?\s*([0-9\.\,]+)', u)
        if m: vtot = money(m.group(2))
        vunit=None
        m = re.search(r'(VL?\.?\s*UNIT[AÁ]R?IO?)\s*[:\-]?\s*([0-9\.\,]+)', u)
        if m: vunit = money(m.group(2))
        qtd=None
        m = re.search(r'\bQTD[E]?\s*[:\-]?\s*([0-9]+[\,\.]?\d*)\b', u)
        if m:
            try: qtd = float(m.group(1).replace('.','').replace(',','.'))
            except: pass
        rows.append(dict(chave_acesso=chave, arquivo=arquivo, linha_ocr=ln.strip(),
                         ncm=ncm, cfop=cfop, qtd=qtd, vl_unit=vunit, vl_total=vtot))
    return rows

def process(paths):
    idx=[],[]
    index_records=[]
    item_records=[]
    for p in paths:
        img=cv2.imread(p)
        if img is None: continue
        pre=quick_preprocess(img)
        text=ocr_text(pre:=pre)  # pylint: disable=unused-variable
        parsed=parse_invoice(text)
        parsed["arquivo"]=os.path.basename(p)
        parsed["sha256"]=sha256_file(p)
        index_records.append(parsed)
        item_records.extend(parse_items_block(parsed.get("itens_raw"), parsed.get("chave_acesso"), parsed["arquivo"]))
    return index_records, item_records

if __name__=="__main__":
    import sys
    args=sys.argv[1:] or ["."]
    files=[]
    for a in args:
        if os.path.isdir(a):
            for ext in ("*.jpg","*.jpeg","*.png","*.tif","*.tiff","*.bmp","*.pdf"):
                files.extend(glob.glob(os.path.join(a,ext)))
        else:
            files.append(a)
    files=sorted(files)

    idx,itens = process(files)

    df_index = pd.DataFrame(idx, columns=[
        "arquivo","tipo","chave_acesso","numero_nf","serie","data_emissao","cnpj_emitente",
        "razao_emitente_guess","cnpj_destinatario","razao_destinatario_guess","uf","valor_total","sha256","itens_raw"
    ])
    df_items = pd.DataFrame(itens, columns=["chave_acesso","arquivo","ncm","cfop","qtd","vl_unit","vl_total","linha_ocr"])

    def flag(row):
        flags=[]
        if not row.get("chave_acesso"): flags.append("SEM_CHAVE")
        if not row.get("cnpj_emitente"): flags.append("SEM_CNPJ_EMITENTE")
        if not row.get("data_emissao"): flags.append("SEM_DATA")
        if row.get("valor_total") is None: flags.append("SEM_VALOR")
        return ";".join(flags)

    df_valid = df_index.copy()
    df_valid["flags"]=df_valid.apply(flag,axis=1)

    # Saídas
    out_dir="saida_nf"; os.makedirs(out_dir, exist_ok=True)
    df_index.to_excel(os.path.join(out_dir,"nf_index.xlsx"), index=False)
    df_items.to_excel(os.path.join(out_dir,"nf_itens.xlsx"), index=False)
    df_valid.to_excel(os.path.join(out_dir,"nf_validacao.xlsx"), index=False)
    df_index.to_csv(os.path.join(out_dir,"nf_index.csv"), index=False, encoding="utf-8-sig")
    df_items.to_csv(os.path.join(out_dir,"nf_itens.csv"), index=False, encoding="utf-8-sig")

    # JSON por NF
    jdir=os.path.join(out_dir,"json"); os.makedirs(jdir, exist_ok=True)
    for rec in idx:
        with open(os.path.join(jdir, rec["arquivo"]+".json"),"w",encoding="utf-8") as f:
            json.dump(rec, f, ensure_ascii=False, indent=2)

    print(f"[OK] Processados {len(df_index)} arquivos. Saída em: {out_dir}/")
