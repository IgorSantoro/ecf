"""
Cruzamento DARF × DCTFWeb — Streamlit (interface moderna)
Execute:  streamlit run cruzamento_streamlit.py
"""

import re
import io
import calendar
from datetime import datetime

import streamlit as st
import pandas as pd
import pdfplumber
from openpyxl.styles import PatternFill, Font, Alignment

# ─────────────────────────────────────────────────────────────────────────────
# Página
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="DARF × DCTFWeb",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
.block-container { padding: 2rem 2.5rem 3rem; max-width: 1400px; }

/* Header */
.app-header {
    background: linear-gradient(135deg, #0f172a 0%, #1e3a5f 60%, #1d4ed8 100%);
    border-radius: 16px;
    padding: 28px 36px;
    margin-bottom: 28px;
    display: flex;
    align-items: center;
    gap: 20px;
}
.app-header h1 { color: white; font-size: 1.6rem; font-weight: 800; margin: 0; }
.app-header p  { color: rgba(255,255,255,.65); font-size: .88rem; margin: 4px 0 0; }

/* Upload boxes */
.upload-box {
    background: white;
    border: 2px dashed #cbd5e1;
    border-radius: 14px;
    padding: 24px;
    transition: border-color .2s;
}

/* Metric cards */
.metric-card {
    background: white;
    border-radius: 14px;
    padding: 20px 22px;
    border-left: 5px solid #e2e8f0;
    box-shadow: 0 1px 6px rgba(0,0,0,.06);
    height: 100%;
}
.metric-card .val  { font-size: 1.6rem; font-weight: 800; line-height: 1.1; margin-bottom: 4px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.metric-card .lbl  { font-size: .78rem; font-weight: 600; text-transform: uppercase; letter-spacing: .05em; color: #64748b; }
.metric-card .sub  { font-size: .8rem; color: #94a3b8; margin-top: 2px; }
.card-ok    { border-color: #10b981; } .card-ok .val    { color: #059669; }
.card-erro  { border-color: #ef4444; } .card-erro .val  { color: #dc2626; }
.card-alert { border-color: #f59e0b; } .card-alert .val { color: #d97706; }
.card-info  { border-color: #3b82f6; } .card-info .val  { color: #2563eb; }
.card-warn  { border-color: #f97316; } .card-warn .val  { color: #ea580c; }

/* Section title */
.section-title {
    font-size: 1rem; font-weight: 700; color: #1e293b;
    margin: 0 0 14px; display: flex; align-items: center; gap: 8px;
}

/* Tabela HTML */
.tabela-wrap {
    background: white; border-radius: 14px; overflow: hidden;
    box-shadow: 0 1px 8px rgba(0,0,0,.07); margin-top: 4px;
}
.tabela-wrap table { width: 100%; border-collapse: collapse; font-size: .8rem; }
.tabela-wrap th {
    background: #f8fafc; padding: 11px 13px; text-align: left;
    font-weight: 700; color: #475569; border-bottom: 2px solid #e2e8f0;
    white-space: nowrap; font-size: .75rem; text-transform: uppercase; letter-spacing: .04em;
}
.tabela-wrap td { padding: 10px 13px; border-bottom: 1px solid #f1f5f9; vertical-align: middle; }
.tabela-wrap tr:last-child td { border-bottom: none; }

/* Badges */
.badge {
    display: inline-block; padding: 3px 10px; border-radius: 20px;
    font-size: .73rem; font-weight: 700; white-space: nowrap;
}
.badge-ok    { background: #d1fae5; color: #065f46; }
.badge-erro  { background: #fee2e2; color: #991b1b; }
.badge-alert { background: #fef3c7; color: #92400e; }
.badge-info  { background: #dbeafe; color: #1e40af; }
.badge-warn  { background: #ffedd5; color: #9a3412; }

/* Destaque dos valores sem DCTFWeb */
.sem-dctf-card {
    background: linear-gradient(135deg, #fff7ed, #fff);
    border: 1.5px solid #fed7aa;
    border-radius: 12px;
    padding: 16px 20px;
    margin-bottom: 10px;
}
.sem-dctf-card .comp { font-weight: 700; color: #1e293b; font-size: .95rem; }
.sem-dctf-card .val  { font-size: 1.4rem; font-weight: 800; color: #ea580c; }
.sem-dctf-card .desc { font-size: .78rem; color: #78716c; margin-top: 2px; }

/* Tag multa */
.multa-tag {
    display: inline-block; background: #fee2e2; color: #b91c1c;
    border-radius: 4px; padding: 1px 7px; font-size: .7rem; font-weight: 700;
}

/* Pill info */
.pill {
    display: inline-block; background: #f1f5f9; color: #475569;
    border-radius: 6px; padding: 2px 8px; font-size: .75rem; font-weight: 600;
}

/* Botão primário */
div.stButton > button[kind="primary"] {
    background: linear-gradient(135deg, #2563eb, #1d4ed8) !important;
    border: none !important;
    border-radius: 10px !important;
    font-weight: 700 !important;
    font-size: 1rem !important;
    padding: 14px !important;
    color: white !important;
}

/* Download button */
div.stDownloadButton > button {
    background: linear-gradient(135deg, #059669, #047857) !important;
    border: none !important;
    border-radius: 10px !important;
    font-weight: 700 !important;
    color: white !important;
}

/* Info box */
.info-box {
    background: #eff6ff; border: 1px solid #bfdbfe;
    border-radius: 10px; padding: 14px 18px;
    color: #1e40af; font-size: .85rem;
}
.warn-box {
    background: #fff7ed; border: 1px solid #fed7aa;
    border-radius: 10px; padding: 14px 18px;
    color: #92400e; font-size: .85rem;
}

/* Oculta lista nativa de arquivos do file_uploader */
[data-testid="stFileUploaderFileData"] {
    display: none !important;
}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# Utilitários
# ─────────────────────────────────────────────────────────────────────────────

def parse_data_br(texto):
    try: return datetime.strptime(texto.strip(), "%d/%m/%Y")
    except: return None

def parse_valor(texto) -> float:
    texto = str(texto).strip()
    if texto in ("-", "", "–", "Sem Movimento"): return 0.0
    try: return float(texto.replace(".", "").replace(",", "."))
    except: return 0.0

def competencia_para_dt(comp_str):
    try:
        mm, aaaa = comp_str.split("/")
        ultimo = calendar.monthrange(int(aaaa), int(mm))[1]
        return datetime(int(aaaa), int(mm), ultimo)
    except: return None

def chave(dt) -> str:
    return f"{dt.month:02d}/{dt.year}" if dt else ""

def fmt_brl(v) -> str:
    if v is None: return "—"
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

# ─────────────────────────────────────────────────────────────────────────────
# Extração DARFs
# ─────────────────────────────────────────────────────────────────────────────

CODIGO_TRIBUTO = {
    "2484": "CSLL", "6912": "CSLL",
    "2362": "IRPJ", "2390": "IRPJ",
    "1708": "COFINS", "0086": "PIS",
    "5993": "IRRF", "0422": "IRRF",
}

def extrair_darfs(pdf_bytes: bytes) -> list[dict]:
    darfs = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for num, pagina in enumerate(pdf.pages, 1):
            texto = pagina.extract_text() or ""
            d = dict(pagina=num, cnpj=None, razao_social=None,
                     periodo_apuracao=None, data_vencimento=None,
                     numero_documento=None, codigo=None, tributo=None,
                     descricao=None, principal=0.0, multa=0.0,
                     juros=0.0, total=0.0, data_arrecadacao=None)

            m = re.search(r"(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})\s*(.+)", texto)
            if m:
                d["cnpj"] = m.group(1).strip()
                d["razao_social"] = m.group(2).strip()

            datas = re.findall(r"\d{2}/\d{2}/\d{4}", texto)
            if len(datas) >= 1: d["periodo_apuracao"] = parse_data_br(datas[0])
            if len(datas) >= 2: d["data_vencimento"]  = parse_data_br(datas[1])
            if len(datas) >= 3: d["data_arrecadacao"] = parse_data_br(datas[2])

            nums = re.findall(r"\b0\d{16}\b", texto)
            if nums: d["numero_documento"] = nums[0]

            m2 = re.search(
                r"^(\d{4})\s+(CSLL|IRPJ|COFINS|PIS|IRRF|IOF|IPI|CSRF)(.+?)"
                r"([\d.,]+)\s+([\d.,\-–]+)\s+([\d.,\-–]+)\s+([\d.,]+)\s*$",
                texto, re.MULTILINE)
            if m2:
                d["codigo"]    = m2.group(1)
                d["descricao"] = (m2.group(2) + m2.group(3)).strip()
                d["principal"] = parse_valor(m2.group(4))
                d["multa"]     = parse_valor(m2.group(5))
                d["juros"]     = parse_valor(m2.group(6))
                d["total"]     = parse_valor(m2.group(7))

            if d["total"] == 0.0:
                m3 = re.search(r"Totais\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)", texto)
                if m3:
                    d["principal"] = parse_valor(m3.group(1))
                    d["multa"]     = parse_valor(m3.group(2))
                    d["juros"]     = parse_valor(m3.group(3))
                    d["total"]     = parse_valor(m3.group(4))

            if d["codigo"]:
                d["tributo"] = CODIGO_TRIBUTO.get(d["codigo"], f"CÓD {d['codigo']}")

            darfs.append(d)
    return darfs

# ─────────────────────────────────────────────────────────────────────────────
# Extração DCTFWeb
# ─────────────────────────────────────────────────────────────────────────────

TRIBUTOS_RECIBO = [
    "Contribuição Previdenciária Segurados",
    "Contribuição Previdenciária Patronal",
    "Contribuição para Outras Entidades e Fundos",
    "Contribuições Diversas",
    "COFINS","COSIRF","CPSS","CSLL","CSRF",
    "IOF","IPI","IRPJ","IRRF","PIS",
    "RET/Pagamento Unificado",
]

def extrair_dctfweb(pdf_bytes: bytes) -> dict:
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        texto = "\n".join(p.extract_text() or "" for p in pdf.pages)

    d = dict(formato=None, cnpj=None, razao_social=None,
             competencia_str=None, competencia_dt=None,
             numero_recibo=None, data_transmissao=None,
             retificadora=None, situacao=None, tributos={})

    m = re.search(r"(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})", texto)
    if m: d["cnpj"] = m.group(1)

    m = re.search(r"(?:Nome/Razão Social|Nome)\s*(.+)", texto)
    if m: d["razao_social"] = m.group(1).strip()

    m = re.search(r"Número Recibo\s*(\d+)|Nº do recibo de entrega\s*(\d+)", texto)
    if m: d["numero_recibo"] = m.group(1) or m.group(2)

    m = re.search(r"Data da Transmissão\s*(\d{2}/\d{2}/\d{4})|recebida[^\d]+(\d{2}/\d{2}/\d{4})", texto)
    if m: d["data_transmissao"] = parse_data_br(m.group(1) or m.group(2))

    m = re.search(r"Declaração Retificadora\s*(Sim|Não)", texto)
    if m: d["retificadora"] = m.group(1)

    m = re.search(r"Situação Declaração\s*(\S+)", texto)
    if m: d["situacao"] = m.group(1)

    m = re.search(r"Período de apuração\s*(\d{2}/\d{4})", texto)
    if not m:
        m = re.search(r"Geral\s*[-–]\s*(\d{2}/\d{4})", texto, re.IGNORECASE)
    if m:
        comp = m.group(1)
        d["competencia_str"] = comp
        d["competencia_dt"]  = competencia_para_dt(comp)

    d["formato"] = (
        "Recibo de Entrega" if "Recibo de Entrega" in texto
        else "Extrato do Processamento" if "Extrato do Processamento" in texto
        else "Desconhecido"
    )

    if d["formato"] == "Recibo de Entrega":
        bloco_m = re.search(
            r"Totalização dos tributos apurados no período(.+?)O presente Recibo",
            texto, re.DOTALL
        )
        bloco = bloco_m.group(1) if bloco_m else texto

        for trib in TRIBUTOS_RECIBO:
            linha_m = re.search(re.escape(trib) + r"[^\n]*", bloco)
            if not linha_m:
                continue
            linha = linha_m.group(0).strip()
            resto = linha[len(trib):].strip()

            if not resto:
                d["tributos"][trib] = {"debito": 0.0, "saldo": 0.0, "sem_movimento": True}

            elif "Sem Movimento" in resto:
                d["tributos"][trib] = {"debito": 0.0, "saldo": 0.0, "sem_movimento": True}

            else:
                nums = []
                for v in re.findall(r"[\d]+(?:[.,]\d+)*", resto):
                    try:
                        nums.append(float(v.replace(".", "").replace(",", ".")))
                    except Exception:
                        pass

                if nums:
                    d["tributos"][trib] = {
                        "debito": nums[0],
                        "saldo":  nums[1] if len(nums) > 1 else nums[0],
                        "sem_movimento": False,
                    }
                else:
                    d["tributos"][trib] = {"debito": 0.0, "saldo": 0.0, "sem_movimento": True}

    elif d["formato"] == "Extrato do Processamento":
        for linha in texto.splitlines():
            m2 = re.match(
                r"\s*(\d{4})\s+(.+?)\s+([\d.,]+)\s+([\d.,\-]+)\s+([\d.,\-]+)\s+([\d.,]+)\s*$", linha)
            if m2:
                d["tributos"][m2.group(2).strip()] = {
                    "debito": parse_valor(m2.group(3)),
                    "saldo":  parse_valor(m2.group(6)),
                    "sem_movimento": False,
                }
    return d

# ─────────────────────────────────────────────────────────────────────────────
# Cruzamento
# ─────────────────────────────────────────────────────────────────────────────

def cruzar(darfs: list[dict], dctfwebs: list[dict]) -> list[dict]:
    idx = {chave(d.get("competencia_dt")): d for d in dctfwebs if d.get("competencia_dt")}
    disponiveis = ", ".join(sorted(idx.keys()))

    # ── Agrupa DARFs complementares por (competência + tributo) ──────────────
    # Chave: (competencia_str, tributo)
    # Valor: lista de DARFs do mesmo grupo
    grupos: dict[tuple, list[dict]] = {}
    for darf in darfs:
        pa      = darf.get("periodo_apuracao")
        ck      = chave(pa)
        tributo = darf.get("tributo") or "N/D"
        grupos.setdefault((ck, tributo), []).append(darf)

    resultados = []
    for (ck, tributo), grupo in grupos.items():
        # Usa o primeiro DARF como referência para os campos de identificação
        darf_ref = grupo[0]
        pa       = darf_ref.get("periodo_apuracao")
        periodo_fmt = pa.strftime("%d/%m/%Y") if pa else "N/D"

        # Soma os valores de todos os DARFs do grupo
        principal_total = round(sum(d["principal"] for d in grupo), 2)
        multa_total     = round(sum(d["multa"]     for d in grupo), 2)
        juros_total     = round(sum(d["juros"]     for d in grupo), 2)
        total_total     = round(sum(d["total"]      for d in grupo), 2)
        tem_multa       = multa_total > 0 or juros_total > 0

        # Números de documento: lista todos separados por vírgula se houver mais de um
        numeros_doc = ", ".join(
            d["numero_documento"] for d in grupo if d.get("numero_documento")
        ) or "N/D"

        # Data de arrecadação: a mais recente do grupo
        datas_arr = [d["data_arrecadacao"] for d in grupo if d.get("data_arrecadacao")]
        data_arr_fmt = max(datas_arr).strftime("%d/%m/%Y") if datas_arr else "N/D"

        dctf = idx.get(ck)

        if dctf is None:
            status = "sem_dctf"
            obs = f"Nenhuma DCTFWeb carregada para {ck}. Disponíveis: {disponiveis or '—'}"
            val_declarado = None
        else:
            info = dctf["tributos"].get(tributo)
            if info is None:
                status = "nao_encontrado"
                obs = f"Tributo '{tributo}' não encontrado na DCTFWeb {ck}"
                val_declarado = None
            elif info.get("sem_movimento"):
                status = "sem_movimento"
                obs = f"DCTFWeb {ck} declara 'Sem Movimento' para {tributo}"
                val_declarado = 0.0
            else:
                val_declarado = info.get("debito", 0.0)
                # Compara a SOMA dos principais com o valor declarado na DCTFWeb
                diff = round(principal_total - val_declarado, 2)
                n    = len(grupo)
                comp_obs = f" (soma de {n} DARFs)" if n > 1 else ""
                if diff == 0:
                    status = "ok"
                    obs = f"Valor coincide com a DCTFWeb{comp_obs}"
                elif diff > 0:
                    status = "divergente"
                    obs = f"Principal DARF maior que declarado em R$ {diff:.2f}{comp_obs}"
                else:
                    status = "divergente"
                    obs = f"Principal DARF menor que declarado em R$ {abs(diff):.2f}{comp_obs}"

        resultados.append(dict(
            competencia_darf = ck or "N/D",
            periodo_apuracao = periodo_fmt,
            numero_documento = numeros_doc,
            tributo          = tributo,
            codigo           = darf_ref.get("codigo") or "N/D",
            principal        = principal_total,
            multa            = multa_total,
            juros            = juros_total,
            total            = total_total,
            tem_multa        = tem_multa,
            data_arrecadacao = data_arr_fmt,
            competencia_dctf = dctf["competencia_str"] if dctf else "—",
            recibo_dctf      = dctf.get("numero_recibo", "—") if dctf else "—",
            val_declarado    = val_declarado,
            status           = status,
            obs              = obs,
        ))
    return resultados

# ─────────────────────────────────────────────────────────────────────────────
# Excel
# ─────────────────────────────────────────────────────────────────────────────

STATUS_LABEL = {
    "ok":            "✅ Declarado OK",
    "sem_movimento": "❌ Sem Movimento na DCTFWeb",
    "sem_dctf":      "📭 DCTFWeb não fornecida para este período",
    "divergente":    "⚠️ Valor divergente",
    "nao_encontrado":"⚠️ Tributo não encontrado",
}

def gerar_excel(resultados: list[dict]) -> bytes:
    rows = [{
        "Competência DARF":     r["competencia_darf"],
        "Período Apuração":     r["periodo_apuracao"],
        "Nº Documento DARF":    r["numero_documento"],
        "Tributo":              r["tributo"],
        "Código":               r["codigo"],
        "Principal (R$)":       r["principal"],
        "Multa (R$)":           r["multa"],
        "Juros (R$)":           r["juros"],
        "Total DARF (R$)":      r["total"],
        "Com Multa/Juros":      "Sim ⚠️" if r["tem_multa"] else "Não",
        "Data Arrecadação":     r["data_arrecadacao"],
        "Competência DCTFWeb":  r["competencia_dctf"],
        "Nº Recibo DCTFWeb":    r["recibo_dctf"],
        "Valor Declarado (R$)": r["val_declarado"] if r["val_declarado"] is not None else "",
        "Status":               STATUS_LABEL.get(r["status"], r["status"]),
        "Observação":           r["obs"],
    } for r in resultados]

    df = pd.DataFrame(rows)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Cruzamento")
        ws = writer.sheets["Cruzamento"]
        verde = PatternFill("solid", fgColor="C6EFCE")
        verm  = PatternFill("solid", fgColor="FFC7CE")
        amar  = PatternFill("solid", fgColor="FFEB9C")
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(wrap_text=True)
        sc = list(df.columns).index("Status")
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            v = str(row[sc].value or "")
            f = verde if "✅" in v else (verm if "❌" in v else amar)
            for cell in row: cell.fill = f
        for col in ws.columns:
            w = max(len(str(c.value or "")) for c in col)
            ws.column_dimensions[col[0].column_letter].width = min(w + 3, 52)
    return buf.getvalue()

# ─────────────────────────────────────────────────────────────────────────────
# Helpers de HTML
# ─────────────────────────────────────────────────────────────────────────────

def badge(status: str) -> str:
    mp = {
        "ok":            ("badge-ok",    "✅ Declarado OK"),
        "sem_movimento": ("badge-erro",  "❌ Sem Movimento"),
        "sem_dctf":      ("badge-info",  "📭 DCTFWeb ausente"),
        "divergente":    ("badge-alert", "⚠️ Divergente"),
        "nao_encontrado":("badge-warn",  "⚠️ Não encontrado"),
    }
    cls, label = mp.get(status, ("badge-warn", status))
    return f'<span class="badge {cls}">{label}</span>'

def card_metric(valor, label, sub, classe) -> str:
    return f"""
    <div class="metric-card {classe}">
        <div class="val">{valor}</div>
        <div class="lbl">{label}</div>
        <div class="sub">{sub}</div>
    </div>"""

# ─────────────────────────────────────────────────────────────────────────────
# INTERFACE
# ─────────────────────────────────────────────────────────────────────────────

st.markdown("""
<div class="app-header">
  <div style="font-size:2.4rem">📊</div>
  <div>
    <h1>Cruzamento DARF × DCTFWeb</h1>
    <p>Validação automática de recolhimentos declarados — identifique divergências em segundos</p>
  </div>
</div>
""", unsafe_allow_html=True)

# ── Upload ────────────────────────────────────────────────────────────────────
col_up1, col_up2 = st.columns(2, gap="large")

with col_up1:
    st.markdown("#### 🧾 DARFs")
    st.caption("Selecione um ou mais PDFs de comprovantes DARF")
    darf_uploads = st.file_uploader(
        "PDFs dos DARFs", type="pdf",
        accept_multiple_files=True, key="darf",
        label_visibility="collapsed"
    )
    if darf_uploads:
        for f in darf_uploads:
            st.markdown(f'<span class="pill">📄 {f.name}</span>', unsafe_allow_html=True)

with col_up2:
    st.markdown("#### 📋 DCTFWeb")
    st.caption("Selecione um ou mais recibos / extratos de DCTFWeb")
    dctf_uploads = st.file_uploader(
        "PDFs da DCTFWeb", type="pdf",
        accept_multiple_files=True, key="dctf",
        label_visibility="collapsed"
    )
    if dctf_uploads:
        for f in dctf_uploads:
            st.markdown(f'<span class="pill">📄 {f.name}</span>', unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ── Botão ─────────────────────────────────────────────────────────────────────
btn_disabled = not (darf_uploads and dctf_uploads)

if btn_disabled:
    if not darf_uploads and not dctf_uploads:
        st.markdown('<div class="info-box">👆 Carregue os PDFs dos DARFs e da DCTFWeb para começar.</div>', unsafe_allow_html=True)
    elif not darf_uploads:
        st.markdown('<div class="warn-box">📂 Carregue pelo menos um PDF de DARF.</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="warn-box">📂 Carregue pelo menos um PDF de DCTFWeb.</div>', unsafe_allow_html=True)

processar = st.button("🔍  Cruzar e Analisar", type="primary",
                      use_container_width=True, disabled=btn_disabled)

# ── Processamento ─────────────────────────────────────────────────────────────
if processar:
    with st.spinner("Lendo e cruzando os dados..."):

        # DARFs — todos os PDFs combinados
        todos_darfs = []
        for f in darf_uploads:
            try:
                todos_darfs.extend(extrair_darfs(f.read()))
            except Exception as e:
                st.error(f"Erro ao ler DARF '{f.name}': {e}")
                st.stop()

        # DCTFWebs
        dctfwebs = []
        for f in dctf_uploads:
            try:
                dctfwebs.append(extrair_dctfweb(f.read()))
            except Exception as e:
                st.warning(f"⚠️ Não foi possível ler '{f.name}': {e}")

        resultados = cruzar(todos_darfs, dctfwebs)

    if not resultados:
        st.error("Nenhum dado foi extraído. Verifique os PDFs enviados.")
        st.stop()

    # ── Empresa e competências ────────────────────────────────────────────────
    emp_nome = (dctfwebs[0].get("razao_social") if dctfwebs else None) or \
               (todos_darfs[0].get("razao_social") if todos_darfs else "—")
    emp_cnpj = (dctfwebs[0].get("cnpj") if dctfwebs else None) or \
               (todos_darfs[0].get("cnpj") if todos_darfs else "—")
    comps_dctf = sorted({d["competencia_str"] for d in dctfwebs if d.get("competencia_str")})
    comps_darf = sorted({r["competencia_darf"] for r in resultados if r["competencia_darf"] != "N/D"})

    st.divider()

    # Cabeçalho empresa
    ei1, ei2, ei3, ei4 = st.columns([3, 2, 2, 2])
    ei1.markdown(f"**{emp_nome}**")
    ei2.markdown(f"CNPJ: `{emp_cnpj}`")
    ei3.markdown(f"DARFs carregados: **{len(todos_darfs)}**")
    ei4.markdown(f"Competências DCTFWeb: **{', '.join(comps_dctf) or '—'}**")

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Cards de resumo ───────────────────────────────────────────────────────
    n_ok        = sum(1 for r in resultados if r["status"] == "ok")
    n_sem_mov   = sum(1 for r in resultados if r["status"] == "sem_movimento")
    n_sem_dctf  = sum(1 for r in resultados if r["status"] == "sem_dctf")
    n_diverg    = sum(1 for r in resultados if r["status"] == "divergente")
    n_multa     = sum(1 for r in resultados if r["tem_multa"])
    total_rec   = sum(r["total"] for r in resultados)

    # Valor total da divergência: diferença absoluta entre principal e declarado
    total_diverg = round(sum(
        abs(r["principal"] - (r["val_declarado"] or 0))
        for r in resultados
        if r["status"] == "divergente" and r["val_declarado"] is not None
    ), 2)

    c1,c2,c3,c4,c5,c6 = st.columns(6, gap="small")
    cards = [
        (c1, n_ok,                "Declarados OK",          f"de {len(resultados)} grupos",    "card-ok"),
        (c2, n_sem_mov,           "Sem Movimento",           "na DCTFWeb",                      "card-erro"),
        (c3, n_sem_dctf,          "DCTFWeb ausente",         "período não carregado",            "card-info"),
        (c4, fmt_brl(total_diverg), "Divergência DARF × DCTFWeb", f"{n_diverg} grupo(s) divergente(s)", "card-alert"),
        (c5, n_multa,             "Com multa/juros",         "recolhidos em atraso",            "card-warn"),
        (c6, fmt_brl(total_rec),  "Total recolhido",         f"{len(resultados)} grupos de DARFs", "card-info"),
    ]
    for col, val, lbl, sub, cls in cards:
        col.markdown(card_metric(val, lbl, sub, cls), unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Painel de atenção — valores sem DCTFWeb ───────────────────────────────
    problemas = [r for r in resultados if r["status"] != "ok"]
    if problemas:
        st.markdown("### 🔴 Itens que precisam de atenção")

        grupos = {}
        for r in problemas:
            grupos.setdefault(r["status"], []).append(r)

        tab_labels = []
        tab_grupos = []
        ordem = ["sem_dctf", "sem_movimento", "divergente", "nao_encontrado"]
        nomes  = {
            "sem_dctf":       "📭 DCTFWeb ausente",
            "sem_movimento":  "❌ Sem Movimento",
            "divergente":     "⚠️ Valor divergente",
            "nao_encontrado": "⚠️ Não encontrado",
        }
        for k in ordem:
            if k in grupos:
                tab_labels.append(f"{nomes[k]} ({len(grupos[k])})")
                tab_grupos.append((k, grupos[k]))

        tabs = st.tabs(tab_labels)
        for tab, (status_key, items) in zip(tabs, tab_grupos):
            with tab:
                total_grupo = sum(r["total"] for r in items)
                st.markdown(
                    f'<div class="warn-box">💰 Total em aberto neste grupo: <strong>{fmt_brl(total_grupo)}</strong></div>',
                    unsafe_allow_html=True
                )
                st.markdown("<br>", unsafe_allow_html=True)

                cols_cards = st.columns(min(len(items), 4), gap="small")
                for i, r in enumerate(items):
                    with cols_cards[i % 4]:
                        multa_html = '<span class="multa-tag">MULTA</span>' if r["tem_multa"] else ""
                        st.markdown(f"""
                        <div class="sem-dctf-card">
                            <div class="comp">{r["tributo"]} · {r["competencia_darf"]}</div>
                            <div class="val">{fmt_brl(r["total"])} {multa_html}</div>
                            <div class="desc">PA: {r["periodo_apuracao"]}</div>
                            <div class="desc">DARF: <code>{r["numero_documento"]}</code></div>
                            <div class="desc" style="margin-top:6px;color:#9a3412">{r["obs"]}</div>
                        </div>
                        """, unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

    # ── Tabela completa ───────────────────────────────────────────────────────
    st.markdown("### 📋 Resultado completo por DARF")

    filtro_col1, filtro_col2 = st.columns([2, 1])
    with filtro_col1:
        filtro_status = st.multiselect(
            "Filtrar por status",
            options=["ok","sem_dctf","sem_movimento","divergente","nao_encontrado"],
            format_func=lambda x: STATUS_LABEL.get(x, x),
            default=None,
            placeholder="Todos os status"
        )
    with filtro_col2:
        filtro_comp = st.multiselect(
            "Filtrar por competência",
            options=sorted({r["competencia_darf"] for r in resultados}),
            placeholder="Todas"
        )

    res_filtrado = resultados
    if filtro_status:
        res_filtrado = [r for r in res_filtrado if r["status"] in filtro_status]
    if filtro_comp:
        res_filtrado = [r for r in res_filtrado if r["competencia_darf"] in filtro_comp]

    colunas_html = [
        "Competência", "Período AP.", "Nº DARF", "Tributo",
        "Principal", "Multa", "Juros", "Total DARF",
        "Arrecadação", "Comp. DCTFWeb", "Recibo DCTFWeb", "Status"
    ]
    thead = "".join(f"<th>{c}</th>" for c in colunas_html)
    tbody = ""

    cor_map = {
        "ok":            "#f0fdf4",
        "sem_movimento": "#fef2f2",
        "sem_dctf":      "#eff6ff",
        "divergente":    "#fffbeb",
        "nao_encontrado":"#fff7ed",
    }
    for r in res_filtrado:
        cor = cor_map.get(r["status"], "#ffffff")
        multa_tag = '<span class="multa-tag">MULTA</span>' if r["tem_multa"] else ""
        tds = "".join([
            f"<td><strong>{r['competencia_darf']}</strong></td>",
            f"<td>{r['periodo_apuracao']}</td>",
            f"<td style='font-family:monospace;font-size:.72rem'>{r['numero_documento']}</td>",
            f"<td><strong>{r['tributo']}</strong> <span style='color:#94a3b8;font-size:.73rem'>{r['codigo']}</span></td>",
            f"<td>{fmt_brl(r['principal'])}</td>",
            f"<td>{fmt_brl(r['multa']) if r['multa'] else '—'}</td>",
            f"<td>{fmt_brl(r['juros']) if r['juros'] else '—'}</td>",
            f"<td><strong>{fmt_brl(r['total'])}</strong> {multa_tag}</td>",
            f"<td>{r['data_arrecadacao']}</td>",
            f"<td>{r['competencia_dctf']}</td>",
            f"<td style='font-size:.72rem'>{r['recibo_dctf']}</td>",
            f"<td>{badge(r['status'])}</td>",
        ])
        tbody += f"<tr style='background:{cor}'>{tds}</tr>"

    st.markdown(f"""
    <div class="tabela-wrap">
      <div style="overflow-x:auto">
        <table>
          <thead><tr>{thead}</tr></thead>
          <tbody>{tbody}</tbody>
        </table>
      </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown(f"<p style='color:#94a3b8;font-size:.78rem;margin-top:8px'>{len(res_filtrado)} de {len(resultados)} registros exibidos</p>", unsafe_allow_html=True)

    # ── Export ────────────────────────────────────────────────────────────────
    st.markdown("<br>", unsafe_allow_html=True)
    excel_bytes = gerar_excel(resultados)
    st.download_button(
        label="⬇️  Exportar resultado completo em Excel",
        data=excel_bytes,
        file_name="cruzamento_darf_dctfweb.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary",
    )
