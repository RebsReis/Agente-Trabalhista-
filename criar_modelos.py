#!/usr/bin/env python3
"""Recria os modelos em modelos/ a partir dos documentos originais do escritório.

Só é preciso rodar uma vez (ou quando o texto-base mudar):
    python3 criar_modelos.py PROCURACAO_AD_JUDICIA_DEIVID.docx modelos/procuracao_modelo.docx
    python3 criar_modelos.py "Contrato de Honorario DIRCEU.docx" modelos/contrato_honorarios_modelo.docx

Troca os dados do cliente do exemplo pelos campos {{...}}, mantendo o timbre
(cabeçalho e rodapé) e a formatação, e remove a miniatura com dados antigos.
"""
import os, re, sys, zipfile
import docx

src, out = sys.argv[1], sys.argv[2]

def idx(p, pred, start=0):
    for i, r in enumerate(p.runs):
        if i >= start and pred(r.text): return i
    raise SystemExit(f"run not found in: {p.text[:60]}")

def setr(p, a, b, text):
    """Put text in run a, empty runs a+1..b."""
    p.runs[a].text = text
    for r in p.runs[a+1:b+1]: r.text = ""

def find_par(d, needle, start=0):
    for i, p in enumerate(d.paragraphs):
        if i >= start and needle in p.text: return i, p
    raise SystemExit("par not found: " + needle)

d = docx.Document(src)
if "PROCURA" in os.path.basename(src).upper():
    _, p = find_par(d, "Pelo presente instrumento particular de mandato")
    n = idx(p, lambda t: "DEIVIDI" in t); p.runs[n].text = "{{NOME}}"
    k = idx(p, lambda t: t.startswith("nomeia"))
    setr(p, n+1, k-1, ", {{QUALIFICACAO}}, ")
    a = idx(p, lambda t: "AÇÃO INDENIZATORIA" in t)
    p.runs[a].text = " {{ACAO}}."
    i, p = find_par(d, "11 de agosto"); setr(p, 0, len(p.runs)-1, "{{CIDADE}}, {{DATA}}.")
    i, p = find_par(d, "11 de agosto", i+1); setr(p, 0, len(p.runs)-1, "{{CIDADE}}, {{DATA}}.")
    for p in d.paragraphs:
        if p.text.strip() == "DEIVIDI LEANDRO DOS REIS": setr(p, 0, len(p.runs)-1, "{{NOME}}")
    i, p = find_par(d, "DECLARAÇÃO DE HIPOSSUFICIÊNCIA")
    p.paragraph_format.page_break_before = True
    _, p = find_par(d, "Eu, DEIVIDI")
    n = idx(p, lambda t: "DEIVIDI" in t); p.runs[n].text = "{{NOME}}"
    k = idx(p, lambda t: t.startswith("DECLARO"))
    setr(p, n+1, k-1, ", {{QUALIFICACAO}}, ")
else:
    _, p = find_par(d, "Pelo presente instrumento particular de contrato")
    n = idx(p, lambda t: "DIRCEU" in t); p.runs[n].text = "{{NOME}}"
    p.runs[n+1].text = ", {{QUALIFICACAO}}"
    _, p = find_par(d, "no que tange a propor")
    a = idx(p, lambda t: t.startswith("AÇÃO")); setr(p, a, a+1, "{{ACAO}},")
    _, p = find_par(d, "A título de pagamento")
    setr(p, 1, len(p.runs)-1, "o contratante pagará a contratada o percentual de {{PERCENTUAL}}, sobre o valor auferido na demanda, seja por acordo ou sentença.")
    _, p = find_par(d, "forma parcelada")
    a = idx(p, lambda t: t == "valor correspondente a ")
    setr(p, a, len(p.runs)-1, "valor correspondente a {{PERCENTUAL}} do montante auferido, inclusive sobre os valores auferidos no levantamento ou recebimento de FGTS e multa fundiária.")
    _, p = find_par(d, "Seguro Desemprego")
    p.runs[0].text = p.runs[0].text.replace("Pagará ainda 30%", "Pagará ainda {{PERCENTUAL_NUM}}%")
    assert "{{PERCENTUAL_NUM}}" in p.text
    _, p = find_par(d, "As partes elegem o foro")
    setr(p, 0, len(p.runs)-1, "As partes elegem o foro {{FORO}}, para dirimir eventuais controvérsias futuras.")
    _, p = find_par(d, "São Paulo, 07")
    setr(p, 0, len(p.runs)-1, "{{CIDADE}}, {{DATA}}.")
# limpa metadados
d.core_properties.author = "Reis & Oliveira Advogados"; d.core_properties.last_modified_by = ""; d.core_properties.title = ""
d.save(out)

# remove a miniatura (pode mostrar dados do cliente do exemplo)
tmp = out + ".tmp"
with zipfile.ZipFile(out) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
    for it in zin.infolist():
        if it.filename == "docProps/thumbnail.emf":
            continue
        data = zin.read(it.filename)
        if it.filename == "_rels/.rels":
            data = re.sub(rb'<Relationship [^>]*thumbnail[^>]*/>', b"", data)
        zout.writestr(it, data)
os.replace(tmp, out)

for p in d.paragraphs:
    if p.text.strip(): print("|", p.text)
