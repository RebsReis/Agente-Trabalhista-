#!/usr/bin/env python3
"""Gera a Procuração (+ Declaração de Hipossuficiência) e o Contrato de Honorários
de um cliente novo, no papel timbrado do escritório Reis & Oliveira.

Uso:
    python3 gerar_documentos.py dados_cliente.json [--sem-pdf]

Os arquivos são salvos em clientes/<NOME_DO_CLIENTE>/ (pasta ignorada pelo git,
porque contém dados pessoais).
"""
import argparse
import datetime as dt
import json
import re
import shutil
import subprocess
import sys
import unicodedata
from pathlib import Path

import docx

RAIZ = Path(__file__).resolve().parent
MODELO_PROCURACAO = RAIZ / "modelos" / "procuracao_modelo.docx"
MODELO_CONTRATO = RAIZ / "modelos" / "contrato_honorarios_modelo.docx"

MESES = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho",
         "agosto", "setembro", "outubro", "novembro", "dezembro"]

UNIDADES = ["zero", "um", "dois", "três", "quatro", "cinco", "seis", "sete", "oito",
            "nove", "dez", "onze", "doze", "treze", "quatorze", "quinze", "dezesseis",
            "dezessete", "dezoito", "dezenove"]
DEZENAS = ["", "", "vinte", "trinta", "quarenta", "cinquenta", "sessenta", "setenta",
           "oitenta", "noventa"]


def extenso(n: int) -> str:
    if n < 20:
        return UNIDADES[n]
    if n < 100:
        d, u = divmod(n, 10)
        return DEZENAS[d] + ("" if u == 0 else " e " + UNIDADES[u])
    if n == 100:
        return "cem"
    raise ValueError("percentual deve ser entre 0 e 100")


def data_extenso(d: dt.date) -> str:
    return f"{d.day:02d} de {MESES[d.month - 1]} de {d.year}"


def fem(dados) -> bool:
    return str(dados.get("genero", "M")).strip().upper().startswith("F")


def formatar_endereco(dados) -> str:
    if dados.get("endereco"):
        return dados["endereco"].strip().rstrip(",")
    partes = f"{dados['logradouro']}, n° {dados['numero']}"
    if dados.get("complemento"):
        partes += f", {dados['complemento']}"
    partes += f" - {dados['bairro']}, {dados['cidade']} - {dados['uf']}"
    return partes


def qualificacao(dados) -> str:
    f = fem(dados)
    nac = dados.get("nacionalidade") or ("brasileira" if f else "brasileiro")
    rg = dados["rg"]
    if dados.get("orgao_emissor"):
        rg += f" {dados['orgao_emissor']}"
    itens = [
        nac,
        dados["estado_civil"],
        dados["profissao"],
        f"{'portadora' if f else 'portador'} da cédula de identidade RG nº {rg}",
        f"{'inscrita' if f else 'inscrito'} no CPF sob o nº {dados['cpf']}",
        f"residente e {'domiciliada' if f else 'domiciliado'} na "
        f"{formatar_endereco(dados)}, CEP: {dados['cep']}",
    ]
    return ", ".join(i.strip() for i in itens if i and i.strip())


def substituir(doc, valores):
    def trocar(paragrafos):
        for p in paragrafos:
            for r in p.runs:
                if "{{" in r.text:
                    t = r.text
                    for k, v in valores.items():
                        t = t.replace("{{" + k + "}}", v)
                    r.text = t

    trocar(doc.paragraphs)
    for tabela in doc.tables:
        for linha in tabela.rows:
            for celula in linha.cells:
                trocar(celula.paragraphs)
    restantes = [p.text for p in doc.paragraphs if "{{" in p.text]
    if restantes:
        raise SystemExit(f"Campos não preenchidos: {restantes}")


def remover(paragrafo):
    paragrafo._p.getparent().remove(paragrafo._p)


def remover_com_vazios_seguintes(paragrafo):
    """Remove o parágrafo e as linhas em branco logo depois dele."""
    el = paragrafo._p
    prox = el.getnext()
    el.getparent().remove(el)
    # remove uma linha em branco (o espaçamento entre parágrafos), se houver
    if (prox is not None and prox.tag.endswith("}p")
            and not "".join(prox.itertext()).strip() and "sectPr" not in prox.xml):
        prox.getparent().remove(prox)


def remover_declaracao(doc):
    corpo = doc.element.body
    paragrafos = doc.paragraphs
    inicio = next(i for i, p in enumerate(paragrafos) if "DECLARAÇÃO DE HIPOSSUFICIÊNCIA" in p.text)
    # recua sobre as linhas em branco que antecedem o título
    while inicio > 0 and not paragrafos[inicio - 1].text.strip():
        inicio -= 1
    for p in paragrafos[inicio:]:
        if p._p.getparent() is corpo and p._p.find(
                "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr/"
                "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sectPr") is None:
            corpo.remove(p._p)


def remover_clausulas_trabalhistas(doc):
    for chave in ("recurso pelo INSS", "forma parcelada", "Seguro Desemprego"):
        for p in list(doc.paragraphs):
            if chave in p.text:
                remover_com_vazios_seguintes(p)


def nome_pasta(nome: str) -> str:
    s = unicodedata.normalize("NFKD", nome).encode("ascii", "ignore").decode()
    return re.sub(r"[^A-Za-z0-9]+", "_", s).strip("_").upper()


def converter_pdf(arquivo: Path):
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        print("  (LibreOffice não encontrado: PDF não gerado)")
        return
    subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir",
                    str(arquivo.parent), str(arquivo)],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("dados", help="arquivo JSON com os dados do cliente")
    ap.add_argument("--sem-pdf", action="store_true", help="não gerar as versões em PDF")
    ap.add_argument("--saida", help="pasta de saída (padrão: clientes/<NOME>)")
    args = ap.parse_args()

    dados = json.loads(Path(args.dados).read_text(encoding="utf-8"))
    obrigatorios = ["nome", "estado_civil", "profissao", "rg", "cpf", "cep", "acao"]
    if not dados.get("endereco"):
        obrigatorios += ["logradouro", "numero", "bairro", "cidade", "uf"]
    faltando = [c for c in obrigatorios if not str(dados.get(c, "")).strip()]
    if faltando:
        sys.exit(f"Faltam dados obrigatórios: {', '.join(faltando)}")

    data = dt.date.fromisoformat(dados["data"]) if dados.get("data") else dt.date.today()
    pct = int(dados.get("percentual", 30))
    # nome da ação exatamente como deve aparecer (ex.: "AÇÃO TRABALHISTA")
    acao = dados["acao"].strip().rstrip(".,").upper()

    valores = {
        "NOME": dados["nome"].strip().upper(),
        "QUALIFICACAO": qualificacao(dados),
        "ACAO": acao,
        "CIDADE": dados.get("cidade_assinatura", "São Paulo"),
        "DATA": data_extenso(data),
        "PERCENTUAL": f"{pct}% ({extenso(pct)} por cento)",
        "PERCENTUAL_NUM": str(pct),
        "FORO": dados.get("foro", "Regional de Itaquera, desta Comarca"),
    }

    pasta = Path(args.saida) if args.saida else RAIZ / "clientes" / nome_pasta(valores["NOME"])
    pasta.mkdir(parents=True, exist_ok=True)
    sufixo = nome_pasta(valores["NOME"])
    gerados = []

    if dados.get("gerar_procuracao", True):
        doc = docx.Document(MODELO_PROCURACAO)
        if not dados.get("hipossuficiencia", True):
            remover_declaracao(doc)
        substituir(doc, valores)
        destino = pasta / f"PROCURACAO_AD_JUDICIA_{sufixo}.docx"
        doc.save(destino)
        gerados.append(destino)

    if dados.get("gerar_contrato", True):
        doc = docx.Document(MODELO_CONTRATO)
        trabalhista = dados.get("clausulas_trabalhistas")
        if trabalhista is None:
            trabalhista = "TRABALHISTA" in acao or "RECLAMA" in acao
        if not trabalhista:
            remover_clausulas_trabalhistas(doc)
        substituir(doc, valores)
        destino = pasta / f"CONTRATO_HONORARIOS_{sufixo}.docx"
        doc.save(destino)
        gerados.append(destino)

    for arq in gerados:
        if not args.sem_pdf:
            converter_pdf(arq)
        print(arq)
        if not args.sem_pdf and arq.with_suffix(".pdf").exists():
            print(arq.with_suffix(".pdf"))


if __name__ == "__main__":
    main()
