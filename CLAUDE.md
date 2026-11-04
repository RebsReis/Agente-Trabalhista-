# Agente Trabalhista — Reis & Oliveira Advogados

Este repositório é o agente que prepara os documentos de cliente novo:
**Procuração "ad judicia" + Declaração de Hipossuficiência** e **Contrato de Honorários**,
já no papel timbrado do escritório.

Quando o advogado enviar documentos de um cliente (RG, CNH, CPF, comprovante de
residência) ou pedir procuração/contrato, siga a skill
`.claude/skills/novo-cliente/SKILL.md`.

- `modelos/` — modelos .docx com timbre e campos `{{...}}`
- `criar_modelos.py` — cria os modelos a partir dos documentos originais do escritório
- `gerar_documentos.py` — preenche os modelos a partir de um JSON e gera .docx + .pdf
- `exemplos/dados_cliente_exemplo.json` — formato dos dados
- `clientes/` e `entrada/` — dados pessoais; nunca versionar

Responda sempre em português do Brasil.
