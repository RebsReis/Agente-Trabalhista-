# Agente de Procuração e Contrato

Envie os documentos do cliente novo (fotos/PDF de RG ou CNH, CPF e comprovante de
residência) numa conversa com o Claude Code neste repositório, por exemplo:

> Cliente novo, ação trabalhista, 30%. Segue os documentos.

O agente lê os documentos, pergunta só o que faltar (estado civil, profissão,
tipo de ação, percentual) e devolve, com o timbre do escritório:

- `PROCURACAO_AD_JUDICIA_<NOME>.docx/.pdf` — procuração + declaração de hipossuficiência
- `CONTRATO_HONORARIOS_<NOME>.docx/.pdf`

## Uso manual

```bash
pip install python-docx           # LibreOffice opcional, para gerar PDF
mkdir -p entrada && cp exemplos/dados_cliente_exemplo.json entrada/cliente.json   # edite os dados
python3 gerar_documentos.py entrada/cliente.json
```

Os arquivos saem em `clientes/<NOME>/`.

## Alterar o texto padrão

Edite os modelos em `modelos/` no Word, mantendo os campos entre chaves duplas
(`{{NOME}}`, `{{QUALIFICACAO}}`, `{{ACAO}}`, `{{CIDADE}}`, `{{DATA}}`,
`{{PERCENTUAL}}`, `{{PERCENTUAL_NUM}}`, `{{FORO}}`). Digite cada campo de uma vez,
sem formatar só parte dele, para o Word não quebrá-lo em pedaços.
