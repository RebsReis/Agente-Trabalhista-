---
name: novo-cliente
description: Gera a Procuração "ad judicia" (com Declaração de Hipossuficiência) e o Contrato de Honorários de um cliente novo, no papel timbrado da Reis & Oliveira Advogados. Use quando o advogado enviar documentos de um cliente novo (RG, CNH, CPF, comprovante de residência, ficha) ou pedir procuração, declaração ou contrato de honorários.
---

# Agente de Procuração e Contrato — Reis & Oliveira Advogados

Sempre que chegar um cliente novo, o advogado envia os documentos do cliente
e você devolve, prontos para imprimir e assinar:

1. **Procuração "ad judicia"** + **Declaração de Hipossuficiência** (mesmo arquivo)
2. **Contrato de Honorários**

Os dois já saem com o timbre do escritório (cabeçalho R&O e rodapé com endereço),
porque são gerados a partir dos modelos em `modelos/`.

## Passo a passo

### 1. Ler os documentos do cliente
Abra cada arquivo enviado (fotos ou PDF de RG, CNH, CPF, comprovante de
residência, ficha de atendimento) com a ferramenta Read e extraia:

| Campo | Onde costuma estar |
|---|---|
| `nome` (completo) | RG / CNH |
| `genero` (`M` ou `F`) | pelo documento/nome; na dúvida, pergunte |
| `rg` e `orgao_emissor` (ex.: `SSP/SP`) | RG / CNH |
| `cpf` (formato 000.000.000-00) | RG, CNH ou CPF |
| `logradouro`, `numero`, `complemento`, `bairro`, `cidade`, `uf`, `cep` | comprovante de residência |
| `nacionalidade` | padrão "brasileiro(a)" |

Confira dígitos de CPF/RG/CEP com cuidado: não invente nada. Se algo estiver
ilegível, pergunte.

### 2. Perguntar só o que os documentos não mostram
Faça **uma única pergunta agrupada** com o que faltar, sugerindo os padrões:

- **Estado civil** (já no gênero certo: solteiro/solteira, casado/casada, divorciado/divorciada, viúvo/viúva, em união estável)
- **Profissão**
- **Tipo de ação** — como deve aparecer no documento, em maiúsculas (ex.: `AÇÃO TRABALHISTA`, `AÇÃO INDENIZATÓRIA C/C PERDAS E DANOS`)
- **Percentual de honorários** — padrão **30%**
- **Declaração de hipossuficiência** (justiça gratuita)? — padrão **sim**
- **Foro do contrato** — padrão `Regional de Itaquera, desta Comarca`

Se o advogado já informou algo na mensagem (ex.: "trabalhista, 30%"), não pergunte de novo.

### 3. Gravar os dados
Escreva `entrada/<nome_cliente>.json` seguindo `exemplos/dados_cliente_exemplo.json`
(a pasta `entrada/` não vai para o git, pois tem dados pessoais).
Campos opcionais:
- `endereco`: endereço completo em uma linha (alternativa a logradouro/numero/bairro/cidade/uf)
- `data`: `AAAA-MM-DD` (padrão: hoje)
- `clausulas_trabalhistas`: `true`/`false`. Padrão automático: mantém as cláusulas
  de INSS, acordo parcelado/FGTS e Seguro-Desemprego só se a ação for trabalhista.
- `gerar_procuracao` / `gerar_contrato`: `false` para gerar só um dos documentos.

### 4. Gerar os documentos
```bash
python3 gerar_documentos.py entrada/<nome_cliente>.json
```
Saída em `clientes/<NOME_DO_CLIENTE>/`: `.docx` (editável) e `.pdf` de cada documento.

### 5. Conferir e entregar
- Converta o PDF em imagem (`pdftoppm -r 60 -png arquivo.pdf previa`) e olhe as
  páginas: nome, qualificação, ação, percentual e data corretos; nada cortado.
- Envie os arquivos `.docx` e `.pdf` ao advogado (SendUserFile, ou indique o caminho).
- Mostre um resumo curto da qualificação usada para ele conferir.

## Regras
- Nunca faça commit de `clientes/` nem `entrada/` (dados pessoais — já estão no `.gitignore`).
- Não altere o texto jurídico dos modelos por conta própria. Se o advogado pedir
  mudança permanente no texto, edite o modelo em `modelos/` (mantendo os campos
  `{{NOME}}`, `{{QUALIFICACAO}}`, `{{ACAO}}`, `{{CIDADE}}`, `{{DATA}}`,
  `{{PERCENTUAL}}`, `{{PERCENTUAL_NUM}}`, `{{FORO}}` cada um inteiro dentro de um único trecho de texto).
- Para um ajuste só deste cliente, edite o `.docx` gerado.
