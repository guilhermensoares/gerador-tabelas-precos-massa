# GERADOR DE TABELAS DE PREÇOS EM MASSA

**Correção V10:** CSV Protheus ajustado para terminar todas as linhas com `;`, mantendo 4 campos por linha, conforme modelo manual aceito pelo importador DA1.


Este app gera arquivos CSV no modelo de tabelas de preços em massa do Protheus e mantém relatórios analíticos em Excel.

A versão atual possui uma **Home** com dois módulos:

1. **Precificação Comum**
2. **Adequação ML**

## Como rodar localmente

No Windows, extraia o ZIP e clique em:

```txt
executar_app.bat
```

Ou rode pelo terminal:

```bash
pip install -r requirements.txt
streamlit run app.py
```


## Seleção de arquivos antes da geração

Nos dois módulos, antes de processar, o usuário pode marcar exatamente quais saídas deseja gerar.

Exemplos:

- somente relatório completo;
- somente tabelas CSV para Protheus;
- relatório + log;
- pacote completo.

O ZIP final sempre inclui apenas os arquivos selecionados naquela execução.

## Módulo 1 — Precificação Comum

Use este módulo quando os analistas precisarem gerar as tabelas normais de precificação.

### Planilha de entrada

A planilha deve conter:

| Coluna | Campo |
|---|---|
| A | SKU |
| B | Custo Médio |
| C | Preço 001 |

O sistema também tenta reconhecer cabeçalhos como `SKU`, `Custo Médio` e `Preço 001`.

### Regras aplicadas

| Tabela | Regra |
|---|---|
| 001 | Preço 001 informado pelo usuário |
| 004 | Custo médio / 2 |
| 012 | Custo médio / 2 |
| 007 | Preço 001 com 16% de desconto |
| 013 | Preço da 007 + 12% |

### Arquivos gerados

- `protheus_tabela_001_*.csv`
- `protheus_tabela_004_*.csv`
- `protheus_tabela_012_*.csv`
- `protheus_tabela_007_*.csv`
- `protheus_tabela_013_*.csv`
- `relatorio_precificacao_comum_*.xlsx`
- `log_precificacao_comum_*.txt`
- `pacote_precificacao_comum_*.zip`

## Módulo 2 — Adequação ML

Use este módulo para a rotina já existente de adequação da tabela 007 com os preços revisados do Mercado Livre.

### Entradas

- Tabela 007 oficial
- `PREÇOS ML - REVISADO.xlsx`, com:
  - A = SKU
  - B = Preço ML
  - C = Custo Médio
- Relatório anterior opcional para comparativo

### Regras aplicadas

- Se o preço comercial 007 estiver maior que o preço ML, o sistema calcula `Preço ML x 0,90`.
- Se o preço comercial 007 já estiver menor ou igual ao ML, mantém o preço 007.
- Se o preço comercial 007 estiver como `0,01`, o sistema trata como item sem preço de venda cadastrado e calcula `Preço ML x 0,90`.
- A tabela 013 é gerada com `Preço 007 x 1,12`.
- O lucro bruto é calculado pela regra `Preço final da tabela 007 - Custo Médio`.
- O sistema gera alertas de prejuízo, markup baixo, margem baixa e custos inválidos.

## Formato dos arquivos Protheus

Todos os arquivos Protheus são gerados em CSV, sem cabeçalho, com separador `;` e seguem o mesmo padrão:

| Célula/Coluna | Regra |
|---|---|
| Linha 1 | Código da tabela com 3 dígitos e 4 campos no total, exemplo `007;;;` |
| Coluna A, a partir da linha 2 | SKUs normalizados com 5 dígitos, exemplo `00050` |
| Coluna B, a partir de B2 | Preço com ponto decimal, exemplo `25.99` |
| Coluna C, a partir de C2 | Data da alteração em `DD/MM/AAAA` |
| Delimitador final | Todas as linhas terminam com `;`, exemplo `00050;25.99;17/06/2026;` |
| Observação | O relatório analítico permanece em `.xlsx`; apenas os arquivos de subida Protheus são `.csv` |

## Logs

Cada módulo gera um `.txt` com intercorrências, incluindo:

- SKU vazio;
- SKU com mais de 5 dígitos;
- SKU em formato suspeito;
- SKU duplicado;
- preço inválido;
- custo inválido;
- linhas que não foram enviadas para alguma tabela por falta de dados válidos.

## Execução por terminal

### Adequação ML

```bash
python run_cli.py ml --tabela-007 "tabela 007-0526.xlsx" --ml "PREÇOS ML - REVISADO.xlsx" --saida saida_ml
```

### Precificação Comum

```bash
python run_cli.py comum --base "precificacao_comum.xlsx" --saida saida_comum
```


## Correção V11 - CSV Protheus

O CSV de subida agora é gerado obrigatoriamente no layout:

```csv
007;;;
36556;62.00;17/06/2026;
```

Regras reforçadas:
- código da tabela com 3 dígitos no conteúdo e no nome do arquivo;
- todas as linhas com `;` final;
- 4 campos por linha para evitar erro Protheus `array out of bounds (4 of 3)`.

## Feature V12 - seleção de SKUs antes da exportação

Depois de gerar o processamento, o app exibe a etapa **Revisar SKUs para exportação**.

Nessa etapa:

- todos os SKUs vêm marcados por padrão;
- o usuário pode desmarcar os SKUs que não devem ir para os arquivos Protheus;
- os CSVs e o ZIP são montados somente com os SKUs marcados;
- a análise continua visível para auditoria;
- o relatório Excel passa a incluir a coluna `Exportar?`, indicando o que entrou ou não na exportação daquela execução.

Essa função evita que o usuário edite o CSV manualmente no Excel, preservando o layout exigido pelo Protheus, especialmente:

```csv
007;;;
36556;62.00;17/06/2026;
```

## Feature V13 - nome personalizado nos downloads

Depois da etapa de revisão de SKUs, o app exibe o campo **Identificação dos arquivos (opcional)**.

Quando preenchido, esse texto passa a compor o nome dos arquivos baixados, junto com o código da tabela e a data em `DDMMAAAA`.

Exemplo com identificação `FEBI` e data `17/06/2026`:

- `001 - FEBI 17062026.csv`
- `007 - FEBI 17062026.csv`
- `013 - FEBI 17062026.csv`
- `relatorio_precificacao_comum - FEBI 17062026.xlsx`
- `log_precificacao_comum - FEBI 17062026.txt`
- `pacote_precificacao_comum - FEBI 17062026.zip`

Se o campo ficar vazio, o app usa o padrão técnico anterior.
