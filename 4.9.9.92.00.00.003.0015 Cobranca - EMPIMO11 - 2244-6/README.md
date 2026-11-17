# Conciliação

O motor lê consultas SQL de uma pasta só e monta o Excel de cada conciliação. A consulta não fica presa a uma conta: a conciliação é que escolhe quais scripts usar.

```
scripts/      consultas, uma por arquivo
pipelines/    uma conciliação por arquivo — a lista do que rodar
saida/        Excel gerado (não entra no git)
.env          usuário e senha (não entra no git)
```

Ted/Doc, por exemplo, usa a razão do Protheus (financeiro), a razão da Sinqia (contábil), o balancete da Sinqia (saldo anterior) e o movimento de conta corrente. Uma conciliação de PIX repete esses scripts. Passo a passo para incluir um script novo: `docs/como-adicionar-script.md`.

## Preparar

```
uv sync
copy .env.example .env
```

Preencha o `.env`. O usuário do banco precisa ter apenas SELECT. O motor também recusa script que não seja um único SELECT, e os filtros entram como parâmetro (`:data_ini`, `:data_fim`, `:conta`), nunca colados no texto.

O `protheus_ct2_razao.sql` lê a CT2010: toda linha em que a conta está no débito ou no crédito. O valor somado é o `VALOR_SINAL`, negativo quando a conta está no crédito. O `sinqia_ct_razao.sql` lê CT005/CT001. Se o resultado não bater com a consulta de referência, troque só o SELECT daquele arquivo.

O `sinqia_cc_razao.sql` ainda não tem a consulta da WP-04. Ele está marcado como `opcional: true` no yaml: a conciliação roda sem ele e a aba WP-04 não entra no Excel. Quando colar o SELECT, apague a linha `PREENCHER_ORIGEM`.

## Rodar

```
uv run conciliacao listar
uv run conciliacao rodar teddoc_outras_instituicoes --inicio 01/09/2026 --fim 30/09/2026 --conta 11030101051 --conta-sinqia 1886599000020001
uv run conciliacao rodar-todas --inicio 01/09/2026 --fim 30/09/2026
```

`--conta` é a conta do Protheus (CT2). `--conta-sinqia` é a da Sinqia. Sem esses argumentos, o programa pergunta no terminal (Enter usa o padrão do YAML). `rodar.bat` faz a mesma coisa, com dois cliques, depois do `uv sync`.

O Excel sai em `saida/` **da pasta do `rodar.bat` que você clicou**, no nome dos outros meses: conta, nome e ano-mês (`1.8.8.65.99.00.002.0001 - Transações TedDoc Outras Instituições - 09.2026.xlsx`). A pasta `saida` só é criada na hora de gravar o arquivo. Cópia no Ted/Doc grava no Ted/Doc; atalho ou cópia na Área de Trabalho grava lá. Cada script pronto ganha uma aba com os dados como vieram do banco. Fonte opcional ainda sem SELECT (linha `PREENCHER_ORIGEM`) não vira aba. A última conta é a conciliação por dia: débito positivo e crédito negativo nos dois lados; a diferença do dia é a soma. Dia dentro de R$ 0,01 fica conciliado. A aba Composição lista o detalhe só dos dias que passaram da tolerância, quando o script de detalhe existir.

## Acrescentar uma conciliação

Crie `pipelines/pix_outras.yaml`. Conta, filial e tolerância ficam entre aspas.

```yaml
id: pix_outras
nome: Conciliação - Transacoes Pix Outras Instituicoes
empresa: Banco Senff
parametros:
  filial: "030101"
  instituicao: "1414"
  agencia: "19"
  plano: "3"
fontes:
  - script: protheus_ct2_razao
    aba: WP-02 Protheus
    parametros:
      filial: "${filial}"
      conta: "11030101099"
  - script: sinqia_ct_razao
    aba: WP-03 Sinqia CT
    parametros:
      instituicao: "${instituicao}"
      agencia: "${agencia}"
      plano: "${plano}"
      conta: "1886599000020002"
  - script: sinqia_pix
    aba: WP-04 Pix
    parametros:
      conta: "1886599000020002"
conciliacao:
  esquerda: protheus_ct2_razao
  direita: sinqia_ct_razao
  detalhe: sinqia_pix
  operacao: soma
  tolerancia: 0.01
```

O id do script é o nome do arquivo, sem `.sql`. No topo de cada SQL:

```
-- @id: sinqia_pix
-- @engine: sqlserver
-- @descricao: Movimento Pix na Sinqia
-- @param: data_ini tipo=inteiro
-- @param: data_fim tipo=inteiro
-- @param: conta tipo=texto
-- @coluna.data: DATA
-- @coluna.valor: VALOR
```

`@coluna.data` e `@coluna.valor` dizem qual campo entra na conta. O resto da consulta aparece na aba como prova. `tipo=inteiro` manda a data como AAAAMMDD numérico (Sinqia). `tipo=texto` manda AAAAMMDD entre aspas (Protheus). `tipo=data` manda data de calendário.

Para o mesmo SELECT em duas contas, repita o script com `como:` e parâmetros diferentes. Cada `como` vira uma aba.

## O que o motor não faz

Não grava senha no Excel nem no log. Se um script passar de 200 mil linhas, a execução para: o WHERE do período ficou largo demais.
