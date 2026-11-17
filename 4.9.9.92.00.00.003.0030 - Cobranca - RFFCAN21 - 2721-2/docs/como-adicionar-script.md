# Como acrescentar um script

Ted/Doc hoje usa dois SELECT prontos (CT2 e Sinqia CT) e um terceiro ainda vazio (conta corrente). O passo a passo abaixo é o mesmo para o quarto, o quinto, ou para uma conciliação nova que reusa estes e soma os dela.

Há dois caminhos. Escolha um antes de criar arquivo.

1. **A conciliação já existe** (ex.: Ted/Doc precisa da conta corrente). Você cria o `.sql` e **inclui o id na lista `fontes` do YAML**.
2. **É outra conciliação** (ex.: PIX). Você reusa os scripts que já servem, cria só os que faltam, e faz um YAML novo em `pipelines/`.

O motor não lê pasta de conciliação: ele lê a pasta `scripts/` inteira. Quem decide o que entra em cada rodada é o YAML.

---

## Passo 1 — Pegar o SELECT que já roda

Abra a WP no Excel → Dados → Consultas e conexões → a consulta da aba que você quer trazer.

Copie o SELECT inteiro. Confira se é um único `SELECT` (ou `WITH ... SELECT`). O motor recusa `INSERT`, `UPDATE`, `DELETE`, `DECLARE`, dois comandos no mesmo arquivo.

Anote o que está fixo no `WHERE`. Exemplo da Sinqia CT:

| No SELECT original | Vira parâmetro |
|---|---|
| `BETWEEN 20260901 AND 20260930` | `:data_ini` e `:data_fim` |
| conta (`CT001_CT_CONTSM = ...`) | `:conta` |
| instituição, agência, plano | `:instituicao`, `:agencia`, `:plano` |

Data e conta **não** ficam coladas no texto. Sempre `:nome`.

---

## Passo 2 — Criar o arquivo na pasta certa

```
scripts/
  oracle/        → Protheus
  sqlserver/     → Sinqia
```

Nome do arquivo = id do script, só minúscula, número e `_`. Sem espaço, sem acento.

Exemplos:

- `scripts/sqlserver/sinqia_cc_razao.sql` (já existe, ainda sem SELECT)
- `scripts/sqlserver/sinqia_pix.sql` (novo)

O `@id` no cabeçalho **tem que ser igual** ao nome do arquivo, sem `.sql`.

---

## Passo 3 — Cabeçalho no topo do SQL

Cole isto no começo, depois o SELECT. Troque só o que está marcado.

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

Regras do cabeçalho:

- `@engine`: `oracle`, `sqlserver` ou `postgres`.
- Um `@param` para **cada** `:nome` usado no SELECT.
- `tipo=texto` — Protheus (data `20260901` como texto). Coluna CHAR: acrescente `tamanho=N`, o DATA_LENGTH da coluna (`conta` 20, `filial` 6, `data` 8, `natureza` 10) e compare com `coluna = :param`. O motor completa o espaço. `RTRIM` na coluna acha a linha, mas lê a tabela inteira.
- `tipo=inteiro` — Sinqia (data `20260901` como número).
- `tipo=data` — data de calendário (Postgres).
- `@coluna.data` e `@coluna.valor` — nomes das colunas **depois** do SELECT, as que entram na soma por dia. O resto das colunas só aparece na aba, como prova.

Se o SELECT original não tiver uma coluna de valor com sinal, acrescente uma. Na Sinqia CT ficou `VALOR` = débito − crédito. Na CT2 ficou `VALOR_SINAL`.

---

## Passo 4 — Trocar os filtros fixos por `:parametro`

Errado (data e conta coladas):

```
WHERE M.CT005_DT_MOVTO BETWEEN 20260901 AND 20260930
  AND C.CT001_CT_CONTSM = 1886599000020001
```

Certo:

```
WHERE M.CT005_DT_MOVTO BETWEEN :data_ini AND :data_fim
  AND C.CT001_CT_CONTSM = :conta
```

`:data_ini` e `:data_fim` o motor preenche sozinho com o `--inicio` / `--fim`.  
`:conta` na CT2 vem de `--conta`. Na Sinqia, de `--conta-sinqia`.  
Os outros (`:filial`, `:instituicao`, `:agencia`, `:plano`) ficam no YAML, com aspas.

Não deixe a linha `PREENCHER_ORIGEM`. Se ela existir, o script é tratado como “ainda não está pronto”.

---

## Passo 5 — Registrar no YAML da conciliação

Abra o arquivo em `pipelines/`. Ted/Doc é `teddoc_outras_instituicoes.yaml`.

Na lista `fontes`, acrescente um item. O `script:` é o `@id`.

```yaml
  - script: sinqia_pix
    aba: WP-04 Pix
    resumo: Movimento Pix na Sinqia no período.
    parametros:
      instituicao: "${instituicao}"
      agencia: "${agencia}"
      plano: "${plano}"
      conta: "1886599000020001"
```

O que cada campo faz:

- `script` — id do arquivo em `scripts/`.
- `aba` — nome da aba no Excel (máximo 31 caracteres). Não use `Capa`, `Conciliação` nem `Composição`.
- `resumo` — texto no topo da aba daquele script.
- `parametros` — um valor para cada `@param` do SQL, **exceto** `data_ini` e `data_fim`.
- `"${instituicao}"` — reusa o valor do bloco `parametros:` no topo do YAML.
- aspas em conta, filial, instituição. Sem aspas o YAML come o zero à esquerda (`030101` vira outro número).

Papel do script na conta:

```yaml
conciliacao:
  esquerda: protheus_ct2_razao    # um lado da diferença
  direita: sinqia_ct_razao        # o outro lado
  detalhe: sinqia_cc_razao        # só explica os dias que não fecharam
```

- Entra na **diferença** → `esquerda` ou `direita`.
- Só **explica** dia que não fechou → `detalhe`.
- Só **aparece** no Excel, não entra na conta → não coloque em nenhum dos três.

Se o SELECT ainda não está colado, marque `opcional: true`. A conciliação roda sem ele. Quando colar o SELECT, apague o `PREENCHER_ORIGEM` (pode deixar o `opcional: true`; sem aquela linha o script passa a rodar).

---

## Passo 6 — Conferir sem abrir o banco

Na pasta do projeto:

```
uv run conciliacao listar
```

O script novo tem que aparecer como `pronto`. Se aparecer `falta colar o SELECT` ou `bloqueado`, o cabeçalho ou o SELECT ainda não passou na regra (SELECT único, parâmetros declarados, sem `PREENCHER_ORIGEM`).

A conciliação tem que listar o id novo em `scripts: ...`.

---

## Passo 7 — Rodar

```
uv run conciliacao rodar teddoc_outras_instituicoes --inicio 01/09/2026 --fim 30/09/2026 --conta 11030101051 --conta-sinqia 1886599000020001
```

O Excel sai em `saida/` da pasta do `rodar.bat` que você clicou (a pasta só é criada na hora de gravar). Cada script pronto vira uma aba. Fonte opcional ainda sem SELECT não aparece.

---

## Exemplo: terceiro script no Ted/Doc (conta corrente)

O arquivo `scripts/sqlserver/sinqia_cc_razao.sql` já está no YAML como `detalhe`, com `opcional: true`. Falta só o SELECT.

1. Copiar o SELECT da WP-04.
2. Colar no arquivo, no lugar do `SELECT ... WHERE PREENCHER_ORIGEM = 1`.
3. Trocar data e conta por `:data_ini`, `:data_fim`, `:conta`.
4. Garantir colunas `DATA` e `VALOR` (ou ajustar `@coluna.data` / `@coluna.valor`).
5. Apagar qualquer `PREENCHER_ORIGEM`.
6. `uv run conciliacao listar` → `sinqia_cc_razao` como `pronto`.
7. Rodar o mês. A aba WP-04 só aparece depois que o SELECT estiver colado. A aba Composição lista o CC nos dias em que CT2 e Sinqia CT não fecharam.

Não precisa criar outro YAML. O id já está em `fontes`.

---

## Exemplo: outra conciliação reusando os dois e somando um

PIX usa CT2 + Sinqia CT (já existem) e um SELECT só dela.

1. Criar `scripts/sqlserver/sinqia_pix.sql` (passos 1 a 4).
2. Copiar `pipelines/teddoc_outras_instituicoes.yaml` para `pipelines/pix_outras.yaml`.
3. Trocar `id`, `nome`, contas no YAML.
4. Em `fontes`, manter `protheus_ct2_razao` e `sinqia_ct_razao`, acrescentar `sinqia_pix`.
5. Ajustar `esquerda` / `direita` / `detalhe`.
6. `uv run conciliacao listar` e depois `uv run conciliacao rodar pix_outras ...`.

Os dois primeiros arquivos SQL **não se copiam**. Os dois YAML apontam para o mesmo `@id`.

---

## Se der erro

| Mensagem | O que conferir |
|---|---|
| `@id` diferente do arquivo | Nome do `.sql` = linha `-- @id:` |
| `declare @param` | Todo `:nome` do SELECT tem `-- @param:` |
| `coloque o valor entre aspas` | Conta, filial, instituição no YAML entre aspas |
| `ainda não tem a consulta definitiva` | Apagar `PREENCHER_ORIGEM` |
| `foi recusado: encontrei 'DELETE'` | Um SELECT só; nada de outro comando |
| `aba reservada` | Não chamar a aba de Capa / Conciliação / Composição |
| `não achei a coluna DATA` | `@coluna.data` tem que ser o alias que o SELECT devolve |

Credencial continua no `.env`. Script novo no mesmo banco (Oracle ou SQL Server) não pede variável extra. Banco novo (Postgres) precisa das linhas `PG_*` no `.env`.
