-- @id: sinqia_ct_balancete
-- @engine: sqlserver
-- @descricao: Saldo anterior da conta na Sinqia
-- @param: data_ini tipo=inteiro
-- @param: data_fim tipo=inteiro
-- @param: conta tipo=texto
-- @param: instituicao tipo=inteiro
-- @param: agencia tipo=inteiro
-- @param: plano tipo=inteiro
-- @coluna.valor: SALDO_ATUAL

-- Saldo do balancete mensal (CT2004, tipo 20), o mesmo da tela.
-- A tela guarda crédito menos débito. Aqui o sinal fica débito menos crédito,
-- igual à razão. Saldo anterior é o do mês da data inicial.

WITH alvo AS (
    SELECT c.CT001_SQ_CONTA AS sq_conta
    FROM SFINANCE_CT.dbo.CT001 AS c
    WHERE c.CT001_NR_PLANO = :plano
      AND c.CT001_CT_CONTSM = :conta
),
meses AS (
    SELECT
        s.CT2004_DT_SALDO AS ano_mes,
        s.CT2004_VL_SDOANT AS saldo_anterior,
        s.CT2004_VL_DEB AS debito,
        s.CT2004_VL_CRED AS credito,
        s.CT2004_VL_SDO AS saldo_atual
    FROM SFINANCE_CT.dbo.CT2004 AS s WITH (NOLOCK)
    INNER JOIN alvo AS a
        ON a.sq_conta = s.CT2004_SQ_CONTA
    WHERE s.CT2004_NR_INST = :instituicao
      AND s.CT2004_NR_AGEN = :agencia
      AND s.CT2004_TP_SALDO = 20
      AND s.CT2004_DT_SALDO BETWEEN (:data_ini / 100) AND (:data_fim / 100)
),
resumo AS (
    SELECT
        (SELECT TOP (1) -saldo_anterior FROM meses ORDER BY ano_mes) AS saldo_anterior,
        SUM(debito) AS debito,
        SUM(credito) AS credito,
        (SELECT TOP (1) -saldo_atual FROM meses ORDER BY ano_mes DESC) AS saldo_atual
    FROM meses
)
SELECT
    c.CT001_CT_CONTSM AS CONTA,
    c.CT001_DS_NOMCT AS NOME,
    c.CT001_GR_CONT AS GRAU,
    CASE c.CT001_ID_ANLSIN
        WHEN 10 THEN 'sintetica'
        WHEN 0 THEN 'analitica'
        WHEN 20 THEN 'analitica'
        ELSE CONVERT(varchar(20), c.CT001_ID_ANLSIN)
    END AS TIPO,
    CAST(ISNULL(r.saldo_anterior, 0) AS decimal(24, 2)) AS SALDO_ANTERIOR,
    CAST(ISNULL(r.debito, 0) AS decimal(24, 2)) AS DEBITO,
    CAST(ISNULL(r.credito, 0) AS decimal(24, 2)) AS CREDITO,
    CAST(ISNULL(r.debito, 0) - ISNULL(r.credito, 0) AS decimal(24, 2)) AS MOVIMENTO,
    CAST(ISNULL(r.saldo_atual, 0) AS decimal(24, 2)) AS SALDO_ATUAL
FROM alvo
INNER JOIN SFINANCE_CT.dbo.CT001 AS c
    ON c.CT001_SQ_CONTA = alvo.sq_conta
   AND c.CT001_NR_PLANO = :plano
CROSS JOIN resumo AS r
