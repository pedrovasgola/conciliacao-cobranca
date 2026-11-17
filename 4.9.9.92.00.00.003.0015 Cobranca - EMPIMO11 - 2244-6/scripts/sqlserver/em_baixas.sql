-- @id: em_baixas
-- @engine: sqlserver
-- @descricao: Movimento de baixas do Sinqia EM no produto da conta
-- @param: data_ini tipo=inteiro
-- @param: data_fim tipo=inteiro
-- @param: instituicao tipo=inteiro
-- @param: agencia tipo=inteiro
-- @param: produto tipo=inteiro
-- @param: produto_2 tipo=inteiro
-- @coluna.data: DATA_MOVIMENTO
-- @coluna.valor: VALOR

-- O período filtra pela data de movimento. A data de pagamento fica na planilha.
-- Fica de fora só a renegociação (tipo R). Histórico 2 entra.

SELECT
    b.EM07_NR_AGEN AS AGENCIA,
    CASE
        WHEN LEN(cod.PRODUTO) > 1 THEN LEFT(cod.PRODUTO, LEN(cod.PRODUTO) - 1) + '-' + RIGHT(cod.PRODUTO, 1)
        ELSE cod.PRODUTO
    END AS PRODUTO,
    LTRIM(RTRIM(p.EM02_DS_PRODUR)) AS SIGLA,
    LTRIM(RTRIM(p.EM02_DS_PRODUC)) AS DESCRICAO,
    CASE
        WHEN LEN(cod.CLIENTE) > 1 THEN LEFT(cod.CLIENTE, LEN(cod.CLIENTE) - 1) + '-' + RIGHT(cod.CLIENTE, 1)
        ELSE cod.CLIENTE
    END AS CLIENTE,
    LTRIM(RTRIM(cli.CL01_NM_CLIENT)) AS NOME,
    CASE
        WHEN LEN(cod.CONTRATO) = 11 THEN LEFT(cod.CONTRATO, 7) + '-' + SUBSTRING(cod.CONTRATO, 8, 3) + '-' + RIGHT(cod.CONTRATO, 1)
        ELSE cod.CONTRATO
    END AS CONTRATO,
    b.EM07_NR_PRESTA AS PRESTACAO,
    CASE
        WHEN b.EM07_DT_MOVIME BETWEEN 19000101 AND 29991231
        THEN CONVERT(date, CONVERT(varchar(8), b.EM07_DT_MOVIME), 112)
    END AS DATA_MOVIMENTO,
    CONVERT(date, CONVERT(varchar(8), b.EM07_DT_PAGAME), 112) AS DATA_PAGAMENTO,
    b.EM07_VL_PAGAME AS VALOR,
    b.EM07_VL_PRESTA AS VALOR_PRESTACAO,
    b.EM07_ID_TPMOVI AS TIPO_MOVIMENTO,
    b.EM07_ID_ESTORN AS ESTORNO,
    b.EM07_ID_TIPHIS AS TIPO_HISTORICO
FROM SFINANCEEM.dbo.EM07 b
INNER JOIN SFINANCEEM.dbo.EM01 c
    ON c.EM01_NR_INST = b.EM07_NR_INST
   AND c.EM01_NR_AGEN = b.EM07_NR_AGEN
   AND c.EM01_CD_CLIENT = b.EM07_CD_CLIENT
   AND c.EM01_NR_CONTRA = b.EM07_NR_CONTRA
INNER JOIN SFINANCEEM.dbo.EM02 p
    ON p.EM02_NR_INST = c.EM01_NR_INST
   AND p.EM02_CD_PRODUT = c.EM01_CD_PRODUT
LEFT JOIN SFINANCE_CD.dbo.CL01 cli
    ON cli.CL01_NR_INST = c.EM01_NR_INST
   AND cli.CL01_NR_CLIENT = c.EM01_CD_CLIENT
CROSS APPLY (
    SELECT
        CONVERT(varchar(20), CONVERT(bigint, p.EM02_CD_PRODUT)) AS PRODUTO,
        CONVERT(varchar(20), CONVERT(bigint, c.EM01_CD_CLIENT)) AS CLIENTE,
        CONVERT(varchar(20), CONVERT(bigint, c.EM01_NR_CONTRA)) AS CONTRATO
) cod
WHERE b.EM07_NR_INST = :instituicao
  AND b.EM07_NR_AGEN = :agencia
  AND c.EM01_CD_PRODUT IN (:produto, :produto_2)
  AND b.EM07_DT_MOVIME BETWEEN :data_ini AND :data_fim
  AND b.EM07_VL_PAGAME <> 0
  AND b.EM07_ID_TPMOVI <> 'R'
ORDER BY b.EM07_DT_MOVIME, b.EM07_DT_PAGAME, c.EM01_NR_CONTRA, b.EM07_NR_PRESTA
