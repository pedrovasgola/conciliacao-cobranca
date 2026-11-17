-- @id: sinqia_ct_razao
-- @engine: sqlserver
-- @descricao: Razão contábil Sinqia da conta transitória
-- @param: data_ini tipo=inteiro
-- @param: data_fim tipo=inteiro
-- @param: conta tipo=texto
-- @param: instituicao tipo=inteiro
-- @param: agencia tipo=inteiro
-- @param: plano tipo=inteiro
-- @coluna.data: DATA
-- @coluna.valor: VALOR

-- Consulta da WP-03. Na hora de rodar: período e --conta-sinqia.
-- Instituição, agência e plano ficam no YAML (padrão Senff: 1414, 19, 3).
-- VALOR = débito − crédito, para somar com o VALOR_SINAL da CT2.

SELECT
    C.CT001_CT_CONTSM AS CT001_CT_CONT,
    C.CT001_DS_NOMCT AS CT001_DS_NOMCT,
    M.CT005_DT_MOVTO AS CT005_DT_MOVTO,
    CP.CT001_CT_CONTSM AS CT005_CT_CONTCP,
    M.CT005_NR_LOTE AS CT005_NR_LOTE,
    M.CT005_DS_HIST AS CT005_DS_HIST,
    0.00 AS CT005_VL_SDOANT,
    CASE WHEN LTRIM(RTRIM(M.CT005_TP_MOVTO)) = 'D' THEN M.CT005_VL_LANC ELSE 0 END AS CT005_VL_DEB,
    CASE WHEN LTRIM(RTRIM(M.CT005_TP_MOVTO)) = 'C' THEN M.CT005_VL_LANC ELSE 0 END AS CT005_VL_CRED,
    (
        CASE WHEN LTRIM(RTRIM(M.CT005_TP_MOVTO)) = 'D' THEN M.CT005_VL_LANC ELSE 0 END
        - CASE WHEN LTRIM(RTRIM(M.CT005_TP_MOVTO)) = 'C' THEN M.CT005_VL_LANC ELSE 0 END
    ) AS VALOR,
    DATEFROMPARTS(
        M.CT005_DT_MOVTO / 10000,
        (M.CT005_DT_MOVTO / 100) % 100,
        M.CT005_DT_MOVTO % 100
    ) AS DATA,
    0.00 AS CT005_VL_SDOATU,
    U.US01_NM_USR AS CT005_NM_OPESIS
FROM SFINANCE_CT.dbo.CT005 M (NOLOCK)
INNER JOIN SFINANCE_CT.dbo.CT001 C (NOLOCK)
    ON M.CT005_SQ_CONTA = C.CT001_SQ_CONTA
LEFT JOIN SFINANCE_CT.dbo.CT001 CP (NOLOCK)
    ON M.CT005_SQ_CONTCP = CP.CT001_SQ_CONTA
LEFT JOIN SFINANCE_CT.dbo.US01 U (NOLOCK)
    ON M.CT005_CD_OPESIS = U.US01_NR_USR
WHERE M.CT005_NR_INST = :instituicao
  AND M.CT005_NR_AGEN = :agencia
  AND M.CT005_DT_MOVTO BETWEEN :data_ini AND :data_fim
  AND C.CT001_NR_PLANO = :plano
  AND C.CT001_CT_CONTSM = :conta
ORDER BY
    C.CT001_CT_CONTSM,
    M.CT005_DT_MOVTO,
    M.CT005_NR_LOTE
