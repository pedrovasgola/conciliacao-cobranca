-- @id: protheus_se5_cobranca
-- @engine: oracle
-- @descricao: Financeiro do Protheus (SE5) das naturezas da cobrança
-- @param: data_ini tipo=texto tamanho=8
-- @param: data_fim tipo=texto tamanho=8
-- @param: filial tipo=texto tamanho=6
-- @param: natureza tipo=texto tamanho=10
-- @param: natureza_2 tipo=texto tamanho=10
-- @coluna.data: DATA
-- @coluna.valor: VALOR

-- Posição financeira da cobrança. As duas naturezas cobrem a conta que
-- recebe por mais de uma natureza; nas demais, as duas são iguais.
-- Baixa ENF fica de fora, como na consulta que já fecha o mês.
-- O título entra se a emissão ou a baixa cai no período.
-- O dia da linha continua a data da baixa.

SELECT
    S.E5_FILIAL AS FILIAL,
    TRUNC(TO_DATE(NULLIF(TRIM(S.E5_DATA), ''), 'YYYYMMDD')) AS DATA,
    TRUNC(TO_DATE(NULLIF(TRIM(S1.E1_EMISSAO), ''), 'YYYYMMDD')) AS DATA_EMISSAO,
    S.E5_VALOR AS VALOR,
    S.E5_NATUREZ AS NATUREZA,
    S.E5_PREFIXO AS PREFIXO,
    S.E5_NUMERO AS NUMERO,
    S.E5_PARCELA AS PARCELA,
    S.E5_TIPO AS TIPO,
    S.E5_CLIFOR AS CLIENTE,
    S1.E1_NOMCLI AS NOME,
    S.E5_MOTBX AS MOTIVO_BAIXA,
    S.E5_DOCUMEN AS DOCUMENTO,
    S.E5_VENCTO AS VENCIMENTO,
    S1.E1_HIST AS HISTORICO
FROM SE5010 S
LEFT JOIN SE1010 S1
    ON S1.E1_FILIAL = S.E5_FILIAL
   AND S1.E1_PREFIXO = S.E5_PREFIXO
   AND S1.E1_NUM = S.E5_NUMERO
   AND S1.E1_PARCELA = S.E5_PARCELA
   AND S1.E1_TIPO = S.E5_TIPO
   AND S1.E1_CLIENTE = S.E5_CLIFOR
   AND S1.E1_LOJA = S.E5_LOJA
   AND S1.D_E_L_E_T_ = ' '
WHERE S.E5_FILIAL = :filial
  AND S.E5_NATUREZ IN (:natureza, :natureza_2)
  AND (
        S.E5_DATA BETWEEN :data_ini AND :data_fim
        OR S1.E1_EMISSAO BETWEEN :data_ini AND :data_fim
      )
  AND S.D_E_L_E_T_ = ' '
  AND S.E5_MOTBX <> 'ENF'
ORDER BY S.E5_DATA, S.E5_NUMERO
