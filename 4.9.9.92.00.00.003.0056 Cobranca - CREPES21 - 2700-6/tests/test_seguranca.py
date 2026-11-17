import pytest

from conciliacao.erros import ErroConciliacao
from conciliacao.seguranca import parametros_no_sql, redigir, validar_sql


def test_select_e_with_passam():
    validar_sql("SELECT A FROM T WHERE D BETWEEN :data_ini AND :data_fim", "t.sql")
    validar_sql("WITH X AS (SELECT 1 AS A) SELECT A FROM X", "t.sql")


def test_comentario_nao_esconde_outro_comando():
    validar_sql("SELECT 1 FROM T /* DELETE FROM T */", "t.sql")
    with pytest.raises(ErroConciliacao, match="DELETE"):
        validar_sql("SELECT 1 FROM T; DELETE FROM T", "t.sql")


def test_dois_select_sao_recusados():
    with pytest.raises(ErroConciliacao, match="mais de um comando"):
        validar_sql("SELECT 1 FROM T; SELECT 2 FROM T", "t.sql")


def test_ponto_e_virgula_dentro_de_texto_nao_quebra():
    validar_sql("SELECT 1 FROM T WHERE NOME = 'a; DELETE'", "t.sql")


def test_select_into_e_recusado():
    with pytest.raises(ErroConciliacao, match="INTO"):
        validar_sql("SELECT * INTO COPIA FROM T", "t.sql")


def test_script_sem_consulta_definitiva_e_recusado():
    with pytest.raises(ErroConciliacao, match="PREENCHER_ORIGEM"):
        validar_sql("SELECT 1 AS A WHERE PREENCHER_ORIGEM = 1", "sinqia.sql")


def test_parametros_ignoram_cast_do_postgres():
    nomes = parametros_no_sql("SELECT CAST(:conta AS varchar) , campo::int FROM T WHERE D = :data_ini")
    assert nomes == {"conta", "data_ini"}


def test_redige_senha_na_mensagem():
    texto = redigir("oracle+oracledb://usuario:segredo@host:1521/ e PWD=abc123;UID=u")
    assert "segredo" not in texto
    assert "abc123" not in texto
    assert "***" in texto
