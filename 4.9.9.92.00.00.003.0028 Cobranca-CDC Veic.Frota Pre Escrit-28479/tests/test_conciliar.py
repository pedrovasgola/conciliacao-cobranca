from datetime import date, datetime, timezone

import pandas as pd
import pytest

from conciliacao.conciliar import conciliar, filtrar_composicao, normalizar_data


def test_data_aaaammdd_e_datetime():
    serie = pd.Series([20260901, "20260902", datetime(2026, 9, 3)])
    datas = normalizar_data(serie)
    assert list(datas.dt.date) == [date(2026, 9, 1), date(2026, 9, 2), date(2026, 9, 3)]


def test_ct2_yyyymmdd_alinha_com_date_da_sinqia():
    """CT2 CHAR YYYYMMDD e Sinqia date não podem virar dois índices diferentes no groupby."""
    protheus = pd.DataFrame(
        {"CT2_DATA": ["20260901", 20260902, datetime(2026, 9, 3, 0, 0)], "VALOR": [-10.0, 5.0, 1.0]}
    )
    sinqia = pd.DataFrame(
        {"DATA": [date(2026, 9, 1), date(2026, 9, 2), "2026-09-03"], "VALOR": [10.0, -5.0, -1.0]}
    )
    resultado = conciliar(
        protheus, sinqia, "CT2_DATA", "VALOR", "DATA", "VALOR",
        "Protheus", "Sinqia", "soma", 0.01,
    )
    assert len(resultado.por_dia) == 3
    assert resultado.dias_ok == 3
    assert list(resultado.por_dia["DATA"]) == [date(2026, 9, 1), date(2026, 9, 2), date(2026, 9, 3)]


def test_serial_excel_e_datetime_com_hora_viram_o_mesmo_dia():
    serial = (date(2026, 9, 1) - date(1899, 12, 30)).days
    protheus = pd.DataFrame(
        {
            "DATA": [serial, datetime(2026, 9, 2, 15, 30, tzinfo=timezone.utc)],
            "VALOR": [-10.0, 4.0],
        }
    )
    sinqia = pd.DataFrame({"DATA": [date(2026, 9, 1), "2026-09-02"], "VALOR": [10.0, -4.0]})
    resultado = conciliar(
        protheus, sinqia, "DATA", "VALOR", "DATA", "VALOR",
        "Protheus", "Sinqia", "soma", 0.01,
    )
    assert resultado.dias_ok == 2
    assert len(resultado.por_dia) == 2


def test_soma_dos_lados_marca_o_dia():
    protheus = pd.DataFrame(
        {"DATA": [datetime(2026, 9, 1), datetime(2026, 9, 2)], "VALOR_SINAL": [-100.0, 40.0]}
    )
    sinqia = pd.DataFrame(
        {"DATA": [datetime(2026, 9, 1), datetime(2026, 9, 2)], "VALOR": [100.0, -40.0]}
    )
    resultado = conciliar(
        protheus, sinqia, "DATA", "VALOR_SINAL", "DATA", "VALOR",
        "Protheus", "Sinqia", "soma", 0.01,
    )
    assert resultado.dias_ok == 2
    assert resultado.diferenca_total == 0


def test_dia_que_nao_fecha_e_dia_so_de_um_lado():
    protheus = pd.DataFrame(
        {"DATA": [datetime(2026, 9, 1), datetime(2026, 9, 2)], "VALOR": [-6261938.58, 10.0]}
    )
    sinqia = pd.DataFrame(
        {"DATA": [datetime(2026, 9, 1)], "VALOR": [6861938.58]}
    )
    resultado = conciliar(
        protheus, sinqia, "DATA", "VALOR", "DATA", "VALOR",
        "Protheus", "Sinqia", "soma", 0.01,
    )
    assert resultado.dias_diferenca == 2
    assert resultado.diferenca_total == pytest.approx(600010.0)
    por_dia = resultado.por_dia.set_index("DATA")
    assert por_dia.loc[date(2026, 9, 1), "STATUS"] == "DIFERENÇA"
    assert por_dia.loc[date(2026, 9, 2), "Sinqia"] == 0


def test_composicao_fica_so_com_o_dia_divergente():
    protheus = pd.DataFrame({"DATA": [datetime(2026, 9, 1), datetime(2026, 9, 2)], "VALOR": [10.0, 5.0]})
    sinqia = pd.DataFrame({"DATA": [datetime(2026, 9, 1), datetime(2026, 9, 2)], "VALOR": [-10.0, 0.0]})
    resultado = conciliar(
        protheus, sinqia, "DATA", "VALOR", "DATA", "VALOR",
        "Protheus", "Sinqia", "soma", 0.01,
    )
    detalhe = pd.DataFrame(
        {
            "DATA": [datetime(2026, 9, 1), datetime(2026, 9, 2), datetime(2026, 9, 2)],
            "VALOR": [1.0, 9.0, 1.0],
        }
    )
    composicao = filtrar_composicao(detalhe, "DATA", resultado)
    assert list(pd.to_datetime(composicao["DATA"]).dt.day) == [2, 2]


def test_coluna_ausente_explica_o_que_veio():
    with pytest.raises(Exception, match="VALOR"):
        conciliar(
            pd.DataFrame({"DATA": [datetime(2026, 9, 1)]}),
            pd.DataFrame({"DATA": [datetime(2026, 9, 1)], "VALOR": [1.0]}),
            "DATA", "VALOR", "DATA", "VALOR",
            "Protheus", "Sinqia", "soma", 0.01,
        )
