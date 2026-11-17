"""Capa da cobrança compara a soma do período, não o dia."""

from datetime import date
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

from conciliacao.catalogo import carregar_pipelines, carregar_scripts
from conciliacao.cobranca import gravar

RAIZ = Path(__file__).resolve().parents[1]


def _pipeline():
    scripts = carregar_scripts(RAIZ / "scripts")
    return carregar_pipelines(scripts, RAIZ / "pipelines")["cobranca"], scripts


def test_sql_tira_renegociacao_e_mantem_pagamento_normal():
    scripts = carregar_scripts(RAIZ / "scripts")
    sql = scripts["em_baixas"].sql
    assert "EM07_ID_TIPHIS <> 2" in sql
    assert "EM07_ID_TPMOVI <> 'R'" in sql
    assert "EM07_DT_PAGAME BETWEEN" in sql
    assert scripts["protheus_se5_cobranca"].parametros["natureza"].tamanho == 10
    assert scripts["protheus_se5_cobranca"].parametros["filial"].tamanho == 6


def test_capa_soma_o_periodo(tmp_path: Path):
    pipeline, scripts = _pipeline()
    baixas = pd.DataFrame({"DATA": [date(2026, 9, 4), date(2026, 9, 15)], "VALOR": [1200.86, 2851.59]})
    financeiro = pd.DataFrame({"DATA": [date(2026, 9, 8), date(2026, 9, 16)], "VALOR": [1200.86, 2851.59]})
    destino = tmp_path / "capa.xlsx"
    gravar(
        destino,
        pipeline,
        scripts,
        {"em_baixas": baixas, "protheus_se5_cobranca": financeiro},
        date(2026, 9, 1),
        date(2026, 9, 30),
        "dev",
    )
    livro = load_workbook(destino)
    assert livro.sheetnames == ["WP-01 Saldos", "WP-02 Baixas", "WP-03 Financeiro"]
    capa = livro["WP-01 Saldos"]
    assert "Conciliação Contábil" in capa["B2"].value
    assert capa["C5"].value == "Banco Senff"
    assert capa["I5"].value.date() == date(2026, 9, 1)
    assert capa["I7"].value == "dev"
    assert capa["C17"].value == "4.9.9.92.00.00.003.0002"
    assert capa["F17"].value == "=SUM(tbl_em_baixas[VALOR])"
    assert capa["H18"].value == "=SUM(tbl_protheus_se5_cobranca[VALOR])"
    assert capa["I17"].value == "=ROUND(H18-F17,2)"
    assert capa.sheet_view.view == "normal"
    assert abs(livro["WP-02 Baixas"].cell(12, 3).value - 1200.86) < 0.001
