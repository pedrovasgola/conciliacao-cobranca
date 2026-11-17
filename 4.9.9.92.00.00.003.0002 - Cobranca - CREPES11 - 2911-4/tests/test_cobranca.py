"""Cobrança: diferença é baixa do EM menos a SE5. O CT só mostra o saldo anterior."""

from datetime import date
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

from conciliacao.catalogo import carregar_pipelines, carregar_scripts
from conciliacao.conciliar import conciliar
from conciliacao.planilha import gravar

RAIZ = Path(__file__).resolve().parents[1]


def _pipeline():
    scripts = carregar_scripts(RAIZ / "scripts")
    return carregar_pipelines(scripts, RAIZ / "pipelines")["cobranca"], scripts


def test_sql_tira_renegociacao_e_mantem_pagamento_normal():
    scripts = carregar_scripts(RAIZ / "scripts")
    sql = scripts["em_baixas"].sql
    assert "EM07_ID_TIPHIS <> 2" not in sql
    assert "EM07_ID_TPMOVI <> 'R'" in sql
    assert "EM07_DT_MOVIME BETWEEN" in sql
    assert "DATA_PAGAMENTO" in sql
    assert "EM07_DT_PAGAME BETWEEN" not in sql
    assert "CT2004" in scripts["sinqia_ct_balancete"].sql
    assert "CT2004_TP_SALDO = 20" in scripts["sinqia_ct_balancete"].sql
    assert "SALDO_ANTERIOR" in scripts["sinqia_ct_balancete"].sql
    assert scripts["sinqia_ct_razao"].coluna("data") == "DATA"
    assert scripts["sinqia_ct_razao"].coluna("valor") == "VALOR"
    se5 = scripts["protheus_se5_cobranca"].sql
    assert "E1_EMISSAO" in se5
    assert "DATA_EMISSAO" in se5
    assert "S.E5_DATA BETWEEN :data_ini AND :data_fim" in se5
    assert "S1.E1_EMISSAO BETWEEN :data_ini AND :data_fim" in se5
    assert "E5_MOTBX <> 'ENF'" in se5


def test_capa_traz_saldo_anterior_e_conciliacao_por_dia(tmp_path: Path):
    pipeline, scripts = _pipeline()
    assert pipeline.esquerda == "em_baixas"
    assert pipeline.direita == "protheus_se5_cobranca"
    assert pipeline.detalhe == "em_baixas"
    assert pipeline.operacao == "subtracao"
    ct = pd.DataFrame({"DATA": [date(2026, 10, 8)], "VALOR": [-100.0]})
    se5 = pd.DataFrame({"DATA": [date(2026, 10, 8)], "VALOR": [100.0]})
    em = pd.DataFrame({
        "DATA_MOVIMENTO": [date(2026, 10, 4)],
        "DATA_PAGAMENTO": [date(2026, 10, 2)],
        "VALOR": [40.0],
    })
    balancete = pd.DataFrame({"SALDO_ANTERIOR": [-50.0], "SALDO_ATUAL": [-150.0]})
    resultado = conciliar(
        em, se5, "DATA_MOVIMENTO", "VALOR", "DATA", "VALOR",
        "WP-04 Baixas", "WP-02 Financeiro", pipeline.operacao, 0.01,
    )
    destino = tmp_path / "capa.xlsx"
    gravar(
        destino,
        pipeline,
        scripts,
        {
            "sinqia_ct_razao": ct,
            "protheus_se5_cobranca": se5,
            "em_baixas": em,
            "sinqia_ct_balancete": balancete,
        },
        resultado,
        None,
        date(2026, 10, 1),
        date(2026, 10, 31),
        "dev",
    )
    livro = load_workbook(destino)
    assert "Conciliação" in livro.sheetnames
    assert "WP-05 Balancete" in livro.sheetnames
    assert "WP-04 Baixas" in livro.sheetnames
    capa = livro["WP-01 Saldos"]
    assert capa["E17"].value == 0
    assert capa["C9"].value == -50
    assert capa["C17"].value == "29114"
    assert capa["C18"].value == "1010066"
    assert capa["I17"].value == "=ROUND(F17-H18,2)"
    assert "Conciliação" in str(capa["B21"].value)
