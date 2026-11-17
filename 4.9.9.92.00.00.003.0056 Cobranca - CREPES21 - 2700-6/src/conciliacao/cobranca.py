"""Capa da cobrança: soma das baixas do EM contra a soma da SE5 no período."""

from __future__ import annotations

import os
from datetime import date, datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.worksheet import Worksheet

from conciliacao.catalogo import Pipeline, Script
from conciliacao.planilha import (
    ABA_SALDOS,
    _AMARELO,
    _AZUL,
    _CINZA,
    _CINZA_TEXTO,
    _FMT_DATA,
    _FMT_DATA_HORA,
    _FMT_VALOR,
    _FONTE,
    _INPUT,
    _NAVY,
    _VERDE_FUNDO,
    _VERDE_TEXTO,
    _VERMELHO_FUNDO,
    _VERMELHO_TEXTO,
    _aba_dados,
    _cabecalho,
    _faixa,
    _fonte,
    _fundo,
    _nome_tabela,
    _par,
    _q,
    _retrato,
    _secao,
    _status_cf,
)

_ULTIMA = 10
_FINO = Side(style="thin", color="BFBFBF")


def _maiusculas(quadro: pd.DataFrame | None) -> pd.DataFrame:
    if quadro is None:
        return pd.DataFrame()
    saida = quadro.copy()
    saida.columns = [str(coluna).upper() for coluna in saida.columns]
    return saida


def somar(quadro: pd.DataFrame | None) -> float:
    quadro = _maiusculas(quadro)
    if quadro.empty or "VALOR" not in quadro.columns:
        return 0.0
    return float(pd.to_numeric(quadro["VALOR"], errors="coerce").fillna(0).sum())


def _naturezas(pipeline: Pipeline) -> str:
    primeira = pipeline.parametros.get("natureza", "").strip()
    segunda = pipeline.parametros.get("natureza_2", "").strip()
    if segunda and segunda != primeira:
        return f"{primeira} / {segunda}"
    return primeira


def _produtos(pipeline: Pipeline) -> str:
    primeiro = pipeline.parametros.get("produto", "").strip()
    segundo = pipeline.parametros.get("produto_2", "").strip()
    if segundo and segundo != primeiro:
        return f"{primeiro} / {segundo}"
    return primeiro


def _linha_valor(ws: Worksheet, linha: int) -> None:
    ws.row_dimensions[linha].height = 20
    for coluna in range(2, _ULTIMA + 1):
        celula = ws.cell(linha, coluna)
        celula.alignment = Alignment(vertical="center", horizontal="left" if coluna <= 4 else "right")
        celula.border = Border(
            bottom=_FINO,
            left=_FINO if coluna == 2 else None,
            right=_FINO if coluna in (6, 8, _ULTIMA) else None,
        )
        if coluna >= 5:
            celula.number_format = _FMT_VALOR
            celula.font = _fonte(10, coluna >= 9)


def _caminho_gravavel(caminho: Path) -> Path:
    """O nome da pasta passa de 260 caracteres no UNC. O prefixo longo deixa o Excel gravar."""
    absoluto = os.path.abspath(caminho)
    if os.name != "nt" or len(absoluto) < 240:
        return caminho
    if absoluto.startswith("\\\\"):
        return Path("\\\\?\\UNC\\" + absoluto[2:])
    return Path("\\\\?\\" + absoluto)


def gravar(
    caminho: Path,
    pipeline: Pipeline,
    scripts: dict[str, Script],
    extratos: dict[str, pd.DataFrame],
    inicio: date,
    fim: date,
    feito_por: str,
) -> None:
    wb = Workbook()
    capa = wb.active
    capa.title = ABA_SALDOS
    for nome, celula in (("DataInicial", "$I$5"), ("DataFinal", "$I$6"), ("Tolerancia", "$C$7")):
        wb.defined_names[nome] = DefinedName(nome, attr_text=f"{_q(ABA_SALDOS)}!{celula}")

    _faixa(capa, "Conciliação Contábil  ·  WP-01 Saldos", pipeline.nome, _ULTIMA)
    larguras = {"B": 16, "C": 28, "D": 42, "E": 16, "F": 16, "G": 16, "H": 16, "I": 16, "J": 16}
    for letra, largura in larguras.items():
        capa.column_dimensions[letra].width = largura

    _par(capa, 5, 2, "Empresa", pipeline.empresa, mesclar_ate=4)
    _par(capa, 6, 2, "Nome WP", pipeline.nome, mesclar_ate=4)
    tolerancia = _par(
        capa, 7, 2, "Tolerância", pipeline.tolerancia, _FMT_VALOR,
        fonte=_fonte(10, True, "0000FF"), fundo=_AMARELO,
    )
    tolerancia.alignment = Alignment(horizontal="left")
    capa.cell(7, 3).comment = Comment(
        "Valor do YAML. Pode alterar aqui: as fórmulas recalculam.", "Conciliação",
    )
    _par(
        capa, 8, 2, "Regra",
        "Diferença = financeiro (SE5) − baixas (EM), na soma do período",
        fonte=_fonte(10), mesclar_ate=6,
    )
    _par(capa, 5, 8, "Data inicial", inicio, _FMT_DATA, mesclar_ate=10)
    _par(capa, 6, 8, "Data final", fim, _FMT_DATA, mesclar_ate=10)
    _par(capa, 7, 8, "Feito por", feito_por or "", mesclar_ate=10, fonte=_INPUT, fundo=_AMARELO)
    _par(
        capa, 8, 8, "Gerado em", datetime.now().replace(microsecond=0),
        _FMT_DATA_HORA, mesclar_ate=10,
    )
    _par(capa, 9, 8, "Revisado por", None, mesclar_ate=10, fonte=_INPUT, fundo=_AMARELO)

    em = pipeline.fonte(pipeline.esquerda)
    se5 = pipeline.fonte(pipeline.direita)
    tabela_em = _nome_tabela(em.alias)
    tabela_se5 = _nome_tabela(se5.alias)

    capa.row_dimensions[11].height = 18
    capa.row_dimensions[12].height = 28
    blocos = (
        ("B", "C", "Linhas EM", f"=COUNTA({tabela_em}[VALOR])", "0"),
        ("D", "E", "Linhas SE5", f"=COUNTA({tabela_se5}[VALOR])", "0"),
        ("F", "H", "Diferença do período", "=I17", _FMT_VALOR),
        ("I", "J", "Status", '=IF(ABS(I17)<=Tolerancia,"CONCILIADO","COM DIFERENÇA")', "@"),
    )
    for inicio_col, fim_col, rotulo, formula, formato in blocos:
        coluna_ini = capa[f"{inicio_col}1"].column
        coluna_fim = capa[f"{fim_col}1"].column
        for linha in (11, 12):
            for coluna in range(coluna_ini, coluna_fim + 1):
                capa.cell(linha, coluna).fill = _fundo(_CINZA)
        if inicio_col != fim_col:
            capa.merge_cells(f"{inicio_col}11:{fim_col}11")
            capa.merge_cells(f"{inicio_col}12:{fim_col}12")
        rotulo_celula = capa[f"{inicio_col}11"]
        rotulo_celula.value = rotulo
        rotulo_celula.font = _fonte(8, True, _CINZA_TEXTO)
        rotulo_celula.alignment = Alignment(horizontal="center", vertical="bottom")
        valor = capa[f"{inicio_col}12"]
        valor.value = formula
        valor.font = _fonte(16, True, _NAVY)
        valor.number_format = formato
        valor.alignment = Alignment(horizontal="center", vertical="center")
    _status_cf(capa, "I12:J12")

    _secao(capa, 14, "Saldos das contas", _ULTIMA)
    for faixa, texto in (
        ("B15:B16", "Origem"),
        ("C15:C16", "Conta"),
        ("D15:D16", "Descrição"),
        ("E15:F15", "BAIXAS (EM)"),
        ("G15:H15", "FINANCEIRO (SE5)"),
        ("I15:J16", "DIFERENÇA"),
    ):
        capa.merge_cells(faixa)
        capa[faixa.split(":")[0]] = texto
    capa.merge_cells("E16:F16")
    capa.merge_cells("G16:H16")
    capa["E16"] = "=DataFinal"
    capa["G16"] = "=DataFinal"
    for linha in (15, 16):
        for coluna in range(2, _ULTIMA + 1):
            _cabecalho(capa.cell(linha, coluna))
    for celula in ("E16", "G16"):
        capa[celula].number_format = _FMT_DATA
        capa[celula].fill = _fundo(_AZUL)
    capa.row_dimensions[15].height = 20
    capa.row_dimensions[16].height = 18

    origem_em = capa.cell(17, 2, "EM")
    origem_em.font = _fonte(10, True)
    origem_em.hyperlink = f"#{_q(em.aba[:31])}!A1"
    conta_em = capa.cell(17, 3, pipeline.parametros.get("conta", ""))
    conta_em.number_format = "@"
    conta_em.font = _fonte(10)
    capa.cell(17, 4, pipeline.parametros.get("descricao", "")).font = _fonte(10)
    capa["F17"] = f"=SUM({tabela_em}[VALOR])"
    capa["I17"] = "=ROUND(H18-F17,2)"
    capa.merge_cells("I17:J17")
    _linha_valor(capa, 17)
    capa["F17"].font = _fonte(10, True, _NAVY)
    capa["I17"].font = _fonte(10, True)

    origem_se5 = capa.cell(18, 2, "Protheus")
    origem_se5.font = _fonte(10, True)
    origem_se5.hyperlink = f"#{_q(se5.aba[:31])}!A1"
    naturezas = capa.cell(18, 3, _naturezas(pipeline))
    naturezas.number_format = "@"
    naturezas.font = _fonte(10)
    produtos = _produtos(pipeline)
    texto_se5 = "SE5" if not produtos else f"SE5  ·  produto {produtos}"
    capa.cell(18, 4, texto_se5).font = _fonte(10)
    capa["H18"] = f"=SUM({tabela_se5}[VALOR])"
    capa["I18"] = "financeiro − baixas"
    capa["I18"].font = _fonte(8, cor=_CINZA_TEXTO, italico=True)
    _linha_valor(capa, 18)
    capa["H18"].font = _fonte(10, True, _NAVY)
    capa["I18"].font = _fonte(8, cor=_CINZA_TEXTO, italico=True)
    capa["I18"].number_format = "@"

    for faixa, formula, verde in (
        ("I17:J17", "ABS(I17)<=Tolerancia", True),
        ("I17:J17", "ABS(I17)>Tolerancia", False),
    ):
        capa.conditional_formatting.add(
            faixa,
            FormulaRule(
                formula=[formula],
                fill=PatternFill("solid", fgColor=_VERDE_FUNDO if verde else _VERMELHO_FUNDO),
                font=Font(name=_FONTE, bold=True, color=_VERDE_TEXTO if verde else _VERMELHO_TEXTO),
            ),
        )

    _secao(capa, 20, "Conclusão", _ULTIMA)
    capa.merge_cells("B21:J22")
    capa["B21"] = (
        '=IF(ABS(I17)<=Tolerancia,'
        '"As baixas do EM e o financeiro da SE5 estão conciliados no período, dentro da tolerância.",'
        '"A soma do período apresenta diferença. Conferir as abas de baixas e do financeiro antes de concluir.")'
    )
    capa["B21"].font = _fonte(10)
    capa["B21"].alignment = Alignment(wrap_text=True, vertical="center")
    capa.row_dimensions[21].height = 20
    capa.row_dimensions[22].height = 20
    nota = capa.cell(
        24, 2,
        "A data do pagamento no EM e a data da SE5 podem cair em dias diferentes. A conta fecha pela soma do período.",
    )
    nota.font = _fonte(9, italico=True, cor=_CINZA_TEXTO)
    capa.merge_cells("B24:J24")

    capa.freeze_panes = "B15"
    capa.sheet_view.view = "normal"
    capa.sheet_view.zoomScale = 110
    capa.sheet_properties.tabColor = _NAVY
    capa.page_setup.paperSize = capa.PAPERSIZE_A4
    capa.print_title_rows = "1:16"
    _retrato(capa)

    for fonte in pipeline.fontes:
        script = scripts[fonte.script]
        quadro = _maiusculas(extratos.get(fonte.alias, pd.DataFrame()))
        aba = wb.create_sheet(fonte.aba[:31])
        _aba_dados(aba, pipeline, fonte, script, quadro)
        aba.sheet_view.view = "normal"
        aba.sheet_properties.tabColor = _NAVY

    caminho.parent.mkdir(parents=True, exist_ok=True)
    wb.calculation.fullCalcOnLoad = True
    wb.active = 0
    wb.save(_caminho_gravavel(caminho))
