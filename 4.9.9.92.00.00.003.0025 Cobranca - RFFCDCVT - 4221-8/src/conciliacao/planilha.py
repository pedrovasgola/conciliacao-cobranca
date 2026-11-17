"""Gera o Excel da conciliação.

Abas, na ordem:
  WP-01 Saldos   resumo: saldo de cada conta, diferença, composição e status
  WP-0x <fonte>  uma por script pronto, dados como vieram do banco (Tabela do Excel)
  Conciliação    diferença por dia, toda calculada por fórmula sobre as abas de dados
  Composição     linhas de detalhe dos dias fora da tolerância (quando houver script)
  Execução       trilha de auditoria: scripts, parâmetros, linhas e conferência motor x planilha

Os dados do banco entram como valor. Tudo o que é conta (saldo, soma do dia,
diferença, status) entra como fórmula, para quem revisar conseguir seguir a origem.
"""

from __future__ import annotations

import ctypes
import os
import re
import shutil
import tempfile
from datetime import date, datetime, time
from decimal import Decimal
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.worksheet import Worksheet

from conciliacao.catalogo import Pipeline, Script, fonte_pronta, regras_exclusao
from conciliacao.conciliar import ResultadoConciliacao
from conciliacao.erros import ErroConciliacao
from conciliacao.exclusoes import COLUNA_MOTIVO

# --------------------------------------------------------------------------- estilo

ABA_SALDOS = "WP-01 Saldos"
ABA_CONCILIACAO = "Conciliação"
ABA_COMPOSICAO = "Composição"
ABA_EXECUCAO = "Execução"
ABA_EXCLUSOES = "Exclusões"

_FONTE = "Arial"
_NAVY = "1F3864"
_AZUL = "2E75B6"
_AZUL_CLARO = "DDEBF7"
_CINZA = "F2F2F2"
_CINZA_TEXTO = "595959"
_CINZA_BORDA = "BFBFBF"
_VERDE_FUNDO = "C6EFCE"
_VERDE_TEXTO = "006100"
_VERMELHO_FUNDO = "FFC7CE"
_VERMELHO_TEXTO = "9C0006"
_AMARELO = "FFF2CC"
_AZUL_INPUT = "0000FF"

_FMT_VALOR = '_-* #,##0.00_-;\\-* #,##0.00_-;_-* "-"??_-;_-@_-'
_FMT_DATA = "DD/MM/YYYY"
_FMT_DATA_HORA = "DD/MM/YYYY HH:MM"

_FINO = Side(style="thin", color=_CINZA_BORDA)
_MEDIO = Side(style="medium", color=_NAVY)
_DUPLO = Side(style="double", color=_NAVY)


def _fonte(tamanho=10, negrito=False, cor="000000", italico=False) -> Font:
    return Font(name=_FONTE, size=tamanho, bold=negrito, color=cor, italic=italico)


def _fundo(cor: str) -> PatternFill:
    return PatternFill("solid", fgColor=cor)


_TEXTO = _fonte()
_ROTULO = _fonte(9, True, _CINZA_TEXTO)
_SECAO = _fonte(11, True, _NAVY)
_CABECALHO = _fonte(10, True, "FFFFFF")
_INPUT = _fonte(10, False, _AZUL_INPUT)


def _q(aba: str) -> str:
    """Nome de aba pronto para fórmula: 'WP-02 Protheus'."""
    return "'" + aba.replace("'", "''") + "'"


def _coluna_tabela(nome: str) -> str:
    """Escapa o nome da coluna dentro de uma referência estruturada."""
    return re.sub(r"(['\[\]#])", r"'\1", str(nome))


def _nome_tabela(alias: str) -> str:
    base = re.sub(r"[^A-Za-z0-9_]", "_", alias)
    return f"tbl_{base}"[:250]


def _faixa(ws: Worksheet, titulo: str, subtitulo: str, ultima_coluna: int) -> None:
    """Faixa azul do topo (linhas 2 e 3), de B até a última coluna usada."""
    ws.row_dimensions[1].height = 8
    ws.row_dimensions[2].height = 26
    ws.row_dimensions[3].height = 18
    fim = get_column_letter(max(ultima_coluna, 3))
    for linha, texto, fonte in (
        (2, titulo, _fonte(15, True, "FFFFFF")),
        (3, subtitulo, _fonte(10, False, "D9E1F2")),
    ):
        for coluna in range(2, max(ultima_coluna, 3) + 1):
            ws.cell(linha, coluna).fill = _fundo(_NAVY)
        ws.merge_cells(f"B{linha}:{fim}{linha}")
        celula = ws.cell(linha, 2, texto)
        celula.font = fonte
        celula.alignment = Alignment(vertical="center", indent=1)
    ws.column_dimensions["A"].width = 2.5
    ws.sheet_view.showGridLines = False


def _secao(ws: Worksheet, linha: int, texto: str, ultima_coluna: int) -> None:
    celula = ws.cell(linha, 2, texto.upper())
    celula.font = _SECAO
    for coluna in range(2, ultima_coluna + 1):
        ws.cell(linha, coluna).border = Border(bottom=_MEDIO)
    ws.row_dimensions[linha].height = 20


def _par(ws: Worksheet, linha: int, coluna: int, rotulo: str, valor, formato: str | None = None,
         fonte: Font | None = None, mesclar_ate: int | None = None, fundo: str | None = None):
    """Rótulo cinza + valor ao lado, com linha fina embaixo."""
    r = ws.cell(linha, coluna, rotulo)
    r.font = _ROTULO
    r.alignment = Alignment(vertical="center")
    r.border = Border(bottom=_FINO)
    v = ws.cell(linha, coluna + 1, valor)
    v.font = fonte or _fonte(10, True)
    v.alignment = Alignment(horizontal="left", vertical="center")
    if formato:
        v.number_format = formato
    ultima = mesclar_ate or coluna + 1
    for c in range(coluna + 1, ultima + 1):
        ws.cell(linha, c).border = Border(bottom=_FINO)
        if fundo:
            ws.cell(linha, c).fill = _fundo(fundo)
    if mesclar_ate and mesclar_ate > coluna + 1:
        ws.merge_cells(start_row=linha, start_column=coluna + 1, end_row=linha, end_column=mesclar_ate)
    return v


def _cabecalho(celula, centro=True, cor=_NAVY) -> None:
    celula.font = _CABECALHO
    celula.fill = _fundo(cor)
    celula.alignment = Alignment(horizontal="center" if centro else "left", vertical="center", wrap_text=True)
    celula.border = Border(left=_FINO, right=_FINO, top=_FINO, bottom=_FINO)


def _status_cf(ws: Worksheet, faixa: str) -> None:
    """Verde para CONCILIADO, vermelho para o resto."""
    ws.conditional_formatting.add(
        faixa,
        CellIsRule(operator="equal", formula=['"CONCILIADO"'],
                   fill=_fundo(_VERDE_FUNDO), font=Font(name=_FONTE, bold=True, color=_VERDE_TEXTO)),
    )
    ws.conditional_formatting.add(
        faixa,
        CellIsRule(operator="notEqual", formula=['"CONCILIADO"'],
                   fill=_fundo(_VERMELHO_FUNDO), font=Font(name=_FONTE, bold=True, color=_VERMELHO_TEXTO)),
    )


def _retrato(ws: Worksheet, paisagem=True) -> None:
    ws.page_setup.orientation = "landscape" if paisagem else "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = ws.page_margins.right = 0.4


# --------------------------------------------------------------------------- valores

_ISO_DATA = re.compile(r"^(\d{4})-(\d{2})-(\d{2})(?:[ T].*)?$")
_AAAAMMDD = re.compile(r"^\d{8}$")


def _como_data(valor):
    if isinstance(valor, datetime):
        if valor.tzinfo is not None:
            valor = valor.replace(tzinfo=None)
        if valor.time() == time.min:
            return valor.date()
        return valor
    if isinstance(valor, date):
        return valor
    if isinstance(valor, str):
        bruto = valor.strip()
        iso = _ISO_DATA.match(bruto)
        if iso:
            return date(int(iso.group(1)), int(iso.group(2)), int(iso.group(3)))
        if _AAAAMMDD.match(bruto):
            try:
                return datetime.strptime(bruto, "%Y%m%d").date()
            except ValueError:
                return None
    return None


def _limpar(valor):
    if valor is None:
        return None
    if isinstance(valor, np.generic):
        valor = valor.item()
    try:
        if pd.isna(valor):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(valor, pd.Timestamp):
        valor = valor.to_pydatetime()
    convertido = _como_data(valor)
    if convertido is not None:
        return convertido
    if isinstance(valor, str):
        return ILLEGAL_CHARACTERS_RE.sub("", valor)
    if isinstance(valor, (int, float, bool, Decimal, date, datetime)):
        return valor
    return str(valor)


_NOME_MONETARIO = re.compile(r"(VALOR|VL_|_VL|SALDO|SDO|DEBITO_VL|CREDITO_VL)", re.IGNORECASE)


def _coluna_monetaria(nome: str, serie: pd.Series) -> bool:
    """Coluna de valor (formata e entra no total). Códigos numéricos ficam de fora."""
    amostra = serie.dropna().head(2000)
    if amostra.empty:
        return bool(_NOME_MONETARIO.search(nome))
    numeros = []
    for item in amostra:
        if isinstance(item, np.generic):
            item = item.item()
        if isinstance(item, bool) or not isinstance(item, (int, float, Decimal)):
            return False
        numeros.append(item)
    if _NOME_MONETARIO.search(nome):
        return True
    return any(float(x) != int(float(x)) for x in numeros)


# --------------------------------------------------------------------------- nomes de conta

_COLUNAS_NOME_CONTA = ("CT001_DS_NOMCT", "NOME_C_CONTABIL")


def _nome_conta(df: pd.DataFrame, codigo: str = "") -> str:
    if df is None or df.empty:
        return ""
    for coluna in _COLUNAS_NOME_CONTA:
        if coluna not in df.columns:
            continue
        serie = df[coluna].dropna().astype("string").str.strip()
        serie = serie[serie != ""]
        if len(serie):
            return str(serie.iloc[0])
    # CT2: o nome da conta está no lado (débito ou crédito) em que ela aparece.
    if codigo and {"CT2_DEBITO", "CT2_YDESDB"} <= set(df.columns):
        debito = df["CT2_DEBITO"].astype("string").str.strip() == codigo
        nomes = df.loc[debito.fillna(False), "CT2_YDESDB"].dropna().astype("string").str.strip()
        if len(nomes[nomes != ""]):
            return str(nomes[nomes != ""].iloc[0])
    if codigo and {"CT2_CREDIT", "CT2_YDESCR"} <= set(df.columns):
        credito = df["CT2_CREDIT"].astype("string").str.strip() == codigo
        nomes = df.loc[credito.fillna(False), "CT2_YDESCR"].dropna().astype("string").str.strip()
        if len(nomes[nomes != ""]):
            return str(nomes[nomes != ""].iloc[0])
    return ""


def nome_conta(df: pd.DataFrame | None, codigo: str = "") -> str:
    """Nome da conta como vem do banco (CT001_DS_NOMCT na Sinqia, CT2_YDESDB/YDESCR no Protheus)."""
    return _nome_conta(df if df is not None else pd.DataFrame(), (codigo or "").strip())


def _texto_contra(fonte, df: pd.DataFrame) -> str:
    """'contrapartida 11010201029 — BANCO SPIX SENFF' para fontes com parâmetro contra."""
    if fonte is None or "contra" not in fonte.parametros:
        return ""
    contra = (fonte.parametros.get("contra") or "").strip()
    if not contra:
        return "todas as contrapartidas"
    nome = _nome_conta(df, contra)
    return f"contrapartida {contra}" + (f" — {nome.strip()}" if nome else "")


def _codigo_conta(fonte) -> str:
    if fonte is None:
        return ""
    codigo = (fonte.parametros.get("conta") or "").strip()
    if codigo:
        return codigo
    natureza = (fonte.parametros.get("natureza") or "").strip()
    outra = (fonte.parametros.get("natureza_2") or "").strip()
    if natureza and outra and outra != natureza:
        return f"{natureza} / {outra}"
    if natureza:
        return natureza
    produto = (fonte.parametros.get("produto") or "").strip()
    outro = (fonte.parametros.get("produto_2") or "").strip()
    if produto and outro and outro != produto:
        return f"{produto} / {outro}"
    return produto


def _rotulo_conta(fonte, df: pd.DataFrame) -> str:
    if fonte is None:
        return "(não informada)"
    codigo = _codigo_conta(fonte)
    if not codigo:
        natureza = (fonte.parametros.get("natureza") or "").strip()
        outra = (fonte.parametros.get("natureza_2") or "").strip()
        produto = (fonte.parametros.get("produto") or "").strip()
        outro_produto = (fonte.parametros.get("produto_2") or "").strip()
        if natureza and outra and outra != natureza:
            codigo = f"{natureza} / {outra}"
        elif natureza:
            codigo = natureza
        elif produto and outro_produto and outro_produto != produto:
            codigo = f"{produto} / {outro_produto}"
        else:
            codigo = produto
    nome = _nome_conta(df, codigo)
    if codigo and nome:
        return f"{codigo}  —  {nome}"
    return codigo or nome or "(não informada)"


def _sistema(aba: str) -> str:
    """'WP-02 Protheus' -> 'Protheus'."""
    return re.sub(r"^WP-?\d+\s*[-–]?\s*", "", aba).strip() or aba


def _saldo_anterior(quadro: pd.DataFrame | None) -> float:
    """Saldo da Sinqia antes da data inicial. Zero quando o balancete não veio."""
    if quadro is None or quadro.empty or "SALDO_ANTERIOR" not in quadro.columns:
        return 0.0
    valor = pd.to_numeric(quadro["SALDO_ANTERIOR"], errors="coerce").fillna(0).iloc[0]
    return round(float(valor), 2)


def _caminho_unc(absoluto: str) -> str:
    """A: vira o caminho de rede. É esse tamanho que o Windows limita em 260."""
    if os.name != "nt" or absoluto.startswith("\\\\"):
        return absoluto
    tamanho = ctypes.c_ulong(4096)
    buffer = ctypes.create_unicode_buffer(4096)
    codigo = ctypes.windll.mpr.WNetGetUniversalNameW(absoluto, 2, buffer, ctypes.byref(tamanho))
    if codigo != 0:
        return absoluto
    ponteiro = ctypes.cast(buffer, ctypes.POINTER(ctypes.c_void_p))[0]
    if not ponteiro:
        return absoluto
    return ctypes.wstring_at(ponteiro) or absoluto


def _caminho_gravavel(caminho: Path) -> str:
    """Prefixo longo para a pasta da frota, cujo nome na rede passa de 260 caracteres."""
    absoluto = os.path.abspath(caminho)
    if os.name != "nt":
        return absoluto
    alvo = _caminho_unc(absoluto)
    if len(alvo) < 240 or alvo.startswith("\\\\?\\"):
        return alvo
    if alvo.startswith("\\\\"):
        return "\\\\?\\UNC\\" + alvo[2:]
    return "\\\\?\\" + alvo


# --------------------------------------------------------------------------- gravação

def gravar(
    caminho: Path,
    pipeline: Pipeline,
    scripts: dict[str, Script],
    extratos: dict[str, pd.DataFrame],
    resultado: ResultadoConciliacao,
    composicao: pd.DataFrame | None,
    inicio: date,
    fim: date,
    feito_por: str,
    excluidos: dict[str, pd.DataFrame] | None = None,
) -> None:
    destino = Path(caminho)
    fd, bruto = tempfile.mkstemp(prefix="conciliacao_", suffix=".xlsx")
    os.close(fd)
    temporario = Path(bruto)
    movido = False
    try:
        _gravar(
            temporario, pipeline, scripts, extratos, resultado, composicao, inicio, fim, feito_por,
            excluidos or {},
        )
        destino.parent.mkdir(parents=True, exist_ok=True)
        gravavel = _caminho_gravavel(destino)
        try:
            os.replace(temporario, gravavel)
            movido = True
        except OSError:
            shutil.copyfile(temporario, gravavel)
            movido = True
    except PermissionError as exc:
        raise ErroConciliacao(
            f"Não consegui salvar {destino}. Feche o arquivo no Excel e rode de novo."
        ) from exc
    except OSError as exc:
        raise ErroConciliacao(
            f"Não consegui salvar {destino}. {exc}"
        ) from exc
    finally:
        if not movido and temporario.exists():
            try:
                temporario.unlink()
            except OSError:
                pass


class _InfoTabela:
    """Onde ficou cada fonte, para as fórmulas apontarem para ela."""

    def __init__(self, fonte, tabela: str, colunas: list[str], linha_total: int | None):
        self.fonte = fonte
        self.tabela = tabela
        self.colunas = colunas
        self.linha_total = linha_total

    def ref(self, coluna: str) -> str:
        return f"{self.tabela}[{_coluna_tabela(coluna)}]"


def _gravar(
    caminho: Path,
    pipeline: Pipeline,
    scripts: dict[str, Script],
    extratos: dict[str, pd.DataFrame],
    resultado: ResultadoConciliacao,
    composicao: pd.DataFrame | None,
    inicio: date,
    fim: date,
    feito_por: str,
    excluidos: dict[str, pd.DataFrame],
) -> None:
    wb = Workbook()
    excluidos = {alias: df for alias, df in excluidos.items() if df is not None and len(df)}
    saldos = wb.active
    saldos.title = ABA_SALDOS

    # Nomes definidos: as outras abas leem período e tolerância daqui.
    for nome, celula in (("DataInicial", "$I$5"), ("DataFinal", "$I$6"), ("Tolerancia", "$C$7")):
        wb.defined_names[nome] = DefinedName(nome, attr_text=f"{_q(ABA_SALDOS)}!{celula}")

    tabelas: dict[str, _InfoTabela] = {}
    for fonte in pipeline.fontes:
        script = scripts[fonte.script]
        if not fonte_pronta(fonte, script):
            continue
        quadro = extratos.get(fonte.alias, pd.DataFrame())
        ws = wb.create_sheet(fonte.aba[:31])
        tabelas[fonte.alias] = _aba_dados(ws, pipeline, fonte, script, quadro,
                                          len(excluidos.get(fonte.alias, ())))

    esq = tabelas.get(pipeline.esquerda)
    dire = tabelas.get(pipeline.direita)
    if esq is None or dire is None:
        raise ErroConciliacao("Os dois lados da conciliação precisam ter aba de dados.")
    col_esq = scripts[esq.fonte.script].coluna
    col_dir = scripts[dire.fonte.script].coluna

    conc = wb.create_sheet(ABA_CONCILIACAO)
    pos_conc = _aba_conciliacao(conc, pipeline, resultado, esq, dire, col_esq, col_dir)

    if pipeline.detalhe and composicao is not None:
        comp = wb.create_sheet(ABA_COMPOSICAO)
        _aba_composicao(comp, pipeline, composicao)

    if excluidos:
        exc = wb.create_sheet(ABA_EXCLUSOES)
        _aba_exclusoes(exc, pipeline, scripts, excluidos)

    exe = wb.create_sheet(ABA_EXECUCAO)
    _aba_saldos(saldos, wb, pipeline, extratos, resultado, esq, dire, col_esq, col_dir,
                pos_conc, inicio, fim, feito_por, excluidos, scripts)
    _aba_execucao(exe, pipeline, scripts, extratos, resultado, pos_conc, feito_por, excluidos)

    saldos.sheet_properties.tabColor = _NAVY
    conc.sheet_properties.tabColor = _NAVY
    exe.sheet_properties.tabColor = _NAVY
    wb.active = 0
    wb.calculation.fullCalcOnLoad = True
    wb.save(caminho)


# --------------------------------------------------------------------------- abas de dados

_LINHA_TABELA = 11


def _aba_dados(ws: Worksheet, pipeline: Pipeline, fonte, script: Script, df: pd.DataFrame,
               n_excluidas: int = 0) -> _InfoTabela:
    colunas = [str(c) for c in df.columns]
    ultima = max(len(colunas) + 1, 8)
    _faixa(ws, fonte.aba, fonte.resumo or script.descricao, ultima)
    ws.sheet_properties.tabColor = _NAVY

    col_data = script.coluna("data")
    col_valor = script.coluna("valor")
    tabela = _nome_tabela(fonte.alias)

    _par(ws, 5, 2, "Empresa", f"={_q(ABA_SALDOS)}!$C$5", mesclar_ate=5)
    contra = _texto_contra(fonte, df)
    _par(ws, 6, 2, "Conta", _rotulo_conta(fonte, df) + (f"   ·   {contra}" if contra else ""), mesclar_ate=8)
    _par(ws, 7, 2, "Período", "=DataInicial", _FMT_DATA)
    ws.cell(7, 4, "=DataFinal").number_format = _FMT_DATA
    ws.cell(7, 4).font = _fonte(10, True)
    ws.cell(7, 4).alignment = Alignment(horizontal="left")
    ws.cell(7, 4).border = Border(bottom=_FINO)
    _par(ws, 8, 2, "Linhas na conciliação", len(df), "#,##0")
    if n_excluidas:
        aviso = ws.cell(8, 4, f"+ {n_excluidas} linha(s) excluída(s) da soma: ver aba {ABA_EXCLUSOES}")
        aviso.font = Font(name=_FONTE, size=9, color=_VERMELHO_TEXTO, underline="single")
        aviso.hyperlink = f"#{_q(ABA_EXCLUSOES)}!A1"
    if col_valor in colunas:
        total = _par(ws, 9, 2, "Total da conta", f"=SUM({tabela}[{_coluna_tabela(col_valor)}])",
                     _FMT_VALOR, fonte=_fonte(11, True, _NAVY), mesclar_ate=4)
        total.alignment = Alignment(horizontal="left")
        nota = ws.cell(9, 6, f"soma de {col_valor}, todas as linhas, sem filtro")
        nota.font = _fonte(8, italico=True, cor=_CINZA_TEXTO)

    if not colunas:
        ws.cell(_LINHA_TABELA, 2, "A consulta não devolveu colunas.").font = _TEXTO
        return _InfoTabela(fonte, tabela, colunas, None)

    linha = _LINHA_TABELA
    monetarias = set()
    for indice, nome in enumerate(colunas, start=2):
        celula = ws.cell(linha, indice, nome)
        celula.font = _CABECALHO
        celula.fill = _fundo(_AZUL)
        celula.alignment = Alignment(horizontal="center", vertical="center")
        if nome in (col_data, col_valor):
            celula.fill = _fundo(_NAVY)
            papel = "data" if nome == col_data else "valor"
            celula.comment = Comment(
                f"Coluna de {papel} usada na conciliação (@coluna.{papel} de {script.caminho.name}).",
                "Conciliação",
            )
        if _coluna_monetaria(nome, df.iloc[:, indice - 2]):
            monetarias.add(nome)

    valores = df.to_numpy()
    for deslocamento, registro in enumerate(valores, start=1):
        for indice, valor in enumerate(registro, start=2):
            limpo = _limpar(valor)
            celula = ws.cell(linha + deslocamento, indice, limpo)
            celula.font = _TEXTO
            if isinstance(limpo, datetime):
                celula.number_format = _FMT_DATA_HORA if limpo.time() != time.min else _FMT_DATA
            elif isinstance(limpo, date):
                celula.number_format = _FMT_DATA
            elif colunas[indice - 2] in monetarias and isinstance(limpo, (int, float, Decimal)):
                celula.number_format = _FMT_VALOR
            elif isinstance(limpo, float) and limpo.is_integer():
                celula.number_format = "0"

    ultima_linha = linha + max(len(df), 1)  # Tabela do Excel precisa de ao menos uma linha
    fim_col = get_column_letter(len(colunas) + 1)
    tab = Table(displayName=tabela, ref=f"B{linha}:{fim_col}{ultima_linha}")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(tab)
    ws.freeze_panes = f"C{linha + 1}"

    # Linha de total abaixo da tabela: SUBTOTAL respeita o filtro, como na WP antiga.
    linha_total = ultima_linha + 2
    rot = ws.cell(linha_total, 2, "Total (linhas visíveis)")
    rot.font = _fonte(10, True, _NAVY)
    for indice, nome in enumerate(colunas, start=2):
        celula = ws.cell(linha_total, indice)
        celula.border = Border(top=_FINO, bottom=_DUPLO)
        if nome in monetarias:
            celula.value = f"=SUBTOTAL(109,{tabela}[{_coluna_tabela(nome)}])"
            celula.number_format = _FMT_VALOR
            celula.font = _fonte(10, True, _NAVY)
    ws.cell(linha_total, 2).border = Border(top=_FINO, bottom=_DUPLO)

    for indice, nome in enumerate(colunas, start=2):
        amostra = [len(str(nome)) + 3]
        serie = df.iloc[:, indice - 2].head(300)
        for item in serie:
            if item is None:
                continue
            if isinstance(item, float):
                amostra.append(len(f"{item:,.2f}") + 2)
            else:
                amostra.append(len(str(item).rstrip()))
        ws.column_dimensions[get_column_letter(indice)].width = min(max(max(amostra) + 2, 11), 44)
    ws.row_dimensions[linha].height = 22
    _retrato(ws)
    return _InfoTabela(fonte, tabela, colunas, linha_total)


# --------------------------------------------------------------------------- conciliação

_CONC_CAB = 9


def _aba_conciliacao(ws, pipeline: Pipeline, resultado: ResultadoConciliacao,
                     esq: _InfoTabela, dire: _InfoTabela, col_esq, col_dir) -> dict:
    _faixa(ws, "Conciliação diária", pipeline.nome, 7)
    sinal = "+" if pipeline.operacao == "soma" else "−"
    if pipeline.id == "cobranca":
        rotulo_esq, rotulo_dir = "Baixas do EM", "Financeiro"
        texto_regra = "Diferença = baixas − financeiro · conciliado se |diferença| ≤ tolerância"
    else:
        rotulo_esq, rotulo_dir = "Lado contábil", "Lado financeiro"
        texto_regra = f"Diferença = contábil {sinal} financeiro · conciliado se |diferença| ≤ tolerância"
    _par(ws, 5, 2, rotulo_esq, f"{esq.fonte.aba}  ·  soma de {col_esq('valor')} por {col_esq('data')}", mesclar_ate=7)
    _par(ws, 6, 2, rotulo_dir, f"{dire.fonte.aba}  ·  soma de {col_dir('valor')} por {col_dir('data')}", mesclar_ate=7)
    _par(ws, 7, 2, "Regra", texto_regra, mesclar_ate=5)
    r = ws.cell(7, 6, "Tolerância")
    r.font = _ROTULO
    r.alignment = Alignment(horizontal="right")
    t = ws.cell(7, 7, "=Tolerancia")
    t.number_format = _FMT_VALOR
    t.font = _fonte(10, True)

    cab = ["DATA", esq.fonte.aba, dire.fonte.aba, "DIFERENÇA", "STATUS", "JUSTIFICATIVA"]
    for indice, nome in enumerate(cab, start=2):
        _cabecalho(ws.cell(_CONC_CAB, indice, nome))
    ws.cell(_CONC_CAB, 7).fill = _fundo(_AZUL)
    ws.cell(_CONC_CAB, 7).comment = Comment("Preencha a explicação dos dias com diferença.", "Conciliação")
    ws.row_dimensions[_CONC_CAB].height = 30

    ref_esq_v, ref_esq_d = esq.ref(col_esq("valor")), esq.ref(col_esq("data"))
    ref_dir_v, ref_dir_d = dire.ref(col_dir("valor")), dire.ref(col_dir("data"))
    datas = list(resultado.por_dia["DATA"]) if len(resultado.por_dia) else []
    primeira = _CONC_CAB + 1
    for deslocamento, dia in enumerate(datas or [None]):
        linha = primeira + deslocamento
        a = ws.cell(linha, 2, dia)
        a.number_format = _FMT_DATA
        a.alignment = Alignment(horizontal="center")
        if dia is None:
            for coluna in (3, 4, 5):
                ws.cell(linha, coluna, 0).number_format = _FMT_VALOR
            ws.cell(linha, 6, "SEM MOVIMENTO")
        else:
            ws.cell(linha, 3, f"=ROUND(SUMIFS({ref_esq_v},{ref_esq_d},$B{linha}),2)")
            ws.cell(linha, 4, f"=ROUND(SUMIFS({ref_dir_v},{ref_dir_d},$B{linha}),2)")
            ws.cell(linha, 5, f"=ROUND(C{linha}{'+' if pipeline.operacao == 'soma' else '-'}D{linha},2)")
            ws.cell(linha, 6, f'=IF(ABS(E{linha})<=Tolerancia,"CONCILIADO","DIFERENÇA")')
        for coluna in range(2, 8):
            celula = ws.cell(linha, coluna)
            celula.font = _TEXTO
            celula.border = Border(bottom=_FINO, left=_FINO, right=_FINO)
            if coluna in (3, 4, 5):
                celula.number_format = _FMT_VALOR
        ws.cell(linha, 6).alignment = Alignment(horizontal="center")
        ws.cell(linha, 6).font = _fonte(9, True)
        ws.cell(linha, 7).fill = _fundo(_AMARELO)
        ws.cell(linha, 7).font = _INPUT
    ultima = primeira + max(len(datas), 1) - 1

    total = ultima + 1
    ws.cell(total, 2, "Total")
    ws.cell(total, 3, f"=SUBTOTAL(109,C{primeira}:C{ultima})")
    ws.cell(total, 4, f"=SUBTOTAL(109,D{primeira}:D{ultima})")
    ws.cell(total, 5, f"=SUBTOTAL(109,E{primeira}:E{ultima})")
    ws.cell(total, 6, f'=IF(ABS(E{total})<=Tolerancia,"CONCILIADO","DIFERENÇA")')
    for coluna in range(2, 8):
        celula = ws.cell(total, coluna)
        celula.font = _fonte(10, True, _NAVY)
        celula.fill = _fundo(_AZUL_CLARO)
        celula.border = Border(top=_MEDIO, bottom=_DUPLO)
        if coluna in (3, 4, 5):
            celula.number_format = _FMT_VALOR
    ws.cell(total, 6).alignment = Alignment(horizontal="center")

    _status_cf(ws, f"F{primeira}:F{total}")
    ws.conditional_formatting.add(
        f"E{primeira}:E{ultima}",
        FormulaRule(formula=[f"ABS(E{primeira})>Tolerancia"], font=Font(name=_FONTE, bold=True, color=_VERMELHO_TEXTO)),
    )
    ws.auto_filter.ref = f"B{_CONC_CAB}:G{ultima}"
    ws.freeze_panes = f"C{primeira}"
    for letra, largura in zip("BCDEFG", (13, 22, 22, 18, 15, 46)):
        ws.column_dimensions[letra].width = largura
    _retrato(ws, paisagem=False)
    return {"primeira": primeira, "ultima": ultima, "total": total}


# --------------------------------------------------------------------------- composição

def _aba_composicao(ws, pipeline: Pipeline, quadro: pd.DataFrame) -> None:
    colunas = [str(c) for c in quadro.columns]
    _faixa(ws, "Composição dos dias com diferença", pipeline.fonte(pipeline.detalhe).resumo,
           max(len(colunas) + 1, 7))
    ws.sheet_properties.tabColor = _NAVY
    if quadro.empty:
        ws.cell(5, 2, "Nenhuma linha de detalhe nos dias com diferença.").font = _TEXTO
        return
    linha = 5
    for indice, nome in enumerate(colunas, start=2):
        _cabecalho(ws.cell(linha, indice, nome))
    for deslocamento, registro in enumerate(quadro.to_numpy(), start=1):
        for indice, valor in enumerate(registro, start=2):
            limpo = _limpar(valor)
            celula = ws.cell(linha + deslocamento, indice, limpo)
            celula.font = _TEXTO
            celula.border = Border(bottom=_FINO)
            if isinstance(limpo, (date, datetime)):
                celula.number_format = _FMT_DATA
            elif isinstance(limpo, (float, Decimal)):
                celula.number_format = _FMT_VALOR
    ultima = linha + len(quadro)
    ws.auto_filter.ref = f"B{linha}:{get_column_letter(len(colunas) + 1)}{ultima}"
    ws.freeze_panes = f"B{linha + 1}"
    for indice, nome in enumerate(colunas, start=2):
        serie = quadro.iloc[:, indice - 2].head(300)
        largura = max([len(nome) + 3] + [len(str(x)) for x in serie if x is not None])
        ws.column_dimensions[get_column_letter(indice)].width = min(max(largura + 2, 11), 44)
    _retrato(ws)


# --------------------------------------------------------------------------- exclusões

def _aba_exclusoes(ws, pipeline: Pipeline, scripts, excluidos: dict[str, pd.DataFrame]) -> None:
    """Linhas tiradas da conciliação por regra. Ficam aqui, com o motivo, para quem revisar."""
    largura_max = max(len(df.columns) for df in excluidos.values()) + 1
    _faixa(ws, "Linhas excluídas da conciliação",
           "Não entram nas abas de dados, nos saldos nem na soma do dia", max(largura_max, 7))
    ws.sheet_properties.tabColor = _NAVY
    larguras: dict[int, int] = {}
    linha = 5
    for fonte in pipeline.fontes:
        quadro = excluidos.get(fonte.alias)
        if quadro is None or quadro.empty:
            continue
        script = scripts[fonte.script]
        col_valor = script.coluna("valor")
        colunas = [COLUNA_MOTIVO] + [str(c) for c in quadro.columns if c != COLUNA_MOTIVO]
        quadro = quadro[colunas]
        _secao(ws, linha, f"{fonte.aba}  ·  {len(quadro)} linha(s)", max(len(colunas) + 1, 7))
        linha += 1
        for indice, nome in enumerate(colunas, start=2):
            _cabecalho(ws.cell(linha, indice, nome), cor=_NAVY if nome in (COLUNA_MOTIVO, col_valor) else _AZUL)
        primeira = linha + 1
        for deslocamento, registro in enumerate(quadro.to_numpy(), start=1):
            for indice, valor in enumerate(registro, start=2):
                limpo = _limpar(valor)
                celula = ws.cell(linha + deslocamento, indice, limpo)
                celula.font = _TEXTO
                celula.border = Border(bottom=_FINO)
                if isinstance(limpo, (date, datetime)):
                    celula.number_format = _FMT_DATA
                elif isinstance(limpo, (float, Decimal)) and _coluna_monetaria(colunas[indice - 2], quadro.iloc[:, indice - 2]):
                    celula.number_format = _FMT_VALOR
                elif isinstance(limpo, float) and limpo.is_integer():
                    celula.number_format = "0"
                larguras[indice] = max(larguras.get(indice, 0), len(str(limpo if limpo is not None else "")),
                                       len(colunas[indice - 2]) + 3)
        ultima = linha + len(quadro)
        linha = ultima + 1
        ws.cell(linha, 2, "Total excluído").font = _fonte(10, True, _NAVY)
        if col_valor in colunas:
            idx = colunas.index(col_valor) + 2
            letra = get_column_letter(idx)
            t = ws.cell(linha, idx, f"=SUM({letra}{primeira}:{letra}{ultima})")
            t.number_format = _FMT_VALOR
            t.font = _fonte(10, True, _NAVY)
        for col in range(2, len(colunas) + 2):
            ws.cell(linha, col).border = Border(top=_FINO, bottom=_DUPLO)
        linha += 3
    for indice, largura in larguras.items():
        ws.column_dimensions[get_column_letter(indice)].width = min(max(largura + 2, 11), 46)
    _retrato(ws)


# --------------------------------------------------------------------------- WP-01 saldos

def _aba_saldos(ws, wb, pipeline: Pipeline, extratos, resultado, esq: _InfoTabela, dire: _InfoTabela,
                col_esq, col_dir, pos_conc: dict, inicio: date, fim: date, feito_por: str,
                excluidos: dict | None = None, scripts: dict | None = None) -> None:
    ULT = 11  # coluna K
    _faixa(ws, "Conciliação Contábil  ·  WP-01 Saldos", pipeline.nome, ULT)
    larguras = {"B": 14, "C": 20, "D": 44, "E": 17, "F": 19, "G": 17, "H": 19, "I": 17, "J": 17, "K": 19}
    for letra, largura in larguras.items():
        ws.column_dimensions[letra].width = largura

    q = _q(ABA_CONCILIACAO)
    pc, uc, tc = pos_conc["primeira"], pos_conc["ultima"], pos_conc["total"]

    # Identificação ---------------------------------------------------------
    _par(ws, 5, 2, "Empresa", pipeline.empresa, mesclar_ate=4)
    _par(ws, 6, 2, "Nome WP", pipeline.nome, mesclar_ate=4)
    tol = _par(ws, 7, 2, "Tolerância", pipeline.tolerancia, _FMT_VALOR, fonte=_fonte(10, True, _AZUL_INPUT),
               fundo=_AMARELO)
    ws.cell(7, 2).comment = Comment("Valor do YAML. Pode alterar aqui: as fórmulas recalculam.", "Conciliação")
    tol.alignment = Alignment(horizontal="left")
    sinal = "+" if pipeline.operacao == "soma" else "−"
    if pipeline.id == "cobranca":
        regra = "Diferença = baixas do EM − financeiro, na soma do período. Saldo do CT não entra."
    else:
        regra = f"Diferença = saldo contábil {sinal} saldo financeiro (sinais opostos)"
    _par(ws, 8, 2, "Regra", regra, fonte=_fonte(10), mesclar_ate=4)
    if pipeline.id == "cobranca":
        abertura_ct = _saldo_anterior(extratos.get("sinqia_ct_balancete"))
        saldo_ct = _par(ws, 9, 2, "Saldo anterior CT", abertura_ct, "#,##0.00", fonte=_fonte(10, True))
        saldo_ct.comment = Comment(
            "Veio do balancete do Sinqia CT. Só visualização: não entra na diferença.",
            "Conciliação",
        )

    _par(ws, 5, 8, "Data inicial", inicio, _FMT_DATA, mesclar_ate=10)
    _par(ws, 6, 8, "Data final", fim, _FMT_DATA, mesclar_ate=10)
    _par(ws, 7, 8, "Feito por", feito_por or "", mesclar_ate=10)
    _par(ws, 8, 8, "Gerado em", datetime.now().replace(microsecond=0), _FMT_DATA_HORA, mesclar_ate=10,
         fonte=_fonte(10))
    _par(ws, 9, 8, "Revisado por", None, mesclar_ate=10, fonte=_INPUT, fundo=_AMARELO)

    # Indicadores -----------------------------------------------------------
    if pipeline.id == "cobranca":
        blocos = [
            ("B", "C", "Dias com movimento", f"=COUNT({q}!$B${pc}:$B${uc})", "0"),
            ("D", "H", "Diferença do período", "=$I$17", "#,##0.00"),
            ("I", "K", "Status geral", '=IF(ABS($I$17)<=Tolerancia,"CONCILIADO","COM DIFERENÇA")', "@"),
        ]
    else:
        blocos = [
            ("B", "C", "Dias com movimento", f"=COUNT({q}!$B${pc}:$B${uc})", "0"),
            ("D", "D", "Dias conciliados", f'=COUNTIF({q}!$F${pc}:$F${uc},"CONCILIADO")', "0"),
            ("E", "F", "Dias com diferença", f'=COUNTIF({q}!$F${pc}:$F${uc},"DIFERENÇA")', "0"),
            ("G", "H", "Maior diferença em um dia", f"=MAX(MAX({q}!$E${pc}:$E${uc}),-MIN({q}!$E${pc}:$E${uc}))", _FMT_VALOR),
            ("I", "K", "Status geral", '=IF(ABS($I$17)<=Tolerancia,"CONCILIADO","COM DIFERENÇA")', "@"),
        ]
    ws.row_dimensions[11].height = 18
    ws.row_dimensions[12].height = 34
    for ini, fim_col, rotulo, formula, formato in blocos:
        for linha in (11, 12):
            for col in range(ws[f"{ini}1"].column, ws[f"{fim_col}1"].column + 1):
                ws.cell(linha, col).fill = _fundo(_CINZA)
        if ini != fim_col:
            ws.merge_cells(f"{ini}11:{fim_col}11")
            ws.merge_cells(f"{ini}12:{fim_col}12")
        r = ws[f"{ini}11"]
        r.value = rotulo
        r.font = _fonte(8, True, _CINZA_TEXTO)
        r.alignment = Alignment(horizontal="center", vertical="bottom")
        v = ws[f"{ini}12"]
        v.value = formula
        v.font = _fonte(16, True, _NAVY)
        v.number_format = formato
        v.alignment = Alignment(horizontal="center", vertical="center")
    _status_cf(ws, "I12:K12")
    if pipeline.id != "cobranca":
        ws.conditional_formatting.add(
            "E12:F12", CellIsRule(operator="greaterThan", formula=["0"], font=Font(name=_FONTE, size=16, bold=True, color=_VERMELHO_TEXTO))
        )

    # Saldos das contas -----------------------------------------------------
    _secao(ws, 14, "Saldos das contas", ULT)
    if pipeline.id == "cobranca":
        cabecalhos = (("B15:C16", "Conta"), ("D15:D16", "Descrição"), ("E15:F15", "BAIXAS EM"),
                      ("G15:H15", "FINANCEIRO"), ("I15:I16", "DIFERENÇA"),
                      ("J15:J16", "COMPOSIÇÃO"), ("K15:K16", "NÃO EXPLICADO"))
    else:
        cabecalhos = (("B15:C16", "Conta contábil"), ("D15:D16", "Descrição"), ("E15:F15", "SALDO CONTÁBIL"),
                      ("G15:H15", "SALDO FINANCEIRO"), ("I15:I16", "DIFERENÇA"),
                      ("J15:J16", "COMPOSIÇÃO"), ("K15:K16", "NÃO EXPLICADO"))
    for faixa, texto in cabecalhos:
        ws.merge_cells(faixa)
        inicio_celula = faixa.split(":")[0]
        ws[inicio_celula].value = texto
    for linha in (15, 16):
        for col in range(2, ULT + 1):
            _cabecalho(ws.cell(linha, col))
    for celula, formula in (("E16", "=DataInicial"), ("F16", "=DataFinal"), ("G16", "=DataInicial"), ("H16", "=DataFinal")):
        ws[celula].value = formula
        ws[celula].number_format = _FMT_DATA
        ws[celula].fill = _fundo(_AZUL)
    ws.row_dimensions[15].height = 20
    ws.row_dimensions[16].height = 18

    nome_esq = _nome_conta(extratos.get(esq.fonte.alias, pd.DataFrame()), _codigo_conta(esq.fonte))
    nome_dir = _nome_conta(extratos.get(dire.fonte.alias, pd.DataFrame()), _codigo_conta(dire.fonte))
    for linha, info, nome in ((17, esq, nome_esq), (18, dire, nome_dir)):
        a = ws.cell(linha, 2, _sistema(info.fonte.aba))
        a.font = _fonte(10, True)
        a.hyperlink = f"#{_q(info.fonte.aba[:31])}!A1"
        b = ws.cell(linha, 3, _codigo_conta(info.fonte))
        b.number_format = "@"
        b.font = _TEXTO
        b.alignment = Alignment(horizontal="left")
        c = ws.cell(linha, 4, nome.upper() if nome else "")
        c.font = _TEXTO
        ws.row_dimensions[linha].height = 20
        for col in range(2, 5):
            ws.cell(linha, col).alignment = Alignment(vertical="center", horizontal="left")
            ws.cell(linha, col).border = Border(bottom=_FINO, left=_FINO if col == 2 else None)

    # saldo inicial e final (fórmula): contábil = esquerda, financeiro = direita
    abertura = _saldo_anterior(extratos.get("sinqia_ct_balancete")) if esq.fonte.script == "sinqia_ct_razao" else 0
    ws["E17"] = abertura
    ws["G18"] = 0
    ws["F17"] = f"=E17+SUM({esq.ref(col_esq('valor'))})"
    ws["H18"] = f"=G18+SUM({dire.ref(col_dir('valor'))})"
    ws["E18"] = None
    ws["F18"] = None
    ws["G17"] = None
    ws["H17"] = None
    for celula in ("E17", "G18"):
        ws[celula].font = _INPUT
        ws[celula].fill = _fundo(_AMARELO)
    ws["E17"].comment = Comment(
        "Saldo anterior da Sinqia, antes da data inicial. Veio do balancete; pode ajustar."
        if pipeline.id != "cobranca"
        else "Saldo inicial das baixas. O saldo anterior do CT está acima e não entra aqui.",
        "Conciliação",
    )
    ws["G18"].comment = Comment(
        "Saldo inicial financeiro (digitado). Transitória: normalmente zero.", "Conciliação"
    )
    operador = "+" if pipeline.operacao == "soma" else "-"
    ws["I17"] = f"=ROUND(F17{operador}H18,2)"
    ws["J17"] = f"={q}!$E${tc}"
    ws["K17"] = "=ROUND(I17-J17,2)"
    for col in range(5, ULT + 1):
        for linha in (17, 18):
            celula = ws.cell(linha, col)
            celula.number_format = _FMT_VALOR
            celula.alignment = Alignment(vertical="center")
            celula.border = Border(bottom=_FINO, right=_FINO if col in (6, 8, ULT) else None)
            if (col, linha) not in ((5, 17), (7, 18)):
                celula.font = _fonte(10, col >= 9)
    if pipeline.id == "cobranca":
        for coord in ("I17", "J17", "K17"):
            ws[coord].number_format = "#,##0.00"
    ws["J18"] = "soma da aba Conciliação"
    ws["J18"].font = _fonte(8, cor=_CINZA_TEXTO, italico=True)
    ws["K18"] = "deve ficar zerado"
    ws["K18"].font = _fonte(8, cor=_CINZA_TEXTO, italico=True)
    ws["I18"] = "baixas − financeiro" if pipeline.id == "cobranca" else "saldo contábil " + sinal + " financeiro"
    ws["I18"].font = _fonte(8, cor=_CINZA_TEXTO, italico=True)
    ws["J17"].hyperlink = f"#{q}!B{_CONC_CAB}"
    ws.conditional_formatting.add(
        "K17", FormulaRule(formula=["ABS(K17)>Tolerancia"], fill=_fundo(_VERMELHO_FUNDO),
                           font=Font(name=_FONTE, bold=True, color=_VERMELHO_TEXTO)))
    ws.conditional_formatting.add(
        "K17", FormulaRule(formula=["ABS(K17)<=Tolerancia"], fill=_fundo(_VERDE_FUNDO),
                           font=Font(name=_FONTE, bold=True, color=_VERDE_TEXTO)))
    ws.conditional_formatting.add(
        "I17", FormulaRule(formula=["ABS(I17)>Tolerancia"], font=Font(name=_FONTE, bold=True, color=_VERMELHO_TEXTO)))

    # Conclusão -------------------------------------------------------------
    _secao(ws, 20, "Conclusão", ULT)
    ws.merge_cells("B21:K22")
    if pipeline.id == "cobranca":
        texto_ok = "As baixas do EM e o financeiro estão conciliados no período, dentro da tolerância definida."
        texto_dif = "As baixas e o financeiro apresentam diferença, integralmente composta pelos dias listados na aba Conciliação. Justificar cada dia na coluna JUSTIFICATIVA."
    else:
        texto_ok = "Os saldos contábil e financeiro estão conciliados no período, dentro da tolerância definida."
        texto_dif = "Os saldos apresentam diferença, integralmente composta pelos dias listados na aba Conciliação. Justificar cada dia na coluna JUSTIFICATIVA."
    ws["B21"] = (
        '=IF(ABS(K17)>Tolerancia,'
        '"Atenção: a soma das diferenças diárias não bate com a diferença entre os saldos. Revisar as abas de dados antes de concluir.",'
        f'IF(ABS(I17)<=Tolerancia,"{texto_ok}","{texto_dif}"))'
    )
    ws["B21"].font = _fonte(10)
    ws["B21"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[21].height = 18
    ws.row_dimensions[22].height = 18

    _secao(ws, 24, "Nota geral", ULT)
    ws.merge_cells("B25:K27")
    ws["B25"].fill = _fundo(_AMARELO)
    ws["B25"].font = _INPUT
    ws["B25"].alignment = Alignment(wrap_text=True, vertical="top")
    ws["B25"].comment = Comment("Espaço para o comentário de quem fez ou revisou.", "Conciliação")

    # Critérios -------------------------------------------------------------
    _secao(ws, 29, "Critérios de análise", ULT)
    criterios = [
        (( "Baixas do EM" if pipeline.id == "cobranca" else "Saldo contábil")
         + f": soma da coluna {col_esq('valor')} da aba {esq.fonte.aba}. {esq.fonte.resumo}".strip()
         + (f" Filtro: {_texto_contra(esq.fonte, extratos.get(esq.fonte.alias, pd.DataFrame()))}."
            if _texto_contra(esq.fonte, extratos.get(esq.fonte.alias, pd.DataFrame())) else "")),
        (("Financeiro" if pipeline.id == "cobranca" else "Saldo financeiro")
         + f": soma da coluna {col_dir('valor')} da aba {dire.fonte.aba}. {dire.fonte.resumo}".strip()),
        f"Período {inicio:%d/%m/%Y} a {fim:%d/%m/%Y}. Os saldos somam todas as linhas da aba, sem considerar filtros.",
        (
            "Saldo anterior do CT fica na capa e na aba Balancete. Não entra na diferença."
            if pipeline.id == "cobranca"
            else f"Saldo do dia {inicio:%d/%m/%Y} (contábil): saldo anterior da Sinqia, débito menos crédito antes dessa data. "
            "O saldo final contábil soma esse valor ao movimento do período."
        ),
        (
            "Diferença do dia = baixas − financeiro, agrupado pela data de cada lado. "
            if pipeline.id == "cobranca"
            else f"Diferença do dia = contábil {sinal} financeiro, agrupado pela data de cada lado. "
        ) + "Dia conciliado quando o valor absoluto da diferença fica dentro da tolerância.",
        "Composição = soma das diferenças diárias. NÃO EXPLICADO = diferença dos saldos − composição; "
        "fora de zero, há linha sem data ou saldo inicial a justificar.",
    ]
    for info in (esq, dire):
        fora = (excluidos or {}).get(info.fonte.alias)
        if fora is None or not len(fora):
            continue
        motivos = "; ".join(sorted(set(fora[COLUNA_MOTIVO].astype(str))))
        criterios.append(
            f"Excluídas da aba {info.fonte.aba} e da soma: {len(fora)} linha(s) ({motivos}). "
            f"As linhas e a regra ficam na aba {ABA_EXCLUSOES}."
        )
    linha = 30
    for numero, texto in enumerate(criterios, start=1):
        n = ws.cell(linha, 2, f"{numero}.")
        n.font = _fonte(10, True, _AZUL)
        n.alignment = Alignment(horizontal="right", vertical="top")
        ws.merge_cells(start_row=linha, start_column=3, end_row=linha, end_column=ULT)
        t = ws.cell(linha, 3, texto)
        t.font = _TEXTO
        t.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[linha].height = 27 if len(texto) > 175 else 15
        linha += 1

    # Índice ----------------------------------------------------------------
    linha += 1
    _secao(ws, linha, "Abas desta pasta", ULT)
    linha += 1
    indice = [(ABA_SALDOS, "Este resumo.")]
    indice += [(info.fonte.aba[:31], info.fonte.resumo or "Dados extraídos do banco.") for info in (esq, dire)]
    outras = [n for n in wb.sheetnames if n not in {ABA_SALDOS, esq.fonte.aba[:31], dire.fonte.aba[:31],
                                                     ABA_CONCILIACAO, ABA_COMPOSICAO, ABA_EXECUCAO,
                                                     ABA_EXCLUSOES}]
    for nome in outras:
        indice.append((nome, "Dados extraídos do banco."))
    indice.append((ABA_CONCILIACAO, "Diferença por dia calculada por fórmula (SUMIFS sobre as abas de dados)."))
    if ABA_COMPOSICAO in wb.sheetnames:
        indice.append((ABA_COMPOSICAO, "Detalhe das linhas nos dias com diferença."))
    if ABA_EXCLUSOES in wb.sheetnames:
        indice.append((ABA_EXCLUSOES, "Linhas tiradas da conciliação por regra (ex.: implantação de saldos), com o motivo."))
    indice.append((ABA_EXECUCAO, "Scripts, parâmetros, linhas extraídas e conferência motor x planilha."))
    for nome, descricao in indice:
        a = ws.cell(linha, 2, nome)
        a.hyperlink = f"#{_q(nome)}!A1"
        a.font = Font(name=_FONTE, size=10, color=_AZUL, underline="single")
        ws.merge_cells(start_row=linha, start_column=4, end_row=linha, end_column=ULT)
        d = ws.cell(linha, 4, descricao)
        d.font = _fonte(10, cor=_CINZA_TEXTO)
        linha += 1

    legenda = linha + 1
    ws.cell(legenda, 2).fill = _fundo(_AMARELO)
    ws.cell(legenda, 3, "Célula para preencher ou ajustar (texto azul)").font = _fonte(9, cor=_CINZA_TEXTO)
    ws.freeze_panes = None
    _retrato(ws)
    ws.page_setup.fitToHeight = 1
    ws.print_area = f"A1:K{legenda}"


# --------------------------------------------------------------------------- execução

def _aba_execucao(ws, pipeline: Pipeline, scripts, extratos, resultado: ResultadoConciliacao,
                  pos_conc: dict, feito_por: str, excluidos: dict | None = None) -> None:
    _faixa(ws, "Execução e trilha de auditoria", "Como esta pasta foi gerada", 7)
    _par(ws, 5, 2, "Gerado em", datetime.now().replace(microsecond=0), _FMT_DATA_HORA, mesclar_ate=4)
    _par(ws, 6, 2, "Feito por", feito_por or "", mesclar_ate=4)
    _par(ws, 7, 2, "Conciliação", f"{pipeline.id}  ({pipeline.caminho.name})", mesclar_ate=4)

    _secao(ws, 9, "Fontes extraídas", 7)
    cab = ["Aba", "Script", "Banco", "Descrição", "Linhas", "Parâmetros"]
    for indice, nome in enumerate(cab, start=2):
        _cabecalho(ws.cell(10, indice, nome))
    linha = 11
    for fonte in pipeline.fontes:
        script = scripts[fonte.script]
        pronta = fonte_pronta(fonte, script)
        parametros = ", ".join(f"{k}={v}" for k, v in fonte.parametros.items())
        valores = [fonte.aba, script.caminho.name, script.engine, script.descricao,
                   len(extratos.get(fonte.alias, pd.DataFrame())) if pronta else "pulado", parametros]
        for indice, valor in enumerate(valores, start=2):
            celula = ws.cell(linha, indice, valor)
            celula.font = _TEXTO
            celula.border = Border(bottom=_FINO)
            celula.alignment = Alignment(vertical="top", wrap_text=indice in (5, 7))
        ws.cell(linha, 6).number_format = "#,##0"
        linha += 1

    linha += 1
    _secao(ws, linha, "Regras de exclusão", 7)
    linha += 1
    for indice, nome in enumerate(["Aba", "Origem da regra", "Motivo", "Condição", "Linhas", "Valor excluído"], start=2):
        _cabecalho(ws.cell(linha, indice, nome))
    linha += 1
    alguma = False
    for fonte in pipeline.fontes:
        script = scripts[fonte.script]
        if not fonte_pronta(fonte, script):
            continue
        regras = regras_exclusao(fonte, script)
        origem = pipeline.caminho.name if fonte.excluir is not None else script.caminho.name
        fora = (excluidos or {}).get(fonte.alias)
        col_valor = script.coluna("valor")
        for regra in regras:
            alguma = True
            parte = fora[fora[COLUNA_MOTIVO] == regra.motivo] if fora is not None and len(fora) else None
            n = 0 if parte is None else len(parte)
            soma = 0.0
            if parte is not None and col_valor in parte.columns:
                soma = round(float(pd.to_numeric(parte[col_valor], errors="coerce").fillna(0).sum()), 2)
            for indice, valor in enumerate([fonte.aba, origem, regra.motivo, regra.texto(), n, soma], start=2):
                celula = ws.cell(linha, indice, valor)
                celula.font = _TEXTO
                celula.border = Border(bottom=_FINO)
                celula.alignment = Alignment(vertical="top", wrap_text=indice in (4, 5))
            ws.cell(linha, 6).number_format = "#,##0"
            ws.cell(linha, 7).number_format = _FMT_VALOR
            linha += 1
    if not alguma:
        ws.cell(linha, 2, "Nenhuma regra de exclusão nesta conciliação.").font = _fonte(9, italico=True, cor=_CINZA_TEXTO)
        linha += 1

    linha += 1
    _secao(ws, linha, "Conferência motor (Python) x fórmulas da planilha", 7)
    linha += 1
    for indice, nome in enumerate(["Item", "Motor", "Planilha", "Divergência"], start=2):
        _cabecalho(ws.cell(linha, indice, nome))
    linha += 1
    q = _q(ABA_CONCILIACAO)
    pc, uc, tc = pos_conc["primeira"], pos_conc["ultima"], pos_conc["total"]
    conferencias = [
        ("Diferença total", resultado.diferenca_total, f"={q}!$E${tc}", _FMT_VALOR),
        ("Dias conciliados", resultado.dias_ok, f'=COUNTIF({q}!$F${pc}:$F${uc},"CONCILIADO")', "0"),
        ("Dias com diferença", resultado.dias_diferenca, f'=COUNTIF({q}!$F${pc}:$F${uc},"DIFERENÇA")', "0"),
    ]
    inicio_conf = linha
    for item, motor, formula, formato in conferencias:
        ws.cell(linha, 2, item).font = _TEXTO
        m = ws.cell(linha, 3, motor)
        p = ws.cell(linha, 4, formula)
        d = ws.cell(linha, 5, f"=ROUND(C{linha}-D{linha},2)")
        for celula in (m, p, d):
            celula.number_format = formato
            celula.font = _TEXTO
        for col in range(2, 6):
            ws.cell(linha, col).border = Border(bottom=_FINO)
        linha += 1
    ws.cell(linha, 2, "Resultado").font = _fonte(10, True)
    r = ws.cell(linha, 3,
                f'=IF(SUMPRODUCT(ABS(E{inicio_conf}:E{linha - 1}))=0,"CONCILIADO","DIVERGENTE: conferir fórmulas")')
    ws.merge_cells(start_row=linha, start_column=3, end_row=linha, end_column=5)
    r.alignment = Alignment(horizontal="center")
    _status_cf(ws, f"C{linha}:E{linha}")
    nota = ws.cell(linha + 1, 2, "O motor calcula em Python na hora da extração; a planilha recalcula pelas fórmulas. "
                                 "Os dois precisam bater (CONCILIADO).")
    nota.font = _fonte(8, cor=_CINZA_TEXTO, italico=True)
    linha += 3

    if resultado.avisos:
        _secao(ws, linha, "Avisos do motor", 7)
        linha += 1
        for aviso in resultado.avisos:
            ws.cell(linha, 2, aviso).font = _fonte(10, cor=_VERMELHO_TEXTO)
            linha += 1

    for letra, largura in zip("BCDEFG", (20, 26, 14, 42, 12, 60)):
        ws.column_dimensions[letra].width = largura
    _retrato(ws)
