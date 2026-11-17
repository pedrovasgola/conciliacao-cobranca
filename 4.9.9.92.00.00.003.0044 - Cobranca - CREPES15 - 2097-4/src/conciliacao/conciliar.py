"""Compara duas bases pelo dia (ou pela data que o script declarar)."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime

import pandas as pd

from conciliacao.erros import ErroConciliacao

_EXCEL_SERIAL_MIN = 20000  # ~1954
_EXCEL_SERIAL_MAX = 80000  # ~2119


@dataclass
class ResultadoConciliacao:
    por_dia: pd.DataFrame
    diferenca_total: float
    dias_ok: int
    dias_diferenca: int
    avisos: list[str] = field(default_factory=list)
    rotulo_esquerda: str = ""
    rotulo_direita: str = ""
    operacao: str = "soma"


def _como_timestamp_naive(valor) -> pd.Timestamp | pd.NaT:
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return pd.NaT
    if isinstance(valor, pd.Timestamp):
        if valor.tzinfo is not None:
            valor = valor.replace(tzinfo=None)
        return pd.Timestamp(valor.year, valor.month, valor.day)
    if isinstance(valor, datetime):
        if valor.tzinfo is not None:
            valor = valor.replace(tzinfo=None)
        return pd.Timestamp(valor.year, valor.month, valor.day)
    if isinstance(valor, date):
        return pd.Timestamp(valor.year, valor.month, valor.day)
    return pd.NaT


def normalizar_data(serie: pd.Series) -> pd.Series:
    """Devolve datetime64[ns] na meia-noite, sem fuso, para os dois lados agruparem no mesmo dia."""
    if pd.api.types.is_datetime64_any_dtype(serie):
        convertido = serie
        if isinstance(convertido.dtype, pd.DatetimeTZDtype):
            convertido = convertido.dt.tz_localize(None)
        return convertido.dt.normalize()

    saida = pd.Series(pd.NaT, index=serie.index, dtype="datetime64[ns]")

    datas_nativas = serie.map(_como_timestamp_naive)
    nativas_ok = datas_nativas.notna()
    saida = saida.where(~nativas_ok, datas_nativas)

    texto = serie.astype("string").str.strip().str.replace(r"\.0$", "", regex=True)
    oito = texto.str.fullmatch(r"\d{8}").fillna(False).astype(bool) & saida.isna()
    if oito.any():
        saida = saida.where(~oito, pd.to_datetime(texto.where(oito), format="%Y%m%d", errors="coerce"))

    resto = saida.isna()
    if resto.any():
        saida = saida.where(~resto, pd.to_datetime(texto.where(resto), errors="coerce"))

    numerico = pd.to_numeric(serie, errors="coerce")
    serial = (
        saida.isna()
        & numerico.notna()
        & numerico.between(_EXCEL_SERIAL_MIN, _EXCEL_SERIAL_MAX)
        & ~numerico.astype("string").str.fullmatch(r"\d{8}").fillna(False)
    )
    if serial.any():
        saida = saida.where(
            ~serial,
            pd.to_datetime(numerico.where(serial), unit="D", origin="1899-12-30", errors="coerce"),
        )

    if isinstance(saida.dtype, pd.DatetimeTZDtype):
        saida = saida.dt.tz_localize(None)
    return saida.dt.normalize()


def _indice_dia(somas: pd.Series) -> pd.Series:
    if somas.empty:
        somas.index = pd.DatetimeIndex([], tz=None)
        return somas
    indice = pd.DatetimeIndex(pd.to_datetime(somas.index, errors="coerce")).tz_localize(None).normalize()
    somas = pd.Series(somas.to_numpy(), index=indice, name=somas.name)
    somas = somas[somas.index.notna()]
    return somas.groupby(level=0).sum().sort_index()


def _somar(df: pd.DataFrame, col_data: str, col_valor: str, origem: str) -> tuple[pd.Series, list[str]]:
    avisos: list[str] = []
    if col_data not in df.columns or col_valor not in df.columns:
        faltando = [c for c in (col_data, col_valor) if c not in df.columns]
        raise ErroConciliacao(
            f"{origem}: não achei a coluna {faltando}. Colunas que vieram: {list(df.columns)}."
        )
    if df.empty:
        vazio = pd.Series(dtype="float64")
        vazio.index = pd.DatetimeIndex([], tz=None)
        return vazio, avisos
    base = pd.DataFrame(
        {
            "DATA": normalizar_data(df[col_data]),
            "VALOR": pd.to_numeric(df[col_valor], errors="coerce"),
        }
    )
    sem_data = int(base["DATA"].isna().sum())
    sem_valor = int(base["VALOR"].isna().sum())
    if sem_data:
        avisos.append(f"{origem}: {sem_data} linha(s) sem data, ignoradas na soma do dia.")
    if sem_valor:
        avisos.append(f"{origem}: {sem_valor} linha(s) com valor inválido, tratadas como zero.")
    base["VALOR"] = base["VALOR"].fillna(0)
    somas = base.dropna(subset=["DATA"]).groupby("DATA", sort=True)["VALOR"].sum().round(2)
    return _indice_dia(somas), avisos


def conciliar(
    esquerda: pd.DataFrame,
    direita: pd.DataFrame,
    col_data_esquerda: str,
    col_valor_esquerda: str,
    col_data_direita: str,
    col_valor_direita: str,
    rotulo_esquerda: str,
    rotulo_direita: str,
    operacao: str,
    tolerancia: float,
) -> ResultadoConciliacao:
    if operacao not in {"soma", "subtracao"}:
        raise ErroConciliacao(f"Operação '{operacao}' inválida.")
    serie_esq, avisos_esq = _somar(esquerda, col_data_esquerda, col_valor_esquerda, rotulo_esquerda)
    serie_dir, avisos_dir = _somar(direita, col_data_direita, col_valor_direita, rotulo_direita)
    junto = pd.concat(
        [serie_esq.rename("ESQUERDA"), serie_dir.rename("DIREITA")],
        axis=1,
    ).fillna(0)
    junto.index = pd.DatetimeIndex(pd.to_datetime(junto.index, errors="coerce")).tz_localize(None).normalize()
    junto = junto[junto.index.notna()].sort_index()
    if operacao == "soma":
        diferenca = (junto["ESQUERDA"] + junto["DIREITA"]).round(2)
    else:
        diferenca = (junto["ESQUERDA"] - junto["DIREITA"]).round(2)
    status = diferenca.abs().le(tolerancia).map({True: "CONCILIADO", False: "DIFERENÇA"})
    por_dia = pd.DataFrame(
        {
            "DATA": [ts.date() for ts in junto.index],
            rotulo_esquerda: junto["ESQUERDA"].round(2).to_numpy(),
            rotulo_direita: junto["DIREITA"].round(2).to_numpy(),
            "DIFERENCA": diferenca.to_numpy(),
            "STATUS": status.to_numpy(),
        }
    )
    dias_ok = int((por_dia["STATUS"] == "CONCILIADO").sum())
    return ResultadoConciliacao(
        por_dia=por_dia,
        diferenca_total=float(diferenca.sum().round(2)) if len(diferenca) else 0.0,
        dias_ok=dias_ok,
        dias_diferenca=int(len(por_dia) - dias_ok),
        avisos=avisos_esq + avisos_dir,
        rotulo_esquerda=rotulo_esquerda,
        rotulo_direita=rotulo_direita,
        operacao=operacao,
    )


def filtrar_composicao(detalhe: pd.DataFrame, col_data: str, resultado: ResultadoConciliacao) -> pd.DataFrame:
    """Linhas do script de detalhe nos dias que ficaram fora da tolerância."""
    if detalhe.empty or resultado.por_dia.empty:
        return detalhe.iloc[0:0].copy()
    if col_data not in detalhe.columns:
        raise ErroConciliacao(
            f"O detalhe não tem a coluna {col_data}. Colunas: {list(detalhe.columns)}."
        )
    divergentes = pd.to_datetime(resultado.por_dia.loc[resultado.por_dia["STATUS"] != "CONCILIADO", "DATA"])
    divergentes = pd.DatetimeIndex(divergentes).tz_localize(None).normalize()
    marcas = normalizar_data(detalhe[col_data])
    return detalhe.loc[marcas.isin(set(divergentes))].copy()
