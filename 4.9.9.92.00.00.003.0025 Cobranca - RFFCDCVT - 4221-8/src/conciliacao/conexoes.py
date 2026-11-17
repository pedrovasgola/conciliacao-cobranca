"""Conexões de leitura. A senha fica só no .env e não entra em log nem no Excel."""

from __future__ import annotations

import os
from urllib.parse import quote_plus

import pandas as pd
from sqlalchemy import create_engine, text
from sqlalchemy.engine import Engine

from conciliacao.erros import ErroConciliacao
from conciliacao.seguranca import redigir, validar_sql


def _env(nome: str, padrao: str | None = None) -> str:
    valor = os.getenv(nome) or padrao
    if valor is None or not str(valor).strip():
        raise ErroConciliacao(f"Variável {nome} não está no .env.")
    return str(valor).strip().strip('"')


def _env_opcional(nome: str, padrao: str = "") -> str:
    valor = os.getenv(nome)
    if valor is None or not str(valor).strip():
        return padrao
    return str(valor).strip().strip('"')


def criar_engine(banco: str) -> Engine:
    try:
        if banco == "oracle":
            url = (
                f"oracle+oracledb://{quote_plus(_env('ORA_USER'))}:{quote_plus(_env('ORA_PASS'))}"
                f"@{_env('ORA_HOST')}:{_env('ORA_PORT', '1521')}/?service_name={_env('ORA_SERVICE')}"
            )
            return create_engine(url, pool_pre_ping=True)
        if banco == "sqlserver":
            try:
                import pyodbc  # noqa: F401
            except ImportError as exc:
                raise ErroConciliacao(
                    "Falta o pyodbc. Na pasta do projeto rode: uv sync"
                ) from exc
            trust = _env_opcional("MSSQL_TRUST_SERVER_CERTIFICATE", "yes") or "yes"
            # O SELECT já traz SFINANCE_CT.dbo.... Sem MSSQL_DB a sessão abre em master.
            banco = _env_opcional("MSSQL_DB", "master") or "master"
            odbc = (
                f"DRIVER={{{_env_opcional('MSSQL_DRIVER', 'SQL Server') or 'SQL Server'}}};"
                f"SERVER={_env('MSSQL_HOST')};DATABASE={banco};"
                f"UID={_env('MSSQL_USER')};PWD={_env('MSSQL_PASS')};"
                f"TrustServerCertificate={trust};"
            )
            return create_engine(
                f"mssql+pyodbc:///?odbc_connect={quote_plus(odbc)}",
                pool_pre_ping=True,
            )
        if banco == "postgres":
            url = (
                f"postgresql+psycopg2://{quote_plus(_env('PG_USER'))}:{quote_plus(_env('PG_PASS'))}"
                f"@{_env('PG_HOST')}:{_env('PG_PORT', '5432')}/{_env('PG_DB')}"
            )
            return create_engine(url, pool_pre_ping=True)
    except ErroConciliacao:
        raise
    except Exception as exc:
        raise ErroConciliacao(f"Não foi possível abrir o {banco}: {redigir(str(exc))}") from None
    raise ErroConciliacao(f"Banco '{banco}' não é suportado.")


def _preencher_char(params: dict, tamanhos: dict[str, int] | None) -> dict:
    """Completa o texto até o DATA_LENGTH. O driver manda VARCHAR2 curto e o CHAR não casa."""
    if not tamanhos:
        return params
    preenchido = dict(params)
    for nome, tamanho in tamanhos.items():
        valor = preenchido.get(nome)
        if valor is None:
            continue
        texto = str(valor)
        if len(texto) > tamanho:
            raise ErroConciliacao(
                f"Parâmetro {nome} tem {len(texto)} caracteres e a coluna tem {tamanho}."
            )
        preenchido[nome] = texto.ljust(tamanho)
    return preenchido


def ler(
    engine: Engine,
    sql: str,
    params: dict,
    origem: str,
    tamanhos: dict[str, int] | None = None,
) -> pd.DataFrame:
    validar_sql(sql, origem)
    params = _preencher_char(params, tamanhos)
    try:
        with engine.connect() as conn:
            df = pd.read_sql(text(sql), conn, params=params)
    except ErroConciliacao:
        raise
    except Exception as exc:
        raise ErroConciliacao(f"{origem}: falha ao ler o banco. {redigir(str(exc))}") from None
    df.columns = [str(coluna).strip().upper() for coluna in df.columns]
    return df
