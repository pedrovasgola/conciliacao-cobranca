"""Executa os scripts de um pipeline e devolve um quadro por fonte."""

from __future__ import annotations

import re
from datetime import date, datetime

import pandas as pd

from conciliacao.catalogo import Fonte, Pipeline, Script, fonte_pronta
from conciliacao.conciliar import normalizar_data
from conciliacao.conexoes import criar_engine, ler
from conciliacao.erros import ErroConciliacao
from conciliacao.seguranca import parametros_no_sql


def _forcar_coluna_data(quadro: pd.DataFrame, script: Script) -> pd.DataFrame:
    coluna = script.coluna("data")
    if not coluna or coluna not in quadro.columns or quadro.empty:
        return quadro
    saida = quadro.copy()
    marcas = normalizar_data(saida[coluna])
    saida[coluna] = marcas.dt.date
    saida.loc[marcas.isna(), coluna] = None
    return saida


def converter(valor, tipo: str, nome: str):
    if tipo == "texto":
        if isinstance(valor, datetime):
            return valor.strftime("%Y%m%d")
        if isinstance(valor, date):
            return valor.strftime("%Y%m%d")
        return str(valor)
    if tipo == "inteiro":
        if isinstance(valor, datetime):
            return int(valor.strftime("%Y%m%d"))
        if isinstance(valor, date):
            return int(valor.strftime("%Y%m%d"))
        digitos = re.sub(r"\D", "", str(valor))
        if not digitos:
            raise ErroConciliacao(f"Parâmetro {nome} não é um inteiro.")
        return int(digitos)
    if tipo == "data":
        if isinstance(valor, datetime):
            return valor.date()
        if isinstance(valor, date):
            return valor
        for formato in ("%Y-%m-%d", "%d/%m/%Y", "%Y%m%d"):
            try:
                return datetime.strptime(str(valor), formato).date()
            except ValueError:
                continue
        raise ErroConciliacao(f"Parâmetro {nome} não é uma data.")
    raise ErroConciliacao(f"Tipo '{tipo}' do parâmetro {nome} é inválido.")


def montar_parametros(script: Script, fonte: Fonte, inicio: date, fim: date) -> dict:
    informados = dict(fonte.parametros)
    informados["data_ini"] = inicio
    informados["data_fim"] = fim
    saida = {}
    for nome in parametros_no_sql(script.sql):
        if nome not in script.parametros:
            raise ErroConciliacao(
                f"{script.caminho.name}: declare @param para {nome} no cabeçalho."
            )
        if nome not in informados:
            raise ErroConciliacao(f"Fonte '{fonte.alias}' não informa o parâmetro {nome}.")
        saida[nome] = converter(informados[nome], script.parametros[nome].tipo, nome)
    return saida


def extrair(
    pipeline: Pipeline,
    scripts: dict[str, Script],
    inicio: date,
    fim: date,
) -> dict[str, pd.DataFrame]:
    if fim < inicio:
        raise ErroConciliacao("A data final é menor que a inicial.")
    motores = {}
    quadros: dict[str, pd.DataFrame] = {}
    try:
        for fonte in pipeline.fontes:
            script = scripts[fonte.script]
            if not fonte_pronta(fonte, script):
                print(f"  {fonte.alias:<28} pulado (SELECT ainda não colado)", flush=True)
                quadros[fonte.alias] = pd.DataFrame()
                continue
            print(f"  Lendo {fonte.alias} ({script.engine})...", flush=True)
            if script.engine not in motores:
                print(f"    conectando {script.engine}...", flush=True)
                motores[script.engine] = criar_engine(script.engine)
            params = montar_parametros(script, fonte, inicio, fim)
            tamanhos = {
                nome: item.tamanho
                for nome, item in script.parametros.items()
                if item.tamanho and nome in params
            }
            quadro = ler(motores[script.engine], script.sql, params, script.caminho.name, tamanhos)
            quadro = _forcar_coluna_data(quadro, script)
            if len(quadro) > pipeline.limite_linhas:
                raise ErroConciliacao(
                    f"{script.caminho.name} devolveu {len(quadro):,} linhas, "
                    f"acima do limite de {pipeline.limite_linhas:,}. Ajuste o WHERE do script."
                )
            quadros[fonte.alias] = quadro
    finally:
        for motor in motores.values():
            motor.dispose()
    return quadros
