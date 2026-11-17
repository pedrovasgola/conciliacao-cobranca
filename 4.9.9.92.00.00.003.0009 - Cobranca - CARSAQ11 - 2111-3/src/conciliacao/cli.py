"""Linha de comando: listar, rodar uma conciliação ou rodar todas."""

from __future__ import annotations

import argparse
import calendar
import os
import re
import sys
from datetime import date, datetime
from pathlib import Path

from dotenv import load_dotenv

from conciliacao.caminhos import pasta_saida, pasta_trabalho, raiz
from conciliacao.catalogo import (
    Pipeline,
    Script,
    aplicar_conta_ct2,
    aplicar_conta_sinqia,
    carregar_pipelines,
    carregar_scripts,
    fonte_pronta,
    preparar_execucao,
)
from conciliacao.conciliar import conciliar, filtrar_composicao
from conciliacao.erros import ErroConciliacao
from conciliacao.extrair import extrair
from conciliacao.planilha import gravar
from conciliacao.seguranca import redigir, validar_sql


def _data(texto: str) -> date:
    try:
        return datetime.strptime(texto.strip(), "%d/%m/%Y").date()
    except ValueError as exc:
        raise ErroConciliacao(f"Data inválida '{texto}'. Use DD/MM/AAAA.") from exc


def _pedir_periodo(args) -> tuple[date, date]:
    if args.inicio or args.fim:
        if not (args.inicio and args.fim):
            raise ErroConciliacao("Informe --inicio e --fim juntos, no formato DD/MM/AAAA.")
        inicio, fim = _data(args.inicio), _data(args.fim)
        if fim < inicio:
            raise ErroConciliacao("A data final é menor que a inicial.")
        return inicio, fim
    if not sys.stdin.isatty():
        raise ErroConciliacao("Informe --inicio e --fim (DD/MM/AAAA).")
    hoje = date.today()
    padrao_ini = hoje.replace(day=1)
    padrao_fim = hoje.replace(day=calendar.monthrange(hoje.year, hoje.month)[1])
    print("Informe o período da conciliação:")
    while True:
        bruto_ini = input(f"  Data inicial (DD/MM/AAAA) [Enter = {padrao_ini:%d/%m/%Y}]: ").strip()
        bruto_fim = input(f"  Data final   (DD/MM/AAAA) [Enter = {padrao_fim:%d/%m/%Y}]: ").strip()
        inicio = _data(bruto_ini) if bruto_ini else padrao_ini
        fim = _data(bruto_fim) if bruto_fim else padrao_fim
        if fim >= inicio:
            return inicio, fim
        print("  A data final é menor que a inicial. Tente de novo.")


def _pedir_conta(args, atributo: str, fonte, rotulo: str, perguntar: bool) -> str | None:
    informada = getattr(args, atributo, None)
    if informada:
        return informada.strip()
    if fonte is None or not perguntar or not sys.stdin.isatty():
        return None
    padrao = fonte.parametros.get("conta", "")
    bruto = input(f"  {rotulo} [Enter = {padrao}]: ").strip()
    return bruto or None


def _carregar_env() -> None:
    vistos: list[Path] = []
    for pasta in (pasta_trabalho(), raiz()):
        arquivo = pasta / ".env"
        if arquivo in vistos:
            continue
        vistos.append(arquivo)
        if arquivo.is_file():
            load_dotenv(arquivo, override=False)
            return
    raise ErroConciliacao(
        "Não achei o arquivo .env. Copie .env.example para .env e preencha as credenciais."
    )


def _anomes(inicio: date, fim: date) -> str:
    if (inicio.year, inicio.month) == (fim.year, fim.month):
        return f"{inicio:%m.%Y}"
    return f"{inicio:%m.%Y} a {fim:%m.%Y}"


def _rotulo_arquivo(pipeline: Pipeline) -> str:
    """'1.8.8.65.99.00.002.0002 - Transações Pix Outras Instituições'."""
    pai = pasta_trabalho().resolve().parent.name.strip()
    if re.match(r"^\d", pai):
        return pai
    sinqia = pipeline.fonte_sinqia()
    digitos = re.sub(r"\D", "", (sinqia.parametros.get("conta", "") if sinqia else ""))
    nome = pipeline.nome.removeprefix("Conciliação - ").strip()
    if len(digitos) == 16 and nome:
        cortes = (1, 1, 1, 2, 2, 2, 3, 4)
        partes = []
        cursor = 0
        for tamanho in cortes:
            partes.append(digitos[cursor:cursor + tamanho])
            cursor += tamanho
        return f"{'.'.join(partes)} - {nome}"
    return nome or pipeline.id


def _destino(pipeline: Pipeline, inicio: date, fim: date, pasta: str | None) -> Path:
    base = Path(pasta) if pasta else pasta_saida()
    return base / f"{_rotulo_arquivo(pipeline)} - {_anomes(inicio, fim)}.xlsx"


def executar(pipeline: Pipeline, scripts: dict[str, Script], inicio: date, fim: date, pasta: str | None) -> Path:
    preparar_execucao(pipeline, scripts)
    destino = _destino(pipeline, inicio, fim, pasta)
    print(f"\n{pipeline.nome}")
    print(f"Pasta em uso: {pasta_trabalho()}")
    print(f"Excel: {destino}")
    print(f"Período {inicio:%d/%m/%Y} a {fim:%d/%m/%Y}")
    ct2 = pipeline.fonte_ct2()
    if ct2 is not None:
        print(f"Conta CT2 {ct2.parametros.get('conta', '')}")
    sinqia = pipeline.fonte_sinqia()
    if sinqia is not None:
        print(f"Conta Sinqia {sinqia.parametros.get('conta', '')}")
    extratos = extrair(pipeline, scripts, inicio, fim)
    for fonte in pipeline.fontes:
        script = scripts[fonte.script]
        if not fonte_pronta(fonte, script):
            print(f"  {fonte.alias:<28} {script.engine:<10} {'pulado':>8}  (cole o SELECT em {script.caminho.name})")
            continue
        print(f"  {fonte.alias:<28} {script.engine:<10} {len(extratos[fonte.alias]):>8,} linhas")

    esq = pipeline.fonte(pipeline.esquerda)
    dire = pipeline.fonte(pipeline.direita)
    script_esq = scripts[esq.script]
    script_dir = scripts[dire.script]
    colunas = [
        script_esq.coluna("data"),
        script_esq.coluna("valor"),
        script_dir.coluna("data"),
        script_dir.coluna("valor"),
    ]
    if any(coluna is None for coluna in colunas):
        raise ErroConciliacao(
            "Os scripts dos dois lados precisam de @coluna.data e @coluna.valor."
        )
    resultado = conciliar(
        extratos[pipeline.esquerda],
        extratos[pipeline.direita],
        colunas[0],
        colunas[1],
        colunas[2],
        colunas[3],
        esq.aba,
        dire.aba,
        pipeline.operacao,
        pipeline.tolerancia,
    )
    composicao = None
    if pipeline.detalhe:
        detalhe = pipeline.fonte(pipeline.detalhe)
        script_detalhe = scripts[detalhe.script]
        if fonte_pronta(detalhe, script_detalhe):
            coluna = script_detalhe.coluna("data")
            if coluna is None:
                raise ErroConciliacao(
                    f"{detalhe.script} entra na composição e precisa de @coluna.data."
                )
            composicao = filtrar_composicao(extratos[pipeline.detalhe], coluna, resultado)
        else:
            print(f"  Composição pulada: {script_detalhe.caminho.name} ainda sem o SELECT definitivo.")

    for aviso in resultado.avisos:
        print(f"  Aviso: {aviso}")
    total_dias = resultado.dias_ok + resultado.dias_diferenca
    print(
        f"  Diferença {resultado.diferenca_total:,.2f} | "
        f"{resultado.dias_ok} de {total_dias} dias conciliados"
    )
    gravar(
        destino,
        pipeline,
        scripts,
        extratos,
        resultado,
        composicao,
        inicio,
        fim,
        os.getenv("CONCILIACAO_FEITO_POR", "").strip(),
    )
    print(f"  Arquivo {destino}")
    return destino


def _listar() -> int:
    scripts = carregar_scripts()
    pipelines = carregar_pipelines(scripts)
    print(f"Scripts ({len(scripts)}) — pasta scripts")
    for script in scripts.values():
        if "PREENCHER_ORIGEM" in script.sql:
            estado = "falta colar o SELECT"
        else:
            try:
                validar_sql(script.sql, script.caminho.name)
                estado = "pronto"
            except ErroConciliacao:
                estado = "bloqueado"
        print(f"  {script.id:<24} {script.engine:<10} {estado:<22} {script.descricao}")
    print(f"\nConciliações ({len(pipelines)}) — pasta pipelines")
    if not pipelines:
        print("  Nenhuma. Crie um arquivo .yaml em pipelines.")
    for pipeline in pipelines.values():
        estado = "ativa" if pipeline.ativo else "inativa"
        usados = ", ".join(fonte.script for fonte in pipeline.fontes)
        print(f"  {pipeline.id} ({estado})")
        print(f"    {pipeline.nome}")
        print(f"    scripts: {usados}")
    return 0


def _rodar(args, todas: bool) -> int:
    scripts = carregar_scripts()
    pipelines = carregar_pipelines(scripts)
    inicio, fim = _pedir_periodo(args)
    _carregar_env()
    if todas:
        alvos = [item for item in pipelines.values() if item.ativo]
        if not alvos:
            raise ErroConciliacao("Não há conciliação ativa em pipelines.")
    else:
        if args.pipeline not in pipelines:
            nomes = ", ".join(pipelines) or "(nenhuma)"
            raise ErroConciliacao(f"Conciliação '{args.pipeline}' não existe. Disponíveis: {nomes}")
        pipeline = pipelines[args.pipeline]
        if not pipeline.ativo:
            raise ErroConciliacao(f"{pipeline.id} está inativa. Mude ativo para true no YAML para rodar.")
        alvos = [pipeline]

    falhas: list[str] = []
    for pipeline in alvos:
        try:
            perguntar = not todas
            conta_ct2 = _pedir_conta(args, "conta", pipeline.fonte_ct2(), "Conta CT2 (Protheus)", perguntar)
            conta_sinqia = _pedir_conta(
                args, "conta_sinqia", pipeline.fonte_sinqia(), "Conta Sinqia", perguntar
            )
            pipeline = aplicar_conta_ct2(pipeline, conta_ct2)
            pipeline = aplicar_conta_sinqia(pipeline, conta_sinqia)
            executar(pipeline, scripts, inicio, fim, args.saida)
        except ErroConciliacao as exc:
            if not todas:
                raise
            falhas.append(f"{pipeline.id}: {exc}")
            print(f"  Falhou: {exc}")
    if falhas:
        print("\nConciliações que não fecharam:")
        for falha in falhas:
            print(f"  {falha}")
        return 1
    return 0


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="conciliacao",
        description="Concilia bases a partir dos scripts SQL da pasta scripts.",
    )
    sub = parser.add_subparsers(dest="comando", required=True)

    sub.add_parser("listar", aliases=["list"], help="Mostra os scripts e as conciliações.")

    rodar = sub.add_parser("rodar", aliases=["run"], help="Roda uma conciliação.")
    rodar.add_argument("pipeline", help="Id do arquivo em pipelines.")
    rodar.add_argument("--inicio", help="Data inicial DD/MM/AAAA.")
    rodar.add_argument("--fim", help="Data final DD/MM/AAAA.")
    rodar.add_argument("--conta", help="Conta do Protheus (CT2).")
    rodar.add_argument("--conta-sinqia", dest="conta_sinqia", help="Conta da Sinqia (CT e CC).")
    rodar.add_argument("--saida", help="Pasta do Excel. Padrão: saida.")

    todas = sub.add_parser("rodar-todas", aliases=["run-all"], help="Roda toda conciliação ativa.")
    todas.add_argument("--inicio", help="Data inicial DD/MM/AAAA.")
    todas.add_argument("--fim", help="Data final DD/MM/AAAA.")
    todas.add_argument("--conta", help="Conta do Protheus (CT2).")
    todas.add_argument("--conta-sinqia", dest="conta_sinqia", help="Conta da Sinqia (CT e CC).")
    todas.add_argument("--saida", help="Pasta do Excel. Padrão: saida.")
    return parser


def main(argv: list[str] | None = None) -> None:
    parser = _parser()
    args = parser.parse_args(argv)
    try:
        if args.comando in {"listar", "list"}:
            codigo = _listar()
        elif args.comando in {"rodar", "run"}:
            codigo = _rodar(args, todas=False)
        elif args.comando in {"rodar-todas", "run-all"}:
            codigo = _rodar(args, todas=True)
        else:
            parser.error(f"Comando desconhecido: {args.comando}")
            codigo = 2
    except ErroConciliacao as exc:
        print(f"Erro: {redigir(str(exc))}", file=sys.stderr)
        codigo = 1
    except Exception as exc:
        print(f"Erro inesperado: {redigir(str(exc))}", file=sys.stderr)
        codigo = 1
    raise SystemExit(codigo)
