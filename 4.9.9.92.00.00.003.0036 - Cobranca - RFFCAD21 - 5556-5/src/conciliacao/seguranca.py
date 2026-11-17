"""Regras para o motor só ler dados e não vazar segredo."""

from __future__ import annotations

import re

from conciliacao.erros import ErroConciliacao

# Palavras que mudam dado, estrutura ou saem do SELECT.
_PROIBIDO = re.compile(
    r"\b("
    r"INSERT|UPDATE|DELETE|MERGE|DROP|ALTER|CREATE|TRUNCATE|GRANT|REVOKE|"
    r"EXEC|EXECUTE|CALL|INTO|BACKUP|RESTORE|BULK|OPENROWSET|OPENDATASOURCE|"
    r"DECLARE|BEGIN|COMMIT|ROLLBACK|DBMS_\w+|UTL_\w+|XP_\w+|SP_\w+"
    r")\b",
    re.IGNORECASE,
)
_FOR_UPDATE = re.compile(r"\bFOR\s+UPDATE\b", re.IGNORECASE)
_PRIMEIRA = re.compile(r"^\s*(\w+)", re.IGNORECASE)


def redigir(texto: str) -> str:
    """Tira senha de mensagem de erro de conexão."""
    texto = re.sub(r"(?i)(PWD=|PASSWORD=)[^;'\s]*", r"\1***", texto)
    texto = re.sub(r"(://[^:/\s]+:)[^@\s]+@", r"\1***@", texto)
    return texto


def texto_para_analise(sql: str) -> str:
    """Tira comentário e conteúdo de string, para a checagem não se enganar."""
    saida: list[str] = []
    i = 0
    n = len(sql)
    while i < n:
        dois = sql[i : i + 2]
        if dois == "--":
            i += 2
            while i < n and sql[i] not in "\n\r":
                i += 1
            saida.append(" ")
            continue
        if dois == "/*":
            fim = sql.find("*/", i + 2)
            if fim < 0:
                raise ErroConciliacao("Comentário /* sem fechar.")
            i = fim + 2
            saida.append(" ")
            continue
        if sql[i] == "'":
            i += 1
            while i < n:
                if sql[i] == "'" and i + 1 < n and sql[i + 1] == "'":
                    i += 2
                    continue
                if sql[i] == "'":
                    i += 1
                    break
                i += 1
            else:
                raise ErroConciliacao("Texto entre aspas sem fechar.")
            saida.append("''")
            continue
        saida.append(sql[i])
        i += 1
    return "".join(saida)


def validar_sql(sql: str, origem: str) -> None:
    """Aceita um único SELECT (ou WITH ... SELECT). Recusa o resto."""
    if "PREENCHER_ORIGEM" in sql:
        raise ErroConciliacao(
            f"{origem} ainda não tem a consulta definitiva. "
            "Abra o arquivo, cole o SELECT que já roda hoje e apague a linha PREENCHER_ORIGEM."
        )
    analisado = texto_para_analise(sql).strip()
    if not analisado:
        raise ErroConciliacao(f"{origem} está vazio.")
    if _FOR_UPDATE.search(analisado):
        raise ErroConciliacao(f"{origem} não pode usar FOR UPDATE.")
    proibido = _PROIBIDO.search(analisado)
    if proibido:
        raise ErroConciliacao(
            f"{origem} foi recusado: encontrei '{proibido.group(0).upper()}'. "
            "Cada script precisa ser um único SELECT, com filtro por parâmetro (:nome)."
        )
    sem_ponto = analisado.rstrip().rstrip(";").strip()
    if ";" in sem_ponto:
        raise ErroConciliacao(
            f"{origem} tem mais de um comando. Deixe só o SELECT."
        )
    primeira = _PRIMEIRA.match(sem_ponto)
    if primeira is None or primeira.group(1).upper() not in {"SELECT", "WITH"}:
        raise ErroConciliacao(
            f"{origem} precisa começar com SELECT ou WITH."
        )


def parametros_no_sql(sql: str) -> set[str]:
    """Nomes :parametro, ignorando cast ::tipo do Postgres."""
    analisado = texto_para_analise(sql)
    nomes = set()
    for match in re.finditer(r":([A-Za-z_][A-Za-z0-9_]*)", analisado):
        inicio = match.start()
        if inicio > 0 and analisado[inicio - 1] == ":":
            continue
        nomes.add(match.group(1))
    return nomes
