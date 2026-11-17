"""Lê a pasta de scripts e a pasta de pipelines.

Um script é uma consulta. Um pipeline é uma conciliação: a lista de scripts
que ela usa, e qual par entra na conta. Outra conciliação repete os ids e
acrescenta os scripts dela.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, replace
from pathlib import Path

import yaml

from conciliacao.caminhos import pasta_pipelines, pasta_scripts
from conciliacao.erros import ErroConciliacao
from conciliacao.seguranca import parametros_no_sql, texto_para_analise, validar_sql

_ID = re.compile(r"^[a-z][a-z0-9_]*$")
SCRIPT_CT2 = "protheus_ct2_razao"
SCRIPTS_SINQIA = frozenset({"sinqia_ct_razao", "sinqia_ct_balancete", "sinqia_cc_razao"})
_CABECALHO = re.compile(r"^--\s*@([a-z0-9_.]+)\s*:\s*(.*?)\s*$")
_TIPOS = {"texto", "inteiro", "data"}
_ENGINES = {"oracle", "sqlserver", "postgres"}
_OPERACOES = {"soma", "subtracao"}
_ABAS_RESERVADAS = {
    "capa",
    "wp-01 saldos",
    "conciliação",
    "conciliacao",
    "composição",
    "composicao",
    "execução",
    "execucao",
    "exclusões",
    "exclusoes",
}
_REF = re.compile(r"^\$\{([a-zA-Z0-9_]+)\}$")


@dataclass(frozen=True)
class ParametroScript:
    nome: str
    tipo: str
    tamanho: int | None = None


@dataclass(frozen=True)
class Script:
    id: str
    engine: str
    descricao: str
    caminho: Path
    parametros: dict[str, ParametroScript]
    colunas: dict[str, str]
    sql: str

    def coluna(self, papel: str) -> str | None:
        nome = self.colunas.get(papel)
        return nome.upper() if nome else None


@dataclass(frozen=True)
class Fonte:
    script: str
    alias: str
    aba: str
    resumo: str
    parametros: dict[str, str]
    opcional: bool = False
    excluir: tuple | None = None


@dataclass(frozen=True)
class Pipeline:
    id: str
    nome: str
    empresa: str
    ativo: bool
    caminho: Path
    parametros: dict[str, str]
    fontes: list[Fonte]
    esquerda: str
    direita: str
    detalhe: str | None
    operacao: str
    tolerancia: float
    limite_linhas: int

    def fonte(self, alias: str) -> Fonte:
        for item in self.fontes:
            if item.alias == alias:
                return item
        raise ErroConciliacao(f"Pipeline {self.id} não tem a fonte '{alias}'.")

    def fonte_ct2(self) -> Fonte | None:
        for item in self.fontes:
            if item.script == SCRIPT_CT2:
                return item
        return None

    def fonte_sinqia(self) -> Fonte | None:
        for item in self.fontes:
            if item.script == "sinqia_ct_razao":
                return item
        for item in self.fontes:
            if item.script in SCRIPTS_SINQIA:
                return item
        return None


def _ler_texto(caminho: Path) -> str:
    return caminho.read_text(encoding="utf-8-sig")


def _parse_script(caminho: Path) -> Script:
    texto = _ler_texto(caminho)
    cabecalho: dict[str, str] = {}
    parametros: dict[str, ParametroScript] = {}
    colunas: dict[str, str] = {}
    for linha in texto.splitlines():
        match = _CABECALHO.match(linha.strip())
        if not match:
            if linha.strip() and not linha.strip().startswith("--"):
                break
            continue
        chave, valor = match.group(1), match.group(2).strip()
        if chave == "param":
            partes = valor.split()
            if not partes:
                raise ErroConciliacao(f"{caminho.name}: @param sem nome.")
            nome = partes[0]
            tipo = "texto"
            tamanho = None
            for parte in partes[1:]:
                if parte.startswith("tipo="):
                    tipo = parte.split("=", 1)[1]
                elif parte.startswith("tamanho="):
                    bruto = parte.split("=", 1)[1]
                    if not bruto.isdigit() or int(bruto) < 1:
                        raise ErroConciliacao(
                            f"{caminho.name}: tamanho de @{nome} precisa ser um inteiro maior que zero."
                        )
                    tamanho = int(bruto)
            if tipo not in _TIPOS:
                raise ErroConciliacao(
                    f"{caminho.name}: tipo '{tipo}' inválido em @{nome}. Use texto, inteiro ou data."
                )
            parametros[nome] = ParametroScript(nome, tipo, tamanho)
        elif chave.startswith("coluna."):
            papel = chave.split(".", 1)[1]
            if not valor:
                raise ErroConciliacao(f"{caminho.name}: @coluna.{papel} sem nome de coluna.")
            colunas[papel] = valor
        elif chave in {"id", "engine", "descricao"}:
            cabecalho[chave] = valor
        else:
            raise ErroConciliacao(f"{caminho.name}: cabeçalho @{chave} não é reconhecido.")

    id_script = cabecalho.get("id", "")
    engine = cabecalho.get("engine", "").lower()
    descricao = cabecalho.get("descricao", "")
    if not _ID.match(id_script):
        raise ErroConciliacao(
            f"{caminho.name}: @id '{id_script}' inválido. Use letras minúsculas, número e _."
        )
    if caminho.stem != id_script:
        raise ErroConciliacao(
            f"{caminho.name}: o @id precisa ser igual ao nome do arquivo ({caminho.stem})."
        )
    if engine not in _ENGINES:
        raise ErroConciliacao(
            f"{caminho.name}: @engine '{engine}' inválido. Use oracle, sqlserver ou postgres."
        )
    if not descricao:
        raise ErroConciliacao(f"{caminho.name}: falta @descricao.")
    return Script(id_script, engine, descricao, caminho, parametros, colunas, texto)


def carregar_scripts(pasta: Path | None = None) -> dict[str, Script]:
    base = (pasta or pasta_scripts()).resolve()
    if not base.is_dir():
        raise ErroConciliacao(f"Pasta de scripts não encontrada: {base}")
    achados: dict[str, Script] = {}
    for caminho in sorted(base.rglob("*.sql")):
        if not caminho.resolve().is_relative_to(base):
            raise ErroConciliacao(f"Script fora da pasta de scripts: {caminho}")
        script = _parse_script(caminho)
        if script.id in achados:
            raise ErroConciliacao(
                f"O id '{script.id}' está em dois arquivos: {achados[script.id].caminho} e {caminho}."
            )
        achados[script.id] = script
    return achados


def _texto_param(valor, origem: str) -> str:
    if isinstance(valor, bool) or valor is None:
        raise ErroConciliacao(f"{origem}: parâmetro vazio ou inválido.")
    if isinstance(valor, (int, float)):
        raise ErroConciliacao(
            f"{origem}: coloque o valor entre aspas. "
            "Número sem aspas perde zero à esquerda (030101 vira outro valor)."
        )
    if isinstance(valor, str):
        return valor.strip()
    raise ErroConciliacao(f"{origem}: parâmetro precisa ser texto.")


def _resolver_valor(valor: str, base: dict[str, str], origem: str) -> str:
    match = _REF.match(valor.strip())
    if not match:
        if "${" in valor:
            raise ErroConciliacao(
                f"{origem}: referência inválida '{valor}'. Use o valor inteiro como ${{nome}}."
            )
        return valor
    nome = match.group(1)
    if nome not in base:
        raise ErroConciliacao(f"{origem}: ${{{nome}}} não está nos parâmetros do pipeline.")
    return base[nome]


def _parse_pipeline(caminho: Path, scripts: dict[str, Script]) -> Pipeline:
    bruto = yaml.safe_load(_ler_texto(caminho))
    if not isinstance(bruto, dict):
        raise ErroConciliacao(f"{caminho.name}: o arquivo precisa ser um mapa YAML.")
    id_pipeline = str(bruto.get("id", "")).strip()
    if not _ID.match(id_pipeline):
        raise ErroConciliacao(f"{caminho.name}: id de pipeline inválido.")
    nome = str(bruto.get("nome", "")).strip()
    empresa = str(bruto.get("empresa", "")).strip()
    if not nome or not empresa:
        raise ErroConciliacao(f"{caminho.name}: preencha nome e empresa.")

    parametros_brutos = bruto.get("parametros") or {}
    if not isinstance(parametros_brutos, dict):
        raise ErroConciliacao(f"{caminho.name}: parametros precisa ser um mapa.")
    parametros = {
        str(chave): _texto_param(valor, f"{caminho.name} parametros.{chave}")
        for chave, valor in parametros_brutos.items()
    }

    fontes_brutas = bruto.get("fontes") or []
    if not isinstance(fontes_brutas, list) or not fontes_brutas:
        raise ErroConciliacao(f"{caminho.name}: inclua ao menos uma fonte.")
    fontes: list[Fonte] = []
    aliases: set[str] = set()
    abas: set[str] = set()
    for indice, item in enumerate(fontes_brutas, start=1):
        if not isinstance(item, dict):
            raise ErroConciliacao(f"{caminho.name}: fonte {indice} inválida.")
        script_id = str(item.get("script", "")).strip()
        if script_id not in scripts:
            raise ErroConciliacao(
                f"{caminho.name}: script '{script_id}' não está na pasta scripts."
            )
        alias = str(item.get("como") or script_id).strip()
        if not _ID.match(alias):
            raise ErroConciliacao(f"{caminho.name}: alias '{alias}' inválido.")
        if alias in aliases:
            raise ErroConciliacao(f"{caminho.name}: a fonte '{alias}' está repetida.")
        aliases.add(alias)
        aba = str(item.get("aba") or alias).strip()
        if aba.casefold() in _ABAS_RESERVADAS:
            raise ErroConciliacao(f"{caminho.name}: a aba '{aba}' é reservada ao motor.")
        if aba.casefold() in abas:
            raise ErroConciliacao(f"{caminho.name}: duas fontes usam a aba '{aba}'.")
        abas.add(aba.casefold())
        resumo = str(item.get("resumo") or scripts[script_id].descricao).strip()
        locais = item.get("parametros") or {}
        if not isinstance(locais, dict):
            raise ErroConciliacao(f"{caminho.name}: parametros da fonte {alias} inválidos.")
        resolvidos = {}
        for chave, valor in locais.items():
            texto = _texto_param(valor, f"{caminho.name} fonte {alias}.{chave}")
            resolvidos[str(chave)] = _resolver_valor(texto, parametros, f"{caminho.name} fonte {alias}")
        fontes.append(
            Fonte(
                script_id,
                alias,
                aba[:31],
                resumo,
                resolvidos,
                bool(item.get("opcional", False)),
            )
        )

    conciliacao = bruto.get("conciliacao") or {}
    if not isinstance(conciliacao, dict):
        raise ErroConciliacao(f"{caminho.name}: bloco conciliacao inválido.")
    esquerda = str(conciliacao.get("esquerda", "")).strip()
    direita = str(conciliacao.get("direita", "")).strip()
    detalhe = conciliacao.get("detalhe")
    detalhe = str(detalhe).strip() if detalhe else None
    for alias in (esquerda, direita, detalhe):
        if alias and alias not in aliases:
            raise ErroConciliacao(
                f"{caminho.name}: a conciliação aponta '{alias}', que não está em fontes."
            )
    if esquerda == direita:
        raise ErroConciliacao(f"{caminho.name}: esquerda e direita precisam ser fontes diferentes.")
    operacao = str(conciliacao.get("operacao", "soma")).strip()
    if operacao not in _OPERACOES:
        raise ErroConciliacao(f"{caminho.name}: operacao '{operacao}' inválida. Use soma ou subtracao.")
    tolerancia = float(conciliacao.get("tolerancia", 0.01))
    if tolerancia < 0:
        raise ErroConciliacao(f"{caminho.name}: tolerancia não pode ser negativa.")
    limite = int(bruto.get("limite_linhas", 200_000))
    if limite < 1:
        raise ErroConciliacao(f"{caminho.name}: limite_linhas inválido.")

    pipeline = Pipeline(
        id=id_pipeline,
        nome=nome,
        empresa=empresa,
        ativo=bool(bruto.get("ativo", True)),
        caminho=caminho,
        parametros=parametros,
        fontes=fontes,
        esquerda=esquerda,
        direita=direita,
        detalhe=detalhe,
        operacao=operacao,
        tolerancia=tolerancia,
        limite_linhas=limite,
    )
    _validar_contrato(pipeline, scripts)
    return pipeline


def _validar_contrato(pipeline: Pipeline, scripts: dict[str, Script]) -> None:
    for alias, papel_data, exige_valor in (
        (pipeline.esquerda, True, True),
        (pipeline.direita, True, True),
        (pipeline.detalhe, True, False),
    ):
        if not alias:
            continue
        script = scripts[pipeline.fonte(alias).script]
        if papel_data and not script.coluna("data"):
            raise ErroConciliacao(
                f"{script.caminho.name}: a fonte '{alias}' entra na conciliação e precisa de @coluna.data."
            )
        if exige_valor and not script.coluna("valor"):
            raise ErroConciliacao(
                f"{script.caminho.name}: a fonte '{alias}' entra na conta e precisa de @coluna.valor."
            )


def carregar_pipelines(
    scripts: dict[str, Script] | None = None,
    pasta: Path | None = None,
) -> dict[str, Pipeline]:
    base = (pasta or pasta_pipelines()).resolve()
    if not base.is_dir():
        raise ErroConciliacao(f"Pasta de pipelines não encontrada: {base}")
    scripts = scripts if scripts is not None else carregar_scripts()
    achados: dict[str, Pipeline] = {}
    for caminho in sorted(base.glob("*.yaml")):
        if caminho.name.startswith("_"):
            continue
        pipeline = _parse_pipeline(caminho, scripts)
        if pipeline.id in achados:
            raise ErroConciliacao(f"Dois pipelines com o id '{pipeline.id}'.")
        achados[pipeline.id] = pipeline
    return achados


def regras_exclusao(fonte: Fonte, script: Script) -> tuple:
    """Regras que tiram linhas da soma. Nesta pasta nenhuma fonte declara exclusão."""
    escolhida = getattr(fonte, "excluir", None)
    if escolhida is not None:
        return escolhida
    return tuple(getattr(script, "excluir", ()) or ())


def _parametro_protegido(sql: str, nome: str) -> str:
    """Tira RTRIM(:nome) e RPAD(:nome, ...). O que sobrar é comparação direta com a coluna."""
    sem = re.sub(rf"RTRIM\(\s*:{nome}\s*\)", " ", sql, flags=re.IGNORECASE)
    return re.sub(rf"RPAD\(\s*:{nome}\b[^)]*\)", " ", sem, flags=re.IGNORECASE)


def _exigir_tamanho_char(script: Script) -> None:
    """Coluna CHAR do Protheus não casa com texto curto, e RTRIM na coluna varre a tabela."""
    if script.engine != "oracle":
        return
    analisado = texto_para_analise(script.sql)
    faltando = []
    for nome, param in script.parametros.items():
        if param.tipo != "texto" or param.tamanho:
            continue
        if f":{nome}" not in _parametro_protegido(analisado, nome):
            continue
        faltando.append(nome)
    if faltando:
        raise ErroConciliacao(
            f"{script.caminho.name}: {', '.join(faltando)} compara com coluna CHAR do Protheus sem tamanho. "
            "Declare tamanho=N, o DATA_LENGTH da coluna. "
            "RTRIM na coluna acha a linha, mas lê a tabela inteira."
        )


def fonte_pronta(fonte: Fonte, script: Script) -> bool:
    """False só quando a fonte é opcional e o SELECT ainda não foi colado."""
    if "PREENCHER_ORIGEM" in script.sql:
        if fonte.opcional:
            return False
        validar_sql(script.sql, script.caminho.name)
    return True


def preparar_execucao(pipeline: Pipeline, scripts: dict[str, Script]) -> None:
    """Confere SQL e parâmetros antes de abrir qualquer banco."""
    obrigatorias = {pipeline.esquerda, pipeline.direita}
    for fonte in pipeline.fontes:
        script = scripts[fonte.script]
        if not fonte_pronta(fonte, script):
            if fonte.alias in obrigatorias:
                raise ErroConciliacao(
                    f"{script.caminho.name} entra na conta e ainda não tem o SELECT definitivo."
                )
            continue
        validar_sql(script.sql, script.caminho.name)
        _exigir_tamanho_char(script)
        usados = parametros_no_sql(script.sql)
        sem_tipo = sorted(nome for nome in usados if nome not in script.parametros)
        if sem_tipo:
            raise ErroConciliacao(
                f"{script.caminho.name}: declare @param para {sem_tipo} no cabeçalho do script."
            )
        periodo = {"data_ini", "data_fim"}
        travados = sorted(nome for nome in fonte.parametros if nome in periodo)
        if travados:
            raise ErroConciliacao(
                f"Fonte '{fonte.alias}': {travados} vêm do período informado ao rodar, não do YAML."
            )
        faltando = sorted(nome for nome in usados if nome not in fonte.parametros and nome not in periodo)
        if faltando:
            raise ErroConciliacao(
                f"{script.caminho.name}: o SQL usa {faltando} e a fonte '{fonte.alias}' não informa esses parâmetros."
            )
        desconhecidos = sorted(
            nome for nome in fonte.parametros if nome not in script.parametros and nome not in usados
        )
        if desconhecidos:
            raise ErroConciliacao(
                f"Fonte '{fonte.alias}' manda {desconhecidos}, que o script {script.id} não declara."
            )


def _aplicar_conta(pipeline: Pipeline, alvos: set[str], conta: str | None, flag: str) -> Pipeline:
    if conta is None:
        return pipeline
    escolhida = conta.strip()
    if not escolhida:
        return pipeline
    if not any(fonte.script in alvos for fonte in pipeline.fontes):
        raise ErroConciliacao(f"{pipeline.id} não usa esses scripts. {flag} não se aplica.")
    novas = []
    for fonte in pipeline.fontes:
        if fonte.script not in alvos:
            novas.append(fonte)
            continue
        parametros = dict(fonte.parametros)
        parametros["conta"] = escolhida
        novas.append(replace(fonte, parametros=parametros))
    return replace(pipeline, fontes=novas)


def aplicar_conta_ct2(pipeline: Pipeline, conta: str | None) -> Pipeline:
    """Troca só a conta do script CT2."""
    return _aplicar_conta(pipeline, {SCRIPT_CT2}, conta, "--conta")


def aplicar_conta_sinqia(pipeline: Pipeline, conta: str | None) -> Pipeline:
    """Troca a conta dos scripts Sinqia (CT e CC, quando existir)."""
    return _aplicar_conta(pipeline, set(SCRIPTS_SINQIA), conta, "--conta-sinqia")
