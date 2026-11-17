import os
from pathlib import Path

from conciliacao.erros import ErroConciliacao


def _eh_projeto(pasta: Path) -> bool:
    return (pasta / "scripts").is_dir() and (
        (pasta / "pyproject.toml").is_file() or (pasta / "rodar.bat").is_file()
    )


def pasta_trabalho() -> Path:
    """Pasta de onde o rodar.bat foi clicado. O Excel nasce aqui, não no código instalado."""
    bruto = os.environ.get("CONCILIACAO_RAIZ", "").strip().strip('"')
    if bruto:
        return Path(bruto).expanduser().resolve()
    return Path.cwd().resolve()


def raiz() -> Path:
    """Scripts, pipelines e .env: pasta de trabalho se for cópia do projeto; senão o código instalado."""
    trabalho = pasta_trabalho()
    for pasta in [trabalho, *trabalho.parents]:
        if _eh_projeto(pasta):
            return pasta
    aqui = Path(__file__).resolve()
    for pasta in [aqui, *aqui.parents]:
        if _eh_projeto(pasta):
            return pasta
    raise ErroConciliacao(
        "Não encontrei a raiz do projeto (pasta scripts junto de pyproject.toml ou rodar.bat). "
        "Abra o rodar.bat de dentro da pasta conciliacao que você quer usar."
    )


def pasta_scripts() -> Path:
    return raiz() / "scripts"


def pasta_pipelines() -> Path:
    return raiz() / "pipelines"


def pasta_saida() -> Path:
    """Criada só na hora de gravar o Excel, ao lado do rodar.bat que você clicou."""
    return pasta_trabalho() / "saida"
