from pathlib import Path

from conciliacao.caminhos import pasta_saida, pasta_trabalho, raiz


def test_raiz_usa_a_pasta_de_onde_o_comando_roda(monkeypatch):
    monkeypatch.delenv("CONCILIACAO_RAIZ", raising=False)
    assert raiz().name == "conciliacao"
    assert (raiz() / "scripts").is_dir()
    assert (raiz() / "pyproject.toml").is_file()


def test_saida_fica_na_pasta_do_bat_mesmo_com_codigo_em_outro_lugar(monkeypatch, tmp_path: Path):
    monkeypatch.setenv("CONCILIACAO_RAIZ", str(tmp_path))
    assert pasta_trabalho() == tmp_path.resolve()
    assert pasta_saida() == tmp_path.resolve() / "saida"
    assert not pasta_saida().exists()
    assert raiz().name == "conciliacao"


def test_saida_nao_existe_antes_de_gravar(monkeypatch, tmp_path: Path):
    monkeypatch.setenv("CONCILIACAO_RAIZ", str(tmp_path / "copia"))
    assert not pasta_saida().exists()
