import pytest

from conciliacao.conexoes import _preencher_char
from conciliacao.erros import ErroConciliacao


def test_texto_curto_e_completado_ate_o_tamanho_da_coluna():
    preenchido = _preencher_char(
        {"natureza": "1011240", "data_ini": "20261001", "conta": "11020202044"},
        {"natureza": 10, "data_ini": 8, "conta": 20},
    )
    assert preenchido["natureza"] == "1011240   "
    assert preenchido["data_ini"] == "20261001"
    assert preenchido["conta"] == "11020202044".ljust(20)
    assert len(preenchido["natureza"]) == 10
    assert len(preenchido["conta"]) == 20


def test_texto_maior_que_a_coluna_e_recusado():
    with pytest.raises(ErroConciliacao, match="natureza"):
        _preencher_char({"natureza": "10112409999"}, {"natureza": 10})
