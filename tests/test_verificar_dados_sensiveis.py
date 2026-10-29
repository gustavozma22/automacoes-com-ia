import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "tools"))
import verificar_dados_sensiveis as v  # noqa: E402


def achar(tmp_path, texto, termos=None):
    (tmp_path / "a.vbs").write_text(texto, encoding="utf-8")
    if termos:
        (tmp_path / ".termos-proibidos.txt").write_text("\n".join(termos), encoding="utf-8")
    return [r for _, _, r, _ in v.verificar(str(tmp_path))]


def test_valores_ficticios_passam(tmp_path):
    texto = "' CNPJ 11222333000181, pedido 4500000001, doc 5100000123, faturas@empresa-exemplo.com.br"
    assert achar(tmp_path, texto) == []


def test_cnpj_real_e_barrado(tmp_path):
    assert "CNPJ" in achar(tmp_path, "' fornecedor 12.345.678/0001-95")


def test_email_real_e_barrado(tmp_path):
    assert "e-mail" in achar(tmp_path, "' contato@fornecedor.com.br")


def test_caminho_de_rede_e_barrado(tmp_path):
    assert "caminho de rede" in achar(tmp_path, r"set BASE=\\servidor\compartilhado")


def test_pedido_real_e_barrado(tmp_path):
    assert "pedido de compra real" in achar(tmp_path, "' pedido 4500012345")


def test_termo_local_e_barrado(tmp_path):
    assert "termo proibido" in achar(tmp_path, "' empresa ACME lancou", termos=["acme"])


def test_documento_sap_real_e_barrado(tmp_path):
    assert "documento SAP real" in achar(tmp_path, "' doc 5100033268")
