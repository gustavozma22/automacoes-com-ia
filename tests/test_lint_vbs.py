"""Testes do lint_vbs com defeitos injetados.

Cada teste escreve um VBScript pequeno com UM defeito conhecido e confere
que o verificador o encontra. Um caso limpo garante que nao ha falso positivo.
"""
import sys, textwrap
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "tools"))
import lint_vbs  # noqa: E402


def rodar(tmp_path, codigo, mapa=None):
    vbs = tmp_path / "teste.vbs"
    vbs.write_text(textwrap.dedent(codigo), encoding="cp1252")
    caminho_mapa = None
    if mapa is not None:
        caminho_mapa = tmp_path / "MAPA.txt"
        caminho_mapa.write_text(textwrap.dedent(mapa), encoding="cp1252")
        caminho_mapa = str(caminho_mapa)
    achados = lint_vbs.analisar(str(vbs), caminho_mapa)
    return [a.texto for a in achados if a.grau == "ERRO"]


LIMPO = """
Option Explicit
Dim total
total = Somar(1, 2)
Function Somar(a, b)
    Somar = a + b
End Function
"""


def test_codigo_limpo_nao_gera_erro(tmp_path):
    assert rodar(tmp_path, LIMPO) == []


def test_rotina_com_mesmo_nome_de_variavel(tmp_path):
    erros = rodar(tmp_path, """
        Dim filho
        Function Filho(x)
            Filho = x
        End Function
    """)
    assert any("nome de rotina E de variavel" in e for e in erros)


def test_numero_errado_de_argumentos(tmp_path):
    erros = rodar(tmp_path, """
        Dim r
        r = Somar(1)
        Function Somar(a, b)
            Somar = a + b
        End Function
    """)
    assert any("espera 2 argumento(s) e recebeu 1" in e for e in erros)


def test_bloco_nao_fechado(tmp_path):
    erros = rodar(tmp_path, """
        Sub Teste()
            If 1 = 1 Then
                MsgBox "ok"
        End Sub
    """)
    assert any("'if'" in e for e in erros)


def test_identificador_terminado_em_underscore(tmp_path):
    erros = rodar(tmp_path, """
        Dim valor_
        valor_ = 1
    """)
    assert any("underscore" in e for e in erros)


def test_chave_de_mapa_inexistente(tmp_path):
    erros = rodar(tmp_path, """
        Dim m
        m = Mapa("CAMPO_QUE_NAO_EXISTE")
        Function Mapa(k)
            Mapa = k
        End Function
    """, mapa="""
        # mapa de teste
        CAMPO_EXISTENTE=wnd[0]/usr/txtX
    """)
    assert any("CAMPO_QUE_NAO_EXISTE" in e for e in erros)


def test_chave_de_mapa_em_minusculas_tambem_e_checada(tmp_path):
    erros = rodar(tmp_path, """
        Dim m
        m = mapa("OUTRO_CAMPO")
    """, mapa="CAMPO_EXISTENTE=wnd[0]/usr/txtX\n")
    assert any("OUTRO_CAMPO" in e for e in erros)


def test_varios_mapas_sao_combinados(tmp_path):
    vbs = tmp_path / "t.vbs"
    vbs.write_text('Dim a, b\na = Mapa("A")\nb = Mapa("B")\n', encoding="cp1252")
    (tmp_path / "m1.txt").write_text("A=1\n", encoding="cp1252")
    (tmp_path / "m2.txt").write_text("B=2\n", encoding="cp1252")
    achados = lint_vbs.analisar(str(vbs), [str(tmp_path / "m1.txt"), str(tmp_path / "m2.txt")])
    assert [a for a in achados if "nao existe no mapa" in a.texto] == []


def test_marcador_L_sem_LinhaId(tmp_path):
    erros = rodar(tmp_path, """
        Dim id
        id = "wnd[0]/usr/tbl/txtCAMPO[0,{L}]"
    """)
    assert any("{L}" in e for e in erros)
