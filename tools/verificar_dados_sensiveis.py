# -*- coding: utf-8 -*-
"""
verificar_dados_sensiveis.py - impede que dado real entre no repositorio.

Procura padroes que nunca deveriam aparecer num portfolio publico:
CNPJ, e-mail, caminho de rede, numero de pedido/documento SAP fora da
faixa ficticia. Termos especificos (nomes de empresa, fornecedores,
usuarios) ficam num arquivo LOCAL que nao vai para o Git:

    .termos-proibidos.txt   (um termo por linha; listado no .gitignore)

Uso:  python tools/verificar_dados_sensiveis.py [pasta]
Sai com codigo 1 se achar algo.
"""
import os, re, sys

EXTENSOES = {".vbs", ".ps1", ".bat", ".txt", ".md", ".json", ".py", ".yml", ".yaml"}
IGNORAR_PASTAS = {".git", "__pycache__", ".pytest_cache", "tests"}  # os testes contem exemplos ruins de proposito

# valores ficticios usados de proposito no repositorio
PERMITIDOS = {
    "11222333000181",            # CNPJ de exemplo
    "noreply@anthropic.com",
}
FICTICIOS = re.compile(r"^(4500000\d{3}|4500001234|5100000\d{3})$")  # faixa ficticia
DOMINIOS_EXEMPLO = (".exemplo", "empresa-exemplo.com.br", "example.com", "users.noreply.github.com")

PADROES = [
    ("CNPJ", re.compile(r"\b\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2}\b")),
    ("e-mail", re.compile(r"\b[\w.+-]+@[\w-]+(?:\.[\w-]+)+\b")),
    ("caminho de rede", re.compile(r"\\\\[A-Za-z0-9_.-]+\\[A-Za-z0-9_$.-]+")),
    ("pedido de compra real", re.compile(r"\b45\d{8}\b")),
    ("documento SAP real", re.compile(r"\b5[01]\d{8}\b")),
]


def termos_locais(raiz):
    caminho = os.path.join(raiz, ".termos-proibidos.txt")
    if not os.path.exists(caminho):
        return []
    with open(caminho, encoding="utf-8") as f:
        return [t.strip() for t in f if t.strip() and not t.startswith("#")]


def ler(caminho):
    with open(caminho, "rb") as f:
        b = f.read()
    for enc in ("utf-8", "cp1252"):
        try:
            return b.decode(enc)
        except UnicodeDecodeError:
            pass
    return b.decode("latin-1")


def verificar(raiz):
    termos = [(t, re.compile(r"(?i)\b" + re.escape(t) + r"\b")) for t in termos_locais(raiz)]
    achados = []
    for pasta, subpastas, arquivos in os.walk(raiz):
        subpastas[:] = [s for s in subpastas if s not in IGNORAR_PASTAS]
        for nome in arquivos:
            if os.path.splitext(nome)[1].lower() not in EXTENSOES or nome == ".termos-proibidos.txt":
                continue
            caminho = os.path.join(pasta, nome)
            for n, linha in enumerate(ler(caminho).splitlines(), 1):
                for rotulo, rx in PADROES:
                    for m in rx.finditer(linha):
                        v = m.group(0)
                        if v in PERMITIDOS or re.sub(r"\D", "", v) in PERMITIDOS or FICTICIOS.match(v):
                            continue
                        if rotulo == "e-mail" and v.lower().endswith(DOMINIOS_EXEMPLO):
                            continue
                        achados.append((os.path.relpath(caminho, raiz), n, rotulo, v))
                for t, rx in termos:
                    if rx.search(linha):
                        achados.append((os.path.relpath(caminho, raiz), n, "termo proibido", t))
    return achados


if __name__ == "__main__":
    raiz = sys.argv[1] if len(sys.argv) > 1 else "."
    achados = verificar(raiz)
    for arq, n, rotulo, v in achados:
        print(f"{arq}:{n}: {rotulo}: {v}")
    print(f"\n{len(achados)} ocorrencia(s) suspeita(s).")
    sys.exit(1 if achados else 0)
