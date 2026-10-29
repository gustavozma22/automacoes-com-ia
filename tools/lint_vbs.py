# -*- coding: utf-8 -*-
"""
lint_vbs.py - verificador estatico de VBScript para scripts de SAP GUI.

Roda ANTES de pedir para alguem executar. Nao prova que a sequencia de
telas funciona - prova apenas que o codigo nao tem erro bobo. Cada
categoria abaixo ja quebrou um lote em producao pelo menos uma vez.

Uso:  python tools/lint_vbs.py SCRIPT.vbs [MAPA_1.txt MAPA_2.txt ...]
Sai com codigo 1 se achar ERRO; avisos nao derrubam.
"""
import re, sys, io, os

PALAVRAS = set("""
if then else elseif end select case for each next to step do loop while until wend
sub function exit dim const set call option explicit on error resume goto not and or
xor mod is new class property get let byref byval public private default erase with
in stop randomize preserve redim nothing true false empty null me
""".split())

EMBUTIDAS = set(x.lower() for x in """
MsgBox InStr InStrRev Len Trim LTrim RTrim Left Right Mid UCase LCase Replace Split Join
Array UBound LBound CStr CLng CInt CDbl CBool CByte CDate CSng IsNull IsObject IsEmpty
IsNumeric IsArray IsDate Now Day Month Year Hour Minute Second String Err CreateObject
GetObject WScript Int Abs Round Rnd Timer Date Time Chr Asc Space StrReverse Filter Eval
Execute TypeName VarType Fix Sgn Sqr Exp Log Hex Oct FormatNumber FormatCurrency
FormatDateTime FormatPercent Weekday DateAdd DateDiff DatePart DateSerial TimeSerial
MonthName WeekdayName RGB InputBox Sin Cos Tan Atn ScriptEngine RegExp
vbCrLf vbTab vbNewLine vbNullString vbYes vbNo vbOK vbCancel vbYesNo vbOKCancel
vbLf vbCr vbBack vbFormFeed vbVerticalTab vbNullChar vbObjectError vbUseDefault
vbCritical vbInformation vbQuestion vbExclamation vbTextCompare vbBinaryCompare
""".split())

ABRE_FECHA = [
    ("sub", "end sub"), ("function", "end function"), ("select", "end select"),
    ("with", "end with"), ("class", "end class"), ("property", "end property"),
]

class Achado:
    def __init__(s, grau, linha, texto):
        s.grau, s.linha, s.texto = grau, linha, texto

def limpar(linhas):
    """Tira literais de texto e comentarios, junta continuacoes de linha."""
    saida = []
    for n, bruta in enumerate(linhas, 1):
        L = bruta.rstrip("\n").rstrip("\r")
        semtexto = re.sub(r'"(?:[^"]|"")*"', '""', L)
        pos = semtexto.find("'")
        if pos >= 0:
            corte = len(semtexto) - len(semtexto[pos:])
            L = L[:corte]
            semtexto = semtexto[:pos]
        if re.match(r'^\s*rem\b', semtexto, re.I):
            L, semtexto = "", ""
        saida.append((n, L, semtexto.strip()))
    # juntar continuacoes " _"
    juntas, buf, ini = [], "", None
    for n, L, S in saida:
        if S.endswith("_") and re.search(r'(^|\s)_$', S):
            buf += re.sub(r'\s*_$', ' ', S)
            if ini is None: ini = n
            continue
        juntas.append((ini or n, (buf + S).strip()))
        buf, ini = "", None
    if buf: juntas.append((ini or 0, buf.strip()))
    return juntas

def identificadores(txt):
    """Nomes usados, ignorando o que vem depois de ponto (propriedades)."""
    achados = []
    for m in re.finditer(r'(\.?)\b([A-Za-z_][A-Za-z0-9_]*)\b', txt):
        if m.group(1) == ".":
            continue
        achados.append((m.group(2), m.start(2)))
    return achados

def analisar(caminho, caminho_mapa=None):
    """caminho_mapa: um caminho ou uma lista de caminhos (o script pode ler mais de um mapa)."""
    if isinstance(caminho_mapa, str):
        caminho_mapa = [caminho_mapa]
    mapas = [m for m in (caminho_mapa or []) if os.path.exists(m)]
    with io.open(caminho, encoding="cp1252", errors="replace") as f:
        linhas = f.readlines()
    codigo = limpar(linhas)
    achados = []

    # ---------- 1. identificador terminado em underscore ----------
    # Em VBScript "_" e continuacao de linha: um nome terminado nele
    # quebra de forma silenciosa e confusa.
    for n, S in codigo:
        for nome, _ in identificadores(S):
            if nome.endswith("_") and nome.lower() not in PALAVRAS:
                achados.append(Achado("ERRO", n, f"identificador terminado em underscore: {nome}"))

    # ---------- 2. assinaturas ----------
    procs = {}          # nome minusculo -> (nome, n_args, linha)
    consts = {}
    globais = set()
    escopo_atual = None
    locais = {}
    for n, S in codigo:
        m = re.match(r'^(?:public\s+|private\s+)?(sub|function)\s+([A-Za-z_]\w*)\s*(\((.*?)\))?\s*$', S, re.I)
        if m:
            nome = m.group(2)
            args = m.group(4) or ""
            lista = [a.strip() for a in args.split(",") if a.strip()]
            chave = nome.lower()
            if chave in procs:
                achados.append(Achado("ERRO", n,
                    f"nome de rotina repetido: {nome} (ja existe na linha {procs[chave][2]}). "
                    "VBScript ignora maiusculas - a chamada falha em silencio"))
            procs[chave] = (nome, len(lista), n)
            escopo_atual = chave
            locais[chave] = set()
            for a in lista:
                locais[chave].add(re.sub(r'^(byval|byref)\s+', '', a, flags=re.I).strip().lower())
            continue
        if re.match(r'^end\s+(sub|function)\b', S, re.I):
            escopo_atual = None
            continue
        m = re.match(r'^const\s+(.*)$', S, re.I)
        if m:
            for parte in m.group(1).split(","):
                nm = parte.split("=")[0].strip()
                if nm: consts[nm.lower()] = n
            continue
        m = re.match(r'^(?:dim|redim(?:\s+preserve)?)\s+(.*)$', S, re.I)
        if m:
            for parte in m.group(1).split(","):
                nm = re.sub(r'\(.*$', '', parte).strip()
                if not nm: continue
                if escopo_atual: locais[escopo_atual].add(nm.lower())
                else: globais.add(nm.lower())

    # ---------- 3. colisao rotina x variavel ----------
    for chave, (nome, _, ln) in procs.items():
        if chave in globais:
            achados.append(Achado("ERRO", ln,
                f"'{nome}' e nome de rotina E de variavel. VBScript ignora maiusculas: "
                "os dois sao o mesmo nome e a chamada falha em silencio"))
        if chave in consts:
            achados.append(Achado("ERRO", ln, f"'{nome}' e nome de rotina E de constante"))

    # ---------- 4. blocos desbalanceados ----------
    pilha = []
    for n, S in codigo:
        b = S.lower()
        if not b: continue
        if re.match(r'^(public\s+|private\s+)?sub\s+\w+', b):        pilha.append(("sub", n))
        elif re.match(r'^(public\s+|private\s+)?function\s+\w+', b): pilha.append(("function", n))
        elif re.match(r'^end\s+sub\b', b):      fechar(pilha, "sub", n, achados)
        elif re.match(r'^end\s+function\b', b): fechar(pilha, "function", n, achados)
        elif re.match(r'^select\s+case\b', b):  pilha.append(("select", n))
        elif re.match(r'^end\s+select\b', b):   fechar(pilha, "select", n, achados)
        elif re.match(r'^with\b', b):           pilha.append(("with", n))
        elif re.match(r'^end\s+with\b', b):     fechar(pilha, "with", n, achados)
        elif re.match(r'^for\b', b):            pilha.append(("for", n))
        elif re.match(r'^next\b', b):           fechar(pilha, "for", n, achados)
        elif re.match(r'^do\b', b):             pilha.append(("do", n))
        elif re.match(r'^loop\b', b):           fechar(pilha, "do", n, achados)
        elif re.match(r'^while\b', b):          pilha.append(("while", n))
        elif re.match(r'^wend\b', b):           fechar(pilha, "while", n, achados)
        elif re.match(r'^end\s+if\b', b):       fechar(pilha, "if", n, achados)
        elif re.match(r'^if\b', b):
            m = re.search(r'\bthen\b(.*)$', b)
            if m and m.group(1).strip() == "":
                pilha.append(("if", n))
            elif not m:
                achados.append(Achado("ERRO", n, "If sem Then"))
    for tipo, n in pilha:
        achados.append(Achado("ERRO", n, f"bloco '{tipo}' aberto e nunca fechado"))

    # ---------- 5. chamada com numero errado de argumentos ----------
    escopo_arg = None
    for n, S in codigo:
        m = re.match(r'^(?:public\s+|private\s+)?(sub|function)\s+([A-Za-z_]\w*)', S, re.I)
        if m:
            escopo_arg = m.group(2).lower()
            continue
        if re.match(r'^end\s+(sub|function)\b', S, re.I):
            escopo_arg = None
            continue
        for chave, (nome, esperados, _) in procs.items():
            # atribuir ao proprio nome dentro da rotina e retorno, nao chamada
            if chave == escopo_arg:
                continue
            for m in re.finditer(r'(?<![.\w])' + re.escape(nome) + r'\s*(\()?', S, re.I):
                resto = S[m.end():]
                # nome seguido de "=" e atribuicao, nao chamada
                if not m.group(1) and re.match(r'^\s*=(?!=)', resto):
                    continue
                if m.group(1):
                    args = ate_fechar(resto)
                    if args is None: continue
                    passados = contar_args(args)
                else:
                    if not re.match(r'^\s*(call\s+)?' + re.escape(nome) + r'\b', S, re.I):
                        continue
                    passados = contar_args(resto)
                if passados != esperados:
                    achados.append(Achado("ERRO", n,
                        f"{nome} espera {esperados} argumento(s) e recebeu {passados}"))
                break

    # ---------- 6. variavel nao declarada ----------
    escopo_atual = None
    for n, S in codigo:
        m = re.match(r'^(?:public\s+|private\s+)?(sub|function)\s+([A-Za-z_]\w*)', S, re.I)
        if m:
            escopo_atual = m.group(2).lower(); continue
        if re.match(r'^end\s+(sub|function)\b', S, re.I):
            escopo_atual = None; continue
        if re.match(r'^(dim|const|redim)', S, re.I): continue
        for nome, _ in identificadores(S):
            b = nome.lower()
            if b in PALAVRAS or b in EMBUTIDAS or b in procs or b in consts or b in globais:
                continue
            if escopo_atual and b in locais.get(escopo_atual, set()):
                continue
            if escopo_atual and b == escopo_atual:   # valor de retorno
                continue
            achados.append(Achado("AVISO", n, f"nome nao declarado neste escopo: {nome}"))

    # ---------- 7. marcador {L} esquecido ----------
    # ATENCAO: roda sobre as linhas CRUAS. O {L} vive dentro de um
    # literal de texto, e o codigo limpo ja apagou os literais - foi
    # exatamente assim que esta checagem nasceu sem enxergar nada.
    # As linhas cruas precisam ter as continuacoes " _" juntadas antes
    # da checagem: senao um Replace(...) quebrado em duas linhas deixa o
    # {L} orfao na segunda e vira falso positivo.
    cruas, buf, ini = [], "", None
    for i, S in enumerate(linhas, 1):
        S = S.rstrip("\n").rstrip("\r")
        if re.search(r'\s_\s*$', S):
            buf += re.sub(r'\s*_\s*$', ' ', S)
            if ini is None: ini = i
            continue
        cruas.append((ini or i, buf + S))
        buf, ini = "", None
    if buf: cruas.append((ini or len(linhas), buf))

    for n, S in cruas:
        if S.strip().startswith("'"):     # comentario nao e codigo
            continue
        if "{L}" in S and "LinhaId" not in S and "Replace" not in S:
            achados.append(Achado("ERRO", n, "id com {L} usado sem passar por LinhaId()"))

    # ---------- 8. chave de mapa inexistente ----------
    if mapas:
        chaves = set()
        for um in mapas:
            with io.open(um, encoding="cp1252", errors="replace") as f:
                for L in f:
                    L = L.strip()
                    if L and not L.startswith("#") and "=" in L:
                        chaves.add(L.split("=", 1)[0].strip())
        for n, bruta in enumerate(linhas, 1):
            for m in re.finditer(r'Mapa\(\s*"([^"]+)"\s*\)', bruta, re.I):
                if m.group(1) not in chaves:
                    achados.append(Achado("ERRO", n,
                        f'Mapa("{m.group(1)}") nao existe no mapa de campos'))
    return achados

def fechar(pilha, tipo, n, achados):
    if not pilha:
        achados.append(Achado("ERRO", n, f"fechamento de '{tipo}' sem abertura"))
        return
    topo = pilha.pop()
    if topo[0] != tipo:
        achados.append(Achado("ERRO", n,
            f"fechamento de '{tipo}' mas o bloco aberto na linha {topo[1]} e '{topo[0]}'"))

def ate_fechar(txt):
    nivel, saida = 0, ""
    for ch in txt:
        if ch == "(": nivel += 1
        elif ch == ")":
            if nivel == 0: return saida
            nivel -= 1
        saida += ch
    return None

def contar_args(txt):
    txt = txt.strip()
    if not txt: return 0
    nivel, total = 0, 1
    for ch in txt:
        if ch in "([": nivel += 1
        elif ch in ")]": nivel -= 1
        elif ch == "," and nivel == 0: total += 1
    return total

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__); sys.exit(2)
    alvo = sys.argv[1]
    mapa = sys.argv[2:]
    res = analisar(alvo, mapa)
    erros = [a for a in res if a.grau == "ERRO"]
    avisos = [a for a in res if a.grau == "AVISO"]
    for a in sorted(erros, key=lambda x: x.linha):
        print(f"ERRO   linha {a.linha:>5}: {a.texto}")
    for a in sorted(avisos, key=lambda x: x.linha):
        print(f"AVISO  linha {a.linha:>5}: {a.texto}")
    print(f"\n{len(erros)} erro(s), {len(avisos)} aviso(s) em {alvo}")
    sys.exit(1 if erros else 0)
