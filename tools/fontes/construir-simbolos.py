#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
construir-simbolos.py - gera a fonte `mOffice Simbolos` (mOffice-Simbolos.ttf), desenhada por nos.

Por que existe (Marco 1 / decisao 1 do gate, 2026-09-19): o motor ONLYOFFICE traz embutida em base64, dentro de
`editors/sdkjs/{word,cell,slide,visio}/sdk-all.js`, uma fonte de simbolos internos da Ascensio (familia "ASCW3",
6.868 bytes, "All rights reserved", sem licenca localizavel). O mOffice troca esse blob pelo desta fonte
(`tools/onlyoffice/aplicar-patches.ps1`), que e' LIVRE (SIL OFL 1.1, titular: o projeto mOffice) e cujos
contornos sao TODOS desenhados aqui, em codigo - formas geometricas simples (setas, caixa, circulos, visto).
Nenhum byte nem contorno da fonte original entra neste script nem no repositorio.

O que e' IGUAL a fonte original, de proposito (sao numeros que o motor pressupoe, nao expressao criativa):
  - os mapeamentos do `cmap`: subtabela 3/0 (simbolica) com U+2610/U+2611 (caixa vazia / com visto),
    U+25C9/U+25CB (olho-de-boi / circulo vazio) e U+F020-U+F022, U+F038-U+F039 (marcas de formatacao); e a
    subtabela 1/0 formato 6 mapeando U+0020-U+0022 e U+0038-U+0039 aos MESMOS glifos das marcas - e' por
    este caminho que o editor chega neles: ele desenha `String.fromCharCode(33)` etc., nao o ponto de codigo
    da area privada (ver docs/licencas/AVISO-fontes.md, secao 1.2);
  - a ordem e os nomes dos 12 glifos (`.notdef`, `.null`, `uni000D` sem cmap, como na original);
  - unitsPerEm, ascent/descent/lineGap e as larguras de avanco de cada glifo - assim o layout das marcas
    nao muda (a seta da tabulacao continua ocupando o mesmo espaco);
  - as constantes de ALINHAMENTO (eixo das setas em y=740, ponta no LSB x=173, centro da caixa e dos
    circulos = metade do avanco): sao escolhas de metrica/posicionamento para as marcas cairem no mesmo
    lugar da linha, nao dados de contorno.

O que e' NOSSO: cada contorno abaixo (proporcoes, espessuras, o desenho do visto, os raios dos circulos), a
tabela `name` inteira (copyright "Copyright (c) 2026 mOffice", licenca OFL 1.1 nos nameID 13/14, sem nenhuma
mencao a marca de terceiros) e a ausencia de hinting (a fonte e' usada em tamanhos de tela; sem instrucoes ela
fica menor e mais simples de auditar).

Saida deterministica: mesmo script -> mesmos bytes (created/modified do `head` fixos; recalcTimestamp desligado).
O teste FonteSimbolosTests regenera a fonte numa pasta temporaria e compara byte a byte com a commitada.

Uso:
  python tools/fontes/construir-simbolos.py                       # grava tools/fontes/mOffice-Simbolos.ttf
  python tools/fontes/construir-simbolos.py --saida X.ttf         # grava em outro lugar
  python tools/fontes/construir-simbolos.py --familia "Outro Nome" # experimento do item 5 do brief da Task 1
  python tools/fontes/construir-simbolos.py --vazia --saida V.ttf # CONTROLE NEGATIVO de teste: mesma estrutura,
                                                                  # mesmos cmaps e metricas, contornos VAZIOS
                                                                  # (nada e' desenhado). Nunca vai para o motor.
Requer Python 3 + fontTools (`pip install fonttools`).
"""
import argparse
import hashlib
import os
import sys

from fontTools.fontBuilder import FontBuilder
from fontTools.misc.timeTools import timestampFromString
from fontTools.pens.ttGlyphPen import TTGlyphPen
from fontTools.ttLib import newTable
from fontTools.ttLib.tables._c_m_a_p import cmap_format_4, cmap_format_6

FAMILIA_PADRAO = "mOffice Simbolos"
VERSAO = "1.000"

# ---------------------------------------------------------------------------------------------------------
# Numeros que o motor pressupoe (iguais aos da fonte que este arquivo substitui - sao metricas, nao desenho).
# ---------------------------------------------------------------------------------------------------------
UPEM = 2048
ASCENT, DESCENT, LINE_GAP = 1491, -431, 307
WIN_ASCENT, WIN_DESCENT = 1854, 434

ORDEM = [".notdef", ".null", "uni000D",
         "uniF020", "uniF021", "uniF022", "uniF038", "uniF039",
         "uni25C9", "circle", "uni2610", "uni2611"]

AVANCOS = {
    ".notdef": 1024, ".null": 0, "uni000D": 1234,
    "uniF020": 512,                       # espaco (marca de formatacao): nada desenhado, so' avanco
    "uniF021": 1825, "uniF022": 1825,     # setas de tabulacao (esquerda / direita)
    "uniF038": 1135, "uniF039": 1135,     # marcas de quebra de linha (esquerda / direita)
    "uni25C9": 1764, "circle": 1764,      # botao de opcao marcado (olho-de-boi) / vazio (circulo)
    "uni2610": 1764, "uni2611": 1764,     # caixa de selecao vazia / com visto
}

CMAP_SIMBOLICA = {  # (3, 0) formato 4
    0x25C9: "uni25C9", 0x25CB: "circle", 0x2610: "uni2610", 0x2611: "uni2611",
    0xF020: "uniF020", 0xF021: "uniF021", 0xF022: "uniF022", 0xF038: "uniF038", 0xF039: "uniF039",
}
CMAP_MAC = {  # (1, 0) formato 6 - o caminho por onde o editor chega nas marcas (fromCharCode(33) etc.)
    0x20: "uniF020", 0x21: "uniF021", 0x22: "uniF022", 0x38: "uniF038", 0x39: "uniF039",
}

# Carimbo fixo do `head` (a data da decisao do gate). Mesmo input -> mesmos bytes.
CARIMBO = timestampFromString("Sat Sep 19 00:00:00 2026")


# ---------------------------------------------------------------------------------------------------------
# Desenho. Convencao TrueType (y para cima): contorno EXTERNO no sentido horario, FURO no anti-horario.
# `poligono`/`circulo` recebem os pontos em qualquer sentido e corrigem pela area com sinal.
# ---------------------------------------------------------------------------------------------------------
def _area_com_sinal(pontos):
    a = 0
    for i, (x0, y0) in enumerate(pontos):
        x1, y1 = pontos[(i + 1) % len(pontos)]
        a += x0 * y1 - x1 * y0
    return a / 2


def poligono(pen, pontos, furo=False):
    """Poligono de vertices retos. Externo -> horario (area negativa); furo -> anti-horario."""
    pontos = list(pontos)
    horario = _area_com_sinal(pontos) < 0
    if horario == furo:
        pontos.reverse()
    pen.moveTo(pontos[0])
    for p in pontos[1:]:
        pen.lineTo(p)
    pen.closePath()


def circulo(pen, cx, cy, r, furo=False):
    """Circulo por 8 arcos quadraticos: pontos on-curve a cada 45 graus, controles nas tangentes
    (raio r / cos 22.5). Sentido: horario para externo, anti-horario para furo."""
    import math
    n = 8
    passo = 2 * math.pi / n
    rc = r / math.cos(passo / 2)
    ons = [(cx + r * math.cos(i * passo), cy + r * math.sin(i * passo)) for i in range(n)]
    offs = [(cx + rc * math.cos((i + 0.5) * passo), cy + rc * math.sin((i + 0.5) * passo)) for i in range(n)]
    # ordem anti-horaria por construcao (angulo crescente); inverte para externo
    seq = [(ons[i], offs[i]) for i in range(n)]
    if not furo:
        seq = [(ons[(-i) % n], offs[(-i - 1) % n]) for i in range(n)]
    pen.moveTo(_arred(seq[0][0]))
    for i in range(n):
        off = seq[i][1]
        on = seq[(i + 1) % n][0]
        pen.qCurveTo(_arred(off), _arred(on))
    pen.closePath()


def _arred(p):
    return (int(round(p[0])), int(round(p[1])))


def _espelhar_x(pontos, largura_caixa):
    """Espelha um poligono na vertical do meio da caixa [0, largura_caixa]."""
    return [(largura_caixa - x, y) for (x, y) in pontos]


# Caixa horizontal das setas: LSB 173 ate 1653 (mesma faixa da largura de avanco 1825). Ponta na ESQUERDA.
_CX_SETA = 173 + 1653          # eixo de espelhamento (x' = 1826 - x)
_SETA_ESQUERDA = [             # seta <- : ponta em (173, 740), cabeca triangular, haste de 120 de espessura
    (173, 740), (640, 1070), (640, 800), (1653, 800), (1653, 680), (640, 680), (640, 410),
]
_QUEBRA_ESQUERDA = [           # marca de quebra: seta <- com uma barra vertical subindo na ponta direita
    (173, 740), (640, 1070), (640, 800), (1533, 800), (1533, 1440), (1653, 1440),
    (1653, 680), (640, 680), (640, 410),
]
_CAIXA_EXT = [(180, 10), (1584, 10), (1584, 1424), (180, 1424)]        # caixa de selecao, moldura de 110
_CAIXA_INT = [(290, 120), (1474, 120), (1474, 1314), (290, 1314)]
_VISTO = [(300, 760), (430, 890), (735, 560), (1300, 1230), (1420, 1110), (735, 300)]  # visto em V assimetrico
_NOTDEF_EXT = [(150, 0), (874, 0), (874, 1400), (150, 1400)]           # retangulo vazado, moldura de 100
_NOTDEF_INT = [(250, 100), (774, 100), (774, 1300), (250, 1300)]
_CIRC_CX, _CIRC_CY = 882, 716
_CIRC_R_EXT, _CIRC_R_INT, _CIRC_R_MIOLO = 720, 600, 360


def desenhar(nome, pen):
    """Contornos de cada glifo. Os tres sem cmap e o espaco ficam vazios."""
    if nome == ".notdef":
        poligono(pen, _NOTDEF_EXT)
        poligono(pen, _NOTDEF_INT, furo=True)
    elif nome in (".null", "uni000D", "uniF020"):
        pass
    elif nome == "uniF021":
        poligono(pen, _SETA_ESQUERDA)
    elif nome == "uniF022":
        poligono(pen, _espelhar_x(_SETA_ESQUERDA, _CX_SETA))
    elif nome == "uniF038":
        poligono(pen, _QUEBRA_ESQUERDA)
    elif nome == "uniF039":
        poligono(pen, _espelhar_x(_QUEBRA_ESQUERDA, _CX_SETA))
    elif nome == "circle":
        circulo(pen, _CIRC_CX, _CIRC_CY, _CIRC_R_EXT)
        circulo(pen, _CIRC_CX, _CIRC_CY, _CIRC_R_INT, furo=True)
    elif nome == "uni25C9":
        circulo(pen, _CIRC_CX, _CIRC_CY, _CIRC_R_EXT)
        circulo(pen, _CIRC_CX, _CIRC_CY, _CIRC_R_INT, furo=True)
        circulo(pen, _CIRC_CX, _CIRC_CY, _CIRC_R_MIOLO)
    elif nome == "uni2610":
        poligono(pen, _CAIXA_EXT)
        poligono(pen, _CAIXA_INT, furo=True)
    elif nome == "uni2611":
        poligono(pen, _CAIXA_EXT)
        poligono(pen, _CAIXA_INT, furo=True)
        poligono(pen, _VISTO)
    else:
        raise ValueError("glifo desconhecido: " + nome)


def _panose_pictorial():
    """PANOSE de fonte de simbolos: bFamilyType 5 (Latin Pictorial), demais campos 'any' (0)."""
    from fontTools.ttLib.tables.O_S_2f_2 import Panose
    p = Panose()
    for campo in ("bFamilyType", "bSerifStyle", "bWeight", "bProportion", "bContrast", "bStrokeVariation",
                  "bArmStyle", "bLetterForm", "bMidline", "bXHeight"):
        setattr(p, campo, 0)
    p.bFamilyType = 5
    return p


# ---------------------------------------------------------------------------------------------------------
def construir(familia=FAMILIA_PADRAO, vazia=False):
    fb = FontBuilder(UPEM, isTTF=True)
    fb.font.recalcTimestamp = False
    fb.setupGlyphOrder(ORDEM)

    glifos, metricas = {}, {}
    for nome in ORDEM:
        pen = TTGlyphPen(None)
        if not vazia:
            desenhar(nome, pen)
        g = pen.glyph()
        g.recalcBounds(None)
        glifos[nome] = g
        lsb = g.xMin if g.numberOfContours > 0 else 0
        metricas[nome] = (AVANCOS[nome], lsb)
    fb.setupGlyf(glifos)
    fb.setupHorizontalMetrics(metricas)
    fb.setupHorizontalHeader(ascent=ASCENT, descent=DESCENT, lineGap=LINE_GAP)

    # cmap: as duas subtabelas que o motor usa (3/0 simbolica formato 4; 1/0 formato 6). Construidas a mao
    # porque FontBuilder.setupCharacterMap so' faz a 3/1 Unicode, que aqui nao serve.
    cmap = newTable("cmap")
    cmap.tableVersion = 0
    t6 = cmap_format_6(6)
    t6.platformID, t6.platEncID, t6.language = 1, 0, 0
    t6.cmap = dict(CMAP_MAC)
    t4 = cmap_format_4(4)
    t4.platformID, t4.platEncID, t4.language = 3, 0, 0
    t4.cmap = dict(CMAP_SIMBOLICA)
    cmap.tables = [t6, t4]
    fb.font["cmap"] = cmap

    ps = "".join(ch for ch in familia if ch.isalnum()) + "-Regular"
    fb.setupNameTable({
        "copyright": "Copyright (c) 2026 mOffice",
        "familyName": familia,
        "styleName": "Regular",
        "uniqueFontIdentifier": "mOffice;" + ps + ";" + VERSAO,
        "fullName": familia,
        "version": "Version " + VERSAO,
        "psName": ps,
        "description": "Simbolos internos do editor do mOffice: marcas de formatacao (tabulacao, quebra), "
                       "caixa de selecao e botao de opcao de formulario. Desenhada em codigo "
                       "(tools/fontes/construir-simbolos.py).",
        "licenseDescription": "This Font Software is licensed under the SIL Open Font License, Version 1.1. "
                              "This license is available with a FAQ at: https://openfontlicense.org",
        "licenseInfoURL": "https://openfontlicense.org",
    })
    avancos_positivos = [a for a in AVANCOS.values() if a > 0]
    fb.setupOS2(
        version=4,
        xAvgCharWidth=int(round(sum(avancos_positivos) / len(avancos_positivos))),
        usWeightClass=400, usWidthClass=5, fsType=0,
        sTypoAscender=ASCENT, sTypoDescender=DESCENT, sTypoLineGap=LINE_GAP,
        usWinAscent=WIN_ASCENT, usWinDescent=WIN_DESCENT,
        fsSelection=0x40,                       # REGULAR
        achVendID="mOff",
        ulUnicodeRange1=0, ulUnicodeRange2=0, ulUnicodeRange3=0, ulUnicodeRange4=0,
        ulCodePageRange1=1 << 31,               # bit 31 = fonte simbolica (cmap 3/0)
        ulCodePageRange2=0,
        usFirstCharIndex=min(CMAP_SIMBOLICA), usLastCharIndex=max(CMAP_SIMBOLICA),
        sxHeight=1062, sCapHeight=1467, usDefaultChar=0, usBreakChar=0x20, usMaxContext=0,
        panose=_panose_pictorial(),             # 5 = Latin Pictorial (fonte de simbolos)
    )
    fb.setupPost()
    # unitsPerEm repetido de proposito: setupHead reinicia a tabela com os padroes (upem 1000) - sem isto a
    # fonte sairia com 1000 e todas as metricas acima ficariam fora de escala.
    fb.setupHead(unitsPerEm=UPEM, created=CARIMBO, modified=CARIMBO, flags=0x3, lowestRecPPEM=8,
                 fontRevision=float(VERSAO))
    return fb.font


def main(argv=None):
    ap = argparse.ArgumentParser(description="Gera mOffice-Simbolos.ttf (fonte de simbolos livre do mOffice).")
    ap.add_argument("--saida", default=os.path.join(os.path.dirname(os.path.abspath(__file__)), "mOffice-Simbolos.ttf"))
    ap.add_argument("--familia", default=FAMILIA_PADRAO, help="nome da familia na tabela name (padrao: %(default)s)")
    ap.add_argument("--vazia", action="store_true", help="controle negativo: mesmos cmaps/metricas, contornos vazios")
    args = ap.parse_args(argv)

    fonte = construir(args.familia, args.vazia)
    os.makedirs(os.path.dirname(os.path.abspath(args.saida)) or ".", exist_ok=True)
    fonte.save(args.saida)
    dados = open(args.saida, "rb").read()
    print("%s: %d bytes, sha256 %s (familia %r%s)" % (
        args.saida, len(dados), hashlib.sha256(dados).hexdigest(), args.familia, ", VAZIA" if args.vazia else ""))
    return 0


if __name__ == "__main__":
    sys.exit(main())
