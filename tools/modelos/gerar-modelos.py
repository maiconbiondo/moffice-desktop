#!/usr/bin/env python3
"""Gera os tres modelos vazios proprios do mOffice (Plano 6, Task 9) a partir dos do motor ONLYOFFICE (AGPL).

Entrada : third_party/onlyoffice/converter/empty/pt-BR/new.{docx,xlsx,pptx}
Saida   : src/mOffice.Nucleo/Modelos/pt-BR/new.{docx,xlsx,pptx}

Trocas EXATAS: cada uma tem de casar o numero de vezes esperado; se qualquer uma falhar o script sai com 1 SEM gravar
nada. As demais partes do pacote sao copiadas byte a byte (mesma ordem, mesma compressao, mesma data de cada entrada),
entao a saida e' reprodutivel. Registrado em docs/licencas/MODIFICACOES-ONLYOFFICE.md.

  docx: tema Calibri Light/Calibri; padrao do Word 2016-2021 (1,08 = line 259, 8 pt depois = after 160). Decisao do dono.
  pptx: tema Calibri Light/Calibri e nomes dos 11 layouts como no PowerPoint pt-BR.
  xlsx: 1a aba "Planilha1" (o do motor vem "Folha1") e sem o caminho de rede do autor original (x15ac:absPath).
"""
import os
import re
import sys
import zipfile

RAIZ = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
ORIGEM = os.path.join(RAIZ, "third_party", "onlyoffice", "converter", "empty", "pt-BR")
DESTINO = os.path.join(RAIZ, "src", "mOffice.Nucleo", "Modelos", "pt-BR")

NOMES_LAYOUT = {
    "title": "Slide de Título", "obj": "Título e Conteúdo", "secHead": "Cabeçalho da Seção", "twoObj": "Duas Partes de Conteúdo",
    "twoTxTwoObj": "Comparação", "titleOnly": "Somente Título", "blank": "Em branco", "objTx": "Conteúdo com Legenda",
    "picTx": "Imagem com Legenda", "vertTx": "Título e Texto Vertical", "vertTitleAndTx": "Título Vertical e Texto",
}
PANOSE_CALIBRI = "020F0502020204030204"
PANOSE_CALIBRI_LIGHT = "020F0302020204030204"


class Falha(Exception):
    pass


def trocar(dados: bytes, de: bytes, para: bytes, vezes: int, onde: str) -> bytes:
    n = dados.count(de)
    if n != vezes:
        raise Falha(f"{onde}: esperava {vezes} ocorrencia(s) de {de[:70]!r}, achei {n}")
    return dados.replace(de, para)


def regex(dados: bytes, padrao: bytes, novo, vezes: int, onde: str) -> bytes:
    r, n = re.subn(padrao, novo, dados)
    if n != vezes:
        raise Falha(f"{onde}: esperava {vezes} casamento(s) de {padrao[:70]!r}, achei {n}")
    return r


def fonte_do_tema(dados: bytes, onde: str) -> bytes:
    # <a:majorFont><a:latin typeface="Arial"[ panose="..."]/> -> Calibri Light / Calibri (com o panose certo da fonte nova)
    dados = regex(dados, rb'<a:majorFont><a:latin typeface="Arial"(?: panose="[0-9A-F]*")?',
                  f'<a:majorFont><a:latin typeface="Calibri Light" panose="{PANOSE_CALIBRI_LIGHT}"'.encode(), 1, onde + " major")
    return regex(dados, rb'<a:minorFont><a:latin typeface="Arial"(?: panose="[0-9A-F]*")?',
                 f'<a:minorFont><a:latin typeface="Calibri" panose="{PANOSE_CALIBRI}"'.encode(), 1, onde + " minor")


def docx(parte: str, dados: bytes) -> bytes:
    if parte == "word/theme/theme1.xml":
        return fonte_do_tema(dados, parte)
    if parte == "word/styles.xml":
        return trocar(dados, b'<w:spacing w:after="200" w:line="276" w:lineRule="auto"/>',
                      b'<w:spacing w:after="160" w:line="259" w:lineRule="auto"/>', 1, parte)
    return dados


def pptx(parte: str, dados: bytes) -> bytes:
    if re.fullmatch(r"ppt/theme/theme\d+\.xml", parte):
        return fonte_do_tema(dados, parte)
    if re.fullmatch(r"ppt/slideLayouts/slideLayout\d+\.xml", parte):
        tipo = re.search(rb'<p:sldLayout[^>]* type="([^"]+)"', dados)
        if not tipo or tipo.group(1).decode() not in NOMES_LAYOUT:
            raise Falha(f"{parte}: tipo de layout desconhecido")
        nome = NOMES_LAYOUT[tipo.group(1).decode()].encode("utf-8")
        return regex(dados, rb'<p:cSld name="[^"]*"', b'<p:cSld name="' + nome + b'"', 1, parte)
    if parte == "docProps/app.xml":  # "Fonts Used" das propriedades do arquivo
        return trocar(dados, b"<vt:lpstr>Arial</vt:lpstr>", b"<vt:lpstr>Calibri</vt:lpstr>", 1, parte)
    return dados


def xlsx(parte: str, dados: bytes) -> bytes:
    if parte == "xl/workbook.xml":
        dados = trocar(dados, b'<sheet name="Folha1"', b'<sheet name="Planilha1"', 1, parte)
        return regex(dados, rb'<mc:AlternateContent [^>]*><mc:Choice Requires="x15"><x15ac:absPath [^>]*/></mc:Choice></mc:AlternateContent>', b"", 1, parte)
    return dados


def gerar(nome: str, transformar, esperado_layouts: int = 0) -> bytes:
    caminho = os.path.join(ORIGEM, nome)
    if not os.path.isfile(caminho):
        raise Falha(f"modelo do motor ausente: {caminho}")
    saida = os.path.join(DESTINO, nome)
    tmp = saida + ".tmp"
    layouts = 0
    with zipfile.ZipFile(caminho) as z, zipfile.ZipFile(tmp, "w") as w:
        for info in z.infolist():
            dados = z.read(info.filename)
            novo = transformar(info.filename, dados)
            if re.fullmatch(r"ppt/slideLayouts/slideLayout\d+\.xml", info.filename):
                layouts += 1
            zi = zipfile.ZipInfo(info.filename, info.date_time)
            zi.compress_type = info.compress_type
            zi.external_attr = info.external_attr
            zi.create_system = info.create_system
            w.writestr(zi, novo)
    if esperado_layouts and layouts != esperado_layouts:
        os.remove(tmp)
        raise Falha(f"{nome}: esperava {esperado_layouts} layouts, achei {layouts}")
    return tmp


def main() -> int:
    os.makedirs(DESTINO, exist_ok=True)
    pendentes = []
    try:
        for nome, f, lay in (("new.docx", docx, 0), ("new.xlsx", xlsx, 0), ("new.pptx", pptx, 11)):
            pendentes.append((nome, gerar(nome, f, lay)))
    except Falha as e:
        for _, t in pendentes:
            if os.path.exists(t):
                os.remove(t)
        # um .tmp parcial do arquivo que falhou
        for nome in ("new.docx", "new.xlsx", "new.pptx"):
            t = os.path.join(DESTINO, nome + ".tmp")
            if os.path.exists(t):
                os.remove(t)
        print("ERRO: " + str(e).encode("ascii", "replace").decode("ascii"))
        return 1
    for nome, t in pendentes:
        os.replace(t, os.path.join(DESTINO, nome))
        print("gerado " + os.path.join("src", "mOffice.Nucleo", "Modelos", "pt-BR", nome))
    return 0


if __name__ == "__main__":
    sys.exit(main())
