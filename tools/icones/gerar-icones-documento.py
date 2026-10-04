"""Gera os .ico por TIPO DE DOCUMENTO (Plano 6, Task 10 — SO-09): o Explorer mostrava o icone do .exe em todo arquivo.

Uma folha com canto dobrado, faixa na cor do app (mTexto azul, mPlanilha verde, mSlides vermelho — as mesmas cores dos logos em
src/mOffice.Comum/Assets/logo-*.svg) e o rotulo por papel:
    documento    linhas de texto
    modelo       folha com borda tracejada
    macro        selo "!" no canto
    apresentacao triangulo "play" (so' mSlides: .ppsx/.odp de apresentacao)
    legado       rotulo "97" na faixa (.doc/.xls/.ppt)

Saida: src/mOffice.Comum/Icones/Documentos/{mtexto,mplanilha,mslides}-{documento,modelo,macro,legado}.ico + mslides-apresentacao.ico (13 arquivos),
cada um com 6 quadros PNG (16/24/32/48/64/256), no formato do mpdf.ico/mtexto.ico. Tudo e' desenhado com formas (nenhuma fonte do sistema):
a saida e' a mesma em qualquer maquina com a mesma versao do Pillow. Os .ico sao COMMITADOS; o IconesTests confere que o script os reproduz.

Uso:
    python tools/icones/gerar-icones-documento.py [pasta-de-saida]     (padrao: src/mOffice.Comum/Icones/Documentos)
Requer Pillow >= 10.
"""
import io
import pathlib
import struct
import sys

import PIL
from PIL import Image, ImageDraw

if int(PIL.__version__.split(".")[0]) < 10:
    sys.exit("Pillow >= 10 e' necessario (pip install 'pillow>=10').")

TAMANHOS = [256, 64, 48, 32, 24, 16]
SUPER = 4  # supersampling: desenha em 1024 e reduz com LANCZOS

# cor principal, cor escura (dobra/contorno), cor clara (linhas)
APPS = {
    "mtexto": ("#1D4E89", "#0F2E52", "#7FA8D9"),
    "mplanilha": ("#1E7145", "#0F4A2B", "#9FD8B3"),
    "mslides": ("#C0392B", "#7B241C", "#F5B7B1"),
}
PAPEIS = {
    "mtexto": ["documento", "modelo", "macro", "legado"],
    "mplanilha": ["documento", "modelo", "macro", "legado"],
    "mslides": ["documento", "modelo", "macro", "legado", "apresentacao"],
}
AMARELO = "#F2A900"


def desenhar(app, papel):
    """Folha 256x256 (unidades) ja' supersampleada; devolve a imagem RGBA grande."""
    principal, escuro, claro = APPS[app]
    n = 256 * SUPER
    img = Image.new("RGBA", (n, n), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)

    def u(v):
        return int(round(v * SUPER))

    def retangulo(x0, y0, x1, y1, r, fill):
        d.rounded_rectangle([u(x0), u(y0), u(x1), u(y1)], radius=u(r), fill=fill)

    # folha: contorno escuro + miolo branco, com o canto superior direito dobrado
    retangulo(36, 12, 220, 244, 16, escuro)
    retangulo(42, 18, 214, 238, 11, "#FFFFFF")
    d.polygon([(u(166), u(12)), (u(222), u(12)), (u(222), u(68))], fill=(0, 0, 0, 0))
    d.polygon([(u(166), u(18)), (u(214), u(66)), (u(166), u(66))], fill=claro)
    d.polygon([(u(166), u(12)), (u(220), u(66)), (u(166), u(66))], outline=escuro, width=u(6), fill=claro)

    # faixa na cor do app, na base da folha
    d.rounded_rectangle([u(42), u(176), u(214), u(238)], radius=u(11), fill=principal)
    d.rectangle([u(42), u(176), u(214), u(200)], fill=principal)

    def linhas(topo, quantas, larguras):
        for i in range(quantas):
            y = topo + i * 22
            retangulo(62, y, 62 + larguras[i % len(larguras)], y + 12, 6, claro)

    if papel in ("documento", "macro", "legado"):
        linhas(78, 4, [70, 100, 84, 100][:4] if papel != "macro" else [70, 100, 84, 60])
    elif papel == "modelo":
        # borda tracejada dentro da folha + duas linhas curtas
        x0, y0, x1, y1 = 60, 84, 196, 160
        passo, traco = 18, 11
        x = x0
        while x < x1:
            f = min(x + traco, x1)
            d.rectangle([u(x), u(y0), u(f), u(y0 + 5)], fill=principal)
            d.rectangle([u(x), u(y1 - 5), u(f), u(y1)], fill=principal)
            x += passo
        y = y0
        while y < y1:
            f = min(y + traco, y1)
            d.rectangle([u(x0), u(y), u(x0 + 5), u(f)], fill=principal)
            d.rectangle([u(x1 - 5), u(y), u(x1), u(f)], fill=principal)
            y += passo
        retangulo(80, 108, 150, 118, 5, claro)
        retangulo(80, 128, 126, 138, 5, claro)
    elif papel == "apresentacao":
        d.polygon([(u(98), u(80)), (u(98), u(160)), (u(166), u(120))], fill=principal)

    if papel == "legado":
        digito7seg(d, u, "97", 84, 186, 40, 22, "#FFFFFF")
    if papel == "macro":
        cx, cy, r = 190, 190, 40
        d.ellipse([u(cx - r - 5), u(cy - r - 5), u(cx + r + 5), u(cy + r + 5)], fill="#FFFFFF")
        d.ellipse([u(cx - r), u(cy - r), u(cx + r), u(cy + r)], fill=AMARELO)
        retangulo(cx - 6, cy - 24, cx + 6, cy + 6, 5, "#FFFFFF")
        d.ellipse([u(cx - 7), u(cy + 13), u(cx + 7), u(cy + 27)], fill="#FFFFFF")
    return img


SEGMENTOS = {  # a b c d e f g
    "9": "abcdfg", "7": "abc",
}


def digito7seg(d, u, texto, x, y, w, h, cor):
    """Desenha digitos de 7 segmentos com retangulos (sem fonte)."""
    esp = 8
    for k, ch in enumerate(texto):
        x0 = x + k * (w + esp)
        t = 8
        meio = y + h
        seg = {
            "a": (x0, y, x0 + w, y + t),
            "b": (x0 + w - t, y, x0 + w, meio + t // 2),
            "c": (x0 + w - t, meio - t // 2, x0 + w, y + 2 * h),
            "d": (x0, y + 2 * h - t, x0 + w, y + 2 * h),
            "e": (x0, meio - t // 2, x0 + t, y + 2 * h),
            "f": (x0, y, x0 + t, meio + t // 2),
            "g": (x0, meio - t // 2, x0 + w, meio + t // 2),
        }
        for s in SEGMENTOS[ch]:
            a, b, c, e = seg[s]
            d.rectangle([u(a), u(b), u(c), u(e)], fill=cor)


def ico(img_grande):
    pngs = []
    for t in TAMANHOS:
        reduzida = img_grande.resize((t, t), Image.LANCZOS)
        buf = io.BytesIO()
        reduzida.save(buf, format="PNG", optimize=False, compress_level=9)
        pngs.append(buf.getvalue())
    cab = struct.pack("<HHH", 0, 1, len(TAMANHOS))
    offset = 6 + 16 * len(TAMANHOS)
    entradas = b""
    for t, png in zip(TAMANHOS, pngs):
        entradas += struct.pack("<BBBBHHII", 0 if t == 256 else t, 0 if t == 256 else t, 0, 0, 1, 32, len(png), offset)
        offset += len(png)
    return cab + entradas + b"".join(pngs)


def main():
    raiz = pathlib.Path(__file__).resolve().parents[2]
    saida = pathlib.Path(sys.argv[1]) if len(sys.argv) > 1 else raiz / "src" / "mOffice.Comum" / "Icones" / "Documentos"
    saida.mkdir(parents=True, exist_ok=True)
    total = 0
    for app, papeis in PAPEIS.items():
        for papel in papeis:
            alvo = saida / f"{app}-{papel}.ico"
            alvo.write_bytes(ico(desenhar(app, papel)))
            total += 1
            print(f"{alvo.name} ({alvo.stat().st_size} bytes)")
    print(f"{total} icones em {saida}")


if __name__ == "__main__":
    main()
