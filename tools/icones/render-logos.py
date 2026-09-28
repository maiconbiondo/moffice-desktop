"""Rasteriza os logos da suite (Logo.mTexto / Logo.mPlanilha / Logo.mSlides) em PNG de 256/48/32/16 px.

FONTE DA VERDADE: src/mOffice.Comum/Temas/mOfficeTema.xaml. Este script NAO transcreve a geometria — ele
faz o PARSE dos <DrawingImage x:Key="Logo.*"> daquele arquivo, entao mexer no logo do tema e rodar
tools/icones/gerar-icones.ps1 e' suficiente; nao ha' uma segunda copia da marca para sair de sincronia.

POR QUE NAO E' O WPF QUEM RASTERIZA (desvio do plano, registrado no relatorio da Task 6): o desenho
original era `mOffice.Sonda --gerar-icones` com RenderTargetBitmap sobre o proprio DrawingImage. Medido
nesta maquina (sessao RDP, RenderCapability.Tier 0): RenderTargetBitmap devolve um bitmap TODO ZERADO
mesmo para um DrawRectangle solido — reproduzido tambem em PowerShell puro, com e sem
ProcessRenderMode=SoftwareOnly e com e sem janela na tela, ou seja, e' a sessao, nao o codigo. Um
gerador que escreve icones em branco em silencio e' pior do que nao ter gerador. Este caminho nao
depende de GPU, de sessao grafica nem de WPF: roda igual no desenvolvedor, no servidor e no CI — que e'
exatamente o que o teste de regeneracao reprodutivel (IconesTests) precisa para valer alguma coisa.

Suporta SO' as tres formas que os logos usam hoje; qualquer outra coisa no XAML levanta erro (falha
FECHADA, nunca desenha "o que deu"):
  - <RectangleGeometry Rect="x,y,l,a" RadiusX="r" RadiusY="r"/> preenchido por Brush
  - Geometry="M x,y L x,y L x,y Z"  (poligono preenchido)
  - Geometry="M x,y H x2" com <Pen Thickness="t" .../> de pontas arredondadas (linha)

Requer **Pillow >= 10** (`pip install "pillow>=10"`) — versao verificada na subida do script. O teste de
regeneracao (IconesTests.Ico_RegeneracaoReprodutivel) compara PIXEIS, nao bytes, justamente porque os bytes
do PNG mudam entre versoes do Pillow sem o desenho mudar. Uso:
    python tools/icones/render-logos.py <pasta-de-saida>
"""
import pathlib
import re
import sys
import xml.etree.ElementTree as ET

from PIL import Image, ImageDraw
from PIL import __version__ as VERSAO_PILLOW

if int(VERSAO_PILLOW.split(".")[0]) < 10:
    raise SystemExit(f"Pillow >= 10 necessario (achado {VERSAO_PILLOW}); rode: pip install 'pillow>=10'")

NS = {
    "p": "http://schemas.microsoft.com/winfx/2006/xaml/presentation",
    "x": "http://schemas.microsoft.com/winfx/2006/xaml",
}
TAMANHOS = [256, 48, 32, 16]
# Supersampling: o Pillow nao tem antialiasing no desenho, entao desenha-se GRANDE e reduz-se com
# LANCZOS. 8x e' o suficiente para a borda arredondada de 16 px nao ficar serrilhada.
SUPER = 8
APPS = {"Logo.mTexto": "mtexto", "Logo.mPlanilha": "mplanilha", "Logo.mSlides": "mslides"}


def cor(texto):
    t = texto.strip().lstrip("#")
    if len(t) == 6:
        return (int(t[0:2], 16), int(t[2:4], 16), int(t[4:6], 16), 255)
    if len(t) == 8:  # AARRGGBB, como o WPF escreve
        return (int(t[2:4], 16), int(t[4:6], 16), int(t[6:8], 16), int(t[0:2], 16))
    raise ValueError(f"cor nao reconhecida: {texto!r}")


def numeros(texto):
    return [float(n) for n in re.split(r"[,\s]+", texto.strip()) if n]


def caminho(d):
    """Traduz o mini-path do XAML para ('poligono', pontos) ou ('linha', (x1, y1, x2, y2))."""
    partes = d.replace(",", " ").split()
    if len(partes) >= 2 and partes[0].startswith("M") and partes[-1].upper() == "Z":
        # "M176,16 L218,16 L218,58 Z" -> poligono
        nums, i = [], 0
        for p in partes[:-1]:
            p = p[1:] if p and p[0].upper() in "ML" else p
            if p:
                nums.append(float(p))
        if len(nums) < 6 or len(nums) % 2:
            raise ValueError(f"poligono malformado: {d!r}")
        return "poligono", [(nums[i], nums[i + 1]) for i in range(0, len(nums), 2)]
    if len(partes) == 3 and partes[0].startswith("M") and partes[2].startswith("H"):
        # "M58,200 H150" -> linha horizontal
        x1, y1, x2 = float(partes[0][1:]), float(partes[1]), float(partes[2][1:])
        return "linha", (x1, y1, x2, y1)
    raise ValueError(f"geometria nao suportada por este gerador: {d!r}")


def formas_do_logo(no):
    """Lista de formas de um <DrawingImage>, ja' com cor e (para linha) espessura."""
    formas = []
    for gd in no.iter(f"{{{NS['p']}}}GeometryDrawing"):
        preenchimento = gd.get("Brush")
        caneta = gd.find(f"{{{NS['p']}}}GeometryDrawing.Pen/{{{NS['p']}}}Pen")
        retangulo = gd.find(f"{{{NS['p']}}}GeometryDrawing.Geometry/{{{NS['p']}}}RectangleGeometry")
        if retangulo is not None:
            if preenchimento is None:
                raise ValueError("RectangleGeometry sem Brush")
            x, y, l, a = numeros(retangulo.get("Rect"))
            r = float(retangulo.get("RadiusX") or 0)
            formas.append(("retangulo", (x, y, l, a, r), cor(preenchimento), 0.0))
            continue
        d = gd.get("Geometry")
        if d is None:
            raise ValueError("GeometryDrawing sem Geometry nem RectangleGeometry")
        tipo, dados = caminho(d)
        if tipo == "poligono":
            if preenchimento is None:
                raise ValueError(f"poligono sem Brush: {d!r}")
            formas.append((tipo, dados, cor(preenchimento), 0.0))
        else:
            if caneta is None:
                raise ValueError(f"linha sem Pen: {d!r}")
            pincel = caneta.find(f"{{{NS['p']}}}Pen.Brush/{{{NS['p']}}}SolidColorBrush")
            if pincel is None:
                raise ValueError(f"Pen sem SolidColorBrush: {d!r}")
            formas.append((tipo, dados, cor(pincel.get("Color")), float(caneta.get("Thickness"))))
    if not formas:
        raise ValueError("logo sem nenhuma forma")
    return formas


def limites(formas):
    """Caixa que contem TODAS as formas, contando a espessura da caneta — e' a mesma definicao de
    Drawing.Bounds do WPF, entao o enquadramento do .ico nao muda se um dia o WPF voltar a rasterizar."""
    xs, ys = [], []
    for tipo, dados, _, espessura in formas:
        if tipo == "retangulo":
            x, y, l, a, _ = dados
            xs += [x, x + l]
            ys += [y, y + a]
        elif tipo == "poligono":
            xs += [p[0] for p in dados]
            ys += [p[1] for p in dados]
        else:
            x1, y1, x2, y2 = dados
            m = espessura / 2
            xs += [min(x1, x2) - m, max(x1, x2) + m]
            ys += [min(y1, y2) - m, max(y1, y2) + m]
    return min(xs), min(ys), max(xs), max(ys)


def desenhar(formas, n):
    x0, y0, x1, y1 = limites(formas)
    lado = n * SUPER
    escala = min(lado / (x1 - x0), lado / (y1 - y0))
    dx = (lado - (x1 - x0) * escala) / 2 - x0 * escala
    dy = (lado - (y1 - y0) * escala) / 2 - y0 * escala

    def t(x, y):
        return (x * escala + dx, y * escala + dy)

    img = Image.new("RGBA", (lado, lado), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    for tipo, dados, c, espessura in formas:
        if tipo == "retangulo":
            x, y, l, a, r = dados
            ax, ay = t(x, y)
            bx, by = t(x + l, y + a)
            if r > 0:
                d.rounded_rectangle([ax, ay, bx, by], radius=r * escala, fill=c)
            else:
                d.rectangle([ax, ay, bx, by], fill=c)
        elif tipo == "poligono":
            d.polygon([t(*p) for p in dados], fill=c)
        else:
            x1l, y1l, x2l, y2l = dados
            larg = espessura * escala
            d.line([t(x1l, y1l), t(x2l, y2l)], fill=c, width=round(larg))
            # Pontas arredondadas (StartLineCap/EndLineCap="Round" no Pen do XAML): o Pillow nao tem
            # cap, entao o circulo de raio = metade da espessura vai a mao nas duas pontas.
            for px, py in (t(x1l, y1l), t(x2l, y2l)):
                d.ellipse([px - larg / 2, py - larg / 2, px + larg / 2, py + larg / 2], fill=c)
    return img.resize((n, n), Image.LANCZOS)


def main():
    saida = pathlib.Path(sys.argv[1])
    saida.mkdir(parents=True, exist_ok=True)
    raiz = pathlib.Path(__file__).resolve().parents[2]
    tema = ET.parse(raiz / "src" / "mOffice.Comum" / "Temas" / "mOfficeTema.xaml").getroot()
    achados = {}
    for no in tema.iter(f"{{{NS['p']}}}DrawingImage"):
        chave = no.get(f"{{{NS['x']}}}Key")
        if chave in APPS:
            achados[chave] = formas_do_logo(no)
    faltando = set(APPS) - set(achados)
    if faltando:
        raise SystemExit(f"logos ausentes em mOfficeTema.xaml: {sorted(faltando)}")
    for chave, app in APPS.items():
        for n in TAMANHOS:
            destino = saida / f"{app}-{n}.png"
            # optimize=True: o mesmo PNG sempre, sem depender do palpite de compressao padrao.
            desenhar(achados[chave], n).save(destino, "PNG", optimize=True)
            print(destino)


if __name__ == "__main__":
    main()
