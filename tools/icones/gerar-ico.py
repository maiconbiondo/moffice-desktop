"""Empacota <pasta>/<app>-{256,48,32,16}.png num .ico de 4 quadros PNG (mesmo formato do mpdf.ico).

Deterministico: so' concatena os bytes dos PNG ja' gravados, sem recomprimir nada. Uso:
    python tools/icones/gerar-ico.py <pasta-dos-png> <app> <arquivo.ico-de-saida>
"""
import struct, sys, pathlib

pasta, app, saida = pathlib.Path(sys.argv[1]), sys.argv[2], pathlib.Path(sys.argv[3])
tamanhos = [256, 48, 32, 16]
pngs = [(pasta / f"{app}-{n}.png").read_bytes() for n in tamanhos]
# ICONDIR: reservado=0, tipo=1 (icone), quantidade de quadros.
cab = struct.pack("<HHH", 0, 1, len(tamanhos))
offset = 6 + 16 * len(tamanhos)
entradas = b""
for n, png in zip(tamanhos, pngs):
    # ICONDIRENTRY: largura, altura (0 = 256), cores da paleta, reservado, planos, bits por pixel,
    # tamanho do quadro, offset do quadro.
    entradas += struct.pack("<BBBBHHII", 0 if n == 256 else n, 0 if n == 256 else n, 0, 0, 1, 32, len(png), offset)
    offset += len(png)
saida.parent.mkdir(parents=True, exist_ok=True)
saida.write_bytes(cab + entradas + b"".join(pngs))
print(f"{saida} ({saida.stat().st_size} bytes)")
