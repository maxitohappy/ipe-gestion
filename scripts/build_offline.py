"""Arma descargas/ipe-gestion-sin-internet.zip con los archivos de la app.

La versión sin internet se abre con doble clic en index.html: no inicia sesión ni comparte el historial
(lo guarda solo en ese computador). En la app publicada, el enlace de descarga se muestra solo con sesión
iniciada. Se ejecuta antes de cada publicación:
    python scripts/build_offline.py
"""
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DST = ROOT / 'descargas' / 'ipe-gestion-sin-internet.zip'
INCLUIR = ['index.html', 'css', 'js', 'vendor']

DST.parent.mkdir(exist_ok=True)
with zipfile.ZipFile(DST, 'w', zipfile.ZIP_DEFLATED) as z:
    for nombre in INCLUIR:
        ruta = ROOT / nombre
        archivos = [ruta] if ruta.is_file() else sorted(p for p in ruta.rglob('*') if p.is_file())
        for f in archivos:
            z.write(f, Path('ipe-gestion') / f.relative_to(ROOT))
print(f'{DST.relative_to(ROOT)}: {DST.stat().st_size // 1024} KB')
