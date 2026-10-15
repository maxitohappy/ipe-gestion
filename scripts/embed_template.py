"""Incrusta templates/memo_template.docx en js/plantilla.js (base64).

Así la app no necesita descargar la plantilla y funciona también abriendo index.html
directamente desde el computador, sin internet (los navegadores bloquean fetch en file://).

Ejecutar cada vez que cambie la plantilla:
    python scripts/embed_template.py
"""
import base64
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / 'templates' / 'memo_template.docx'
DST = ROOT / 'js' / 'plantilla.js'

b64 = base64.b64encode(SRC.read_bytes()).decode('ascii')
DST.write_text(
    '// Generado por scripts/embed_template.py a partir de templates/memo_template.docx. No editar a mano.\n'
    f'(self.IPE = self.IPE || {{}}).plantillaB64 = "{b64}";\n',
    encoding='utf-8',
)
print(f'{DST.relative_to(ROOT)}: {len(b64)} caracteres')
