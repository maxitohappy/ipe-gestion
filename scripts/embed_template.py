"""Incrusta las plantillas Word en js/plantilla.js (base64).

Así la app no necesita descargar las plantillas y funciona también abriendo index.html
directamente desde el computador, sin internet (los navegadores bloquean fetch en file://).

Ejecutar cada vez que cambie una plantilla:
    python scripts/embed_template.py
"""
import base64
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
PLANTILLAS = {
    'plantillaB64': ROOT / 'templates' / 'memo_template.docx',
    'plantillaDevolucionB64': ROOT / 'templates' / 'devolucion_template.docx',
}
DST = ROOT / 'js' / 'plantilla.js'

lineas = ['// Generado por scripts/embed_template.py a partir de templates/*.docx. No editar a mano.',
          'self.IPE = self.IPE || {};']
for nombre, ruta in PLANTILLAS.items():
    b64 = base64.b64encode(ruta.read_bytes()).decode('ascii')
    lineas.append(f'self.IPE.{nombre} = "{b64}";')
    print(f'{ruta.name}: {len(b64)} caracteres')
DST.write_text('\n'.join(lineas) + '\n', encoding='utf-8')
