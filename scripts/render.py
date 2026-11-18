#!/usr/bin/env python3
"""Render a .pptx to one PNG per slide (and a contact sheet) for visual QA.

    python render.py deck.pptx out_dir [--width 1600] [--slides 2,5,7] [--sheet] [--engine auto|powerpoint|libreoffice]

Engines
  powerpoint   Windows + Microsoft PowerPoint, driven through PowerShell COM. Pixel-exact with what
               the audience will see, including the Noto fonts. Preferred whenever available.
  libreoffice  soffice --headless → PDF → PNG (PyMuPDF, or pdftoppm). Works on Linux/macOS; fonts
               must be installed (see assets/fontconfig/README.md), text widths may differ slightly.
"""
from __future__ import annotations

import argparse
import glob
import os
import shutil
import subprocess
import sys
import tempfile

PS_SCRIPT = r'''
param([string]$Pptx, [string]$OutDir, [int]$Width = 1600, [string]$Slides = "")
$ErrorActionPreference = "Stop"
New-Item -ItemType Directory -Force $OutDir | Out-Null
$pp = New-Object -ComObject PowerPoint.Application
try {
  $pres = $pp.Presentations.Open($Pptx, -1, 0, 0)
  $h = [int]($Width * $pres.PageSetup.SlideHeight / $pres.PageSetup.SlideWidth)
  $list = if ($Slides) { $Slides.Split(",") | ForEach-Object { [int]$_ } } else { 1..$pres.Slides.Count }
  foreach ($i in $list) { $pres.Slides.Item($i).Export((Join-Path $OutDir ("s{0:D2}.png" -f $i)), "PNG", $Width, $h) }
  $pres.Close()
} finally { $pp.Quit() }
'''


def render_powerpoint(pptx, out_dir, width, slides):
    if os.name != 'nt' or not shutil.which('powershell'):
        raise RuntimeError('PowerPoint engine needs Windows with PowerShell')
    with tempfile.TemporaryDirectory() as td:
        ps1 = os.path.join(td, 'render.ps1')
        with open(ps1, 'w', encoding='utf-8-sig') as f:
            f.write(PS_SCRIPT)
        cmd = ['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ps1,
               '-Pptx', os.path.abspath(pptx), '-OutDir', os.path.abspath(out_dir), '-Width', str(width)]
        if slides:
            cmd += ['-Slides', ','.join(map(str, slides))]
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=900)
        if r.returncode != 0:
            raise RuntimeError(r.stderr.strip() or r.stdout.strip())


def _soffice():
    for c in ('soffice', 'libreoffice', r'C:\Program Files\LibreOffice\program\soffice.exe',
              '/Applications/LibreOffice.app/Contents/MacOS/soffice'):
        p = shutil.which(c) or (c if os.path.exists(c) else None)
        if p:
            return p
    raise RuntimeError('LibreOffice (soffice) not found')


def render_libreoffice(pptx, out_dir, width, slides):
    with tempfile.TemporaryDirectory() as td:
        subprocess.run([_soffice(), '--headless', '--convert-to', 'pdf', '--outdir', td, os.path.abspath(pptx)],
                       check=True, capture_output=True, timeout=900)
        pdf = os.path.join(td, os.path.splitext(os.path.basename(pptx))[0] + '.pdf')
        try:
            import fitz  # PyMuPDF
            doc = fitz.open(pdf)
            for i, page in enumerate(doc, 1):
                if slides and i not in slides:
                    continue
                zoom = width / page.rect.width
                page.get_pixmap(matrix=fitz.Matrix(zoom, zoom)).save(os.path.join(out_dir, f's{i:02d}.png'))
        except ImportError:
            if not shutil.which('pdftoppm'):
                raise RuntimeError('install PyMuPDF (pip install pymupdf) or poppler (pdftoppm)')
            subprocess.run(['pdftoppm', '-png', '-scale-to-x', str(width), '-scale-to-y', '-1', pdf,
                            os.path.join(out_dir, 's')], check=True)
            for f in glob.glob(os.path.join(out_dir, 's-*.png')):
                n = int(os.path.basename(f)[2:-4])
                os.replace(f, os.path.join(out_dir, f's{n:02d}.png'))


def contact_sheet(out_dir, cols=4, tile=480):
    from PIL import Image, ImageDraw
    files = sorted(glob.glob(os.path.join(out_dir, 's[0-9]*.png')))
    if not files:
        return None
    ims = [Image.open(f).convert('RGB') for f in files]
    th = int(tile * ims[0].height / ims[0].width)
    gap, lab = 12, 22
    rows = (len(ims) + cols - 1) // cols
    sheet = Image.new('RGB', (cols * tile + (cols + 1) * gap, rows * (th + lab + gap) + gap), (232, 232, 232))
    d = ImageDraw.Draw(sheet)
    for k, (f, im) in enumerate(zip(files, ims)):
        r, c = divmod(k, cols)
        x, y = gap + c * (tile + gap), gap + r * (th + lab + gap)
        d.text((x, y + 4), os.path.basename(f)[1:3], fill=(60, 60, 60))
        sheet.paste(im.resize((tile, th)), (x, y + lab))
    path = os.path.join(out_dir, 'sheet.jpg')
    sheet.save(path, quality=88)
    return path


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('pptx')
    ap.add_argument('out_dir')
    ap.add_argument('--width', type=int, default=1600)
    ap.add_argument('--slides', help='comma separated 1-based slide numbers')
    ap.add_argument('--sheet', action='store_true', help='also write sheet.jpg (all slides in a grid)')
    ap.add_argument('--engine', default='auto', choices=['auto', 'powerpoint', 'libreoffice'])
    a = ap.parse_args()
    os.makedirs(a.out_dir, exist_ok=True)
    for f in glob.glob(os.path.join(a.out_dir, 's[0-9]*.png')):
        os.remove(f)
    slides = [int(x) for x in a.slides.split(',')] if a.slides else None
    engines = [a.engine] if a.engine != 'auto' else ['powerpoint', 'libreoffice']
    errors = []
    for eng in engines:
        try:
            (render_powerpoint if eng == 'powerpoint' else render_libreoffice)(a.pptx, a.out_dir, a.width, slides)
            print(f'engine: {eng}')
            break
        except Exception as e:  # try the next engine
            errors.append(f'{eng}: {e}')
    else:
        sys.exit('render failed\n' + '\n'.join(errors))
    for f in sorted(glob.glob(os.path.join(a.out_dir, 's[0-9]*.png'))):
        print(os.path.abspath(f))
    if a.sheet:
        print('sheet:', os.path.abspath(contact_sheet(a.out_dir)))


if __name__ == '__main__':
    main()
