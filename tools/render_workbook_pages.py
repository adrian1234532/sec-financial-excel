"""Decode actual Excel print exports and bind every page to the workbook hash."""

import argparse
import hashlib
import json
import subprocess
from pathlib import Path

from PIL import Image, ImageDraw
from pypdf import PdfReader


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--workbook', type=Path, required=True)
    parser.add_argument('--directory', type=Path, required=True)
    args = parser.parse_args()
    args.directory.mkdir(parents=True, exist_ok=True)
    digest = hashlib.sha256(args.workbook.read_bytes()).hexdigest()
    records = []
    pdfs = sorted(args.directory.glob('*.pdf'))
    if not pdfs:
        raise FileNotFoundError(
            f'{args.directory}: no Excel print-export PDF files found')
    for pdf in pdfs:
        reader = PdfReader(pdf)
        if not reader.pages:
            raise ValueError(f'{pdf.name}: no print pages')
        subprocess.run(['pdftoppm', '-scale-to', '1800', '-png', str(pdf), str(pdf.with_suffix(''))], check=True)
        images = sorted(args.directory.glob(pdf.stem + '-*.png'))
        if len(images) != len(reader.pages):
            raise ValueError(f'{pdf.name}: image/page count mismatch')
        for number, (page, image) in enumerate(zip(reader.pages, images), 1):
            with Image.open(image) as decoded:
                decoded.load()
                if min(decoded.size) == 0:
                    raise ValueError(f'{image}: invalid image')
            records.append({'sheet': pdf.stem, 'page': number, 'image': image.name,
                            'text_characters': len(page.extract_text()), 'workbook_sha256': digest,
                            'image_sha256': hashlib.sha256(image.read_bytes()).hexdigest(), 'review': 'NOT_RUN'})
        tiles = Image.new('RGB', (2000, 790 * ((len(images) + 1) // 2)), 'white')
        draw = ImageDraw.Draw(tiles)
        for index, image in enumerate(images):
            with Image.open(image) as source:
                source.thumbnail((980, 750))
                x, y = (index % 2) * 1000, (index // 2) * 790
                tiles.paste(source, (x, y + 30))
                draw.text((x + 10, y + 8), f'{pdf.stem} page {index + 1}', fill='black')
        tiles.save(args.directory / (pdf.stem + '_contact.png'))
    (args.directory / 'pages.json').write_text(json.dumps(records, indent=2), encoding='utf-8')
    print(f'Rendered {len(records)} actual print pages; visual judgments remain NOT_RUN until viewed.')


if __name__ == '__main__':
    main()
