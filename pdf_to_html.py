import base64
import math
import os
import subprocess
import tempfile
from io import BytesIO, StringIO
from pathlib import Path

import pymupdf
from fontTools.ttLib import TTFont

CWD = Path.cwd()
BASE = Path(__file__).parent
TEMPDIR = None

# You need fontforge installed and accessible in your PATH
# On Windows, it checks from WSL
class FontOps:
    @staticmethod
    def generate_fonts(pdf_path):
        for font in TEMPDIR.glob('*'):
            font.unlink()
        
        command = [
            'wsl',
            'fontforge',
            '-script',
            Path(os.path.relpath(BASE/'fontforge.py', CWD)).as_posix(),
            Path(os.path.relpath(pdf_path, CWD)).as_posix(),
            Path(os.path.relpath(TEMPDIR, CWD)).as_posix()
        ]

        if os.name != 'nt':
            command.pop(0)

        subprocess.run(command, check=True)
    
    @staticmethod
    def woff_to_mupdf():
        """ Convert WOFF font file to pymupdf.Font object """
        fonts = {}

        for woff_file in TEMPDIR.glob('*.woff'):

            ttf_buffer = BytesIO()
            font = TTFont(woff_file)
            font.flavor = None  # Remove WOFF flavor to convert to TTF
            font.save(ttf_buffer)
            ttf_buffer.seek(0)
            try:
                fonts[woff_file.stem] = pymupdf.Font(fontbuffer=ttf_buffer.read())
            except Exception as e:
                print(f"Error processing {woff_file}: {e}")

        return fonts

    @staticmethod
    def extract_fonts_from_pdf(pdf_path):
        """Extract embedded fonts from PDF and return as base64."""
        fonts = {}
        FontOps.generate_fonts(pdf_path)

        woff_files = list(TEMPDIR.glob('*.woff'))
        for woff_file in woff_files:
            with open(woff_file, 'rb') as f:
                font_data = f.read()
            font_base64 = base64.b64encode(font_data).decode('utf-8')
            clean_name = woff_file.stem

            fonts[clean_name] = {
                "name": clean_name,
                "data": font_base64,
                "format": "woff"
            }

        return fonts


class Extractors:
    @staticmethod
    def extract_page_image(page, scale_factor):
        """ Convert page to a PNG image and return as base64 string along with dimensions """
        page.add_redact_annot(page.rect)
        page.apply_redactions(
                images=pymupdf.PDF_REDACT_IMAGE_NONE,     # Keep all images
                graphics=pymupdf.PDF_REDACT_LINE_ART_NONE # Keep all vector graphics/drawings
        )

        mat = pymupdf.Matrix(scale_factor, scale_factor)
        
        pix = page.get_pixmap(matrix=mat, alpha=False)
        img_data = pix.tobytes("png")
        img_base64 = base64.b64encode(img_data).decode('utf-8')

        width = int(pix.width)
        height = int(pix.height)
        return img_base64, width, height

    @staticmethod
    def extract_text_blocks(page, font_buffers, scale_factor):
        """ Extract text blocks with coordinates and font information """
        text_blocks = []

        text_dict = page.get_text("dict")    
        for block in text_dict["blocks"]:
            if "lines" in block:  # Text block
                for line in block["lines"]:
                    rotation = line["dir"]
                    degrees = int(math.degrees(math.atan2(*rotation)))
                    html_degrees = 360 - (degrees - 90)
                    if html_degrees >= 360:
                        html_degrees -= 360

                    for span in line["spans"]:
                        font_name = span["font"]
                        font_size = round(span["size"] * scale_factor, 2)
                        font_flags = span["flags"]

                        color = span["color"]
                        color = f"#{color:06x}"  # Convert to hex

                        font_weight = "bold" if font_flags & 2**4 else "normal"
                        font_style = "italic" if font_flags & 2**1 else "normal"
                        
                        text = span.get("text", "")
                        if text:
                            bbox = span["bbox"]
                            x = round(bbox[0] * scale_factor, 2)
                            y = round(bbox[1] * scale_factor, 2)

                            width = round((bbox[2] - bbox[0]) * scale_factor, 2)

                            if html_degrees != 0:
                                if html_degrees in [90, 270]:
                                    width = abs(round((bbox[3] - bbox[1]) * scale_factor, 2))
                                elif html_degrees != 180:
                                    width = abs(width / math.cos(math.radians(html_degrees)))

                            try:
                                calculated_width = font_buffers[font_name].text_length(text, font_size)
                            except Exception as e:
                                print(f"Error calculating text width for font {font_name}: {e}")
                                calculated_width = width

                            diff = width - calculated_width

                            letter_spacing = 0
                            if abs(diff) > 1.0:
                                letter_spacing = round(diff / len(text), 2)
                            
                            if html_degrees != 0:
                                x = round(span["origin"][0] * scale_factor, 2)
                                y = round(span["origin"][1] * scale_factor, 2)

                            text_blocks.append({
                                "text": text,
                                "x": x,
                                "y": y,
                                "font_name": font_name,
                                "font_size": font_size,
                                "font_weight": font_weight,
                                "font_style": font_style,
                                "color": color,
                                "width": width,
                                "letter_spacing": letter_spacing,
                                "rotation": html_degrees
                            })

        return text_blocks


def generate_html(pages_data, fonts, title):
    """Generate complete HTML with embedded images, fonts, and positioned text."""
    
    html_buffer = StringIO()
    
    font_css_buffer = StringIO()
    for font_name, font_info in fonts.items():
        font_css_buffer.write(f"""
            @font-face {{
                font-family: '{font_name}';
                src: url(data:font/truetype;base64,{font_info["data"]}) format('{font_info["format"]}');
            }}"""
        )
    font_css = font_css_buffer.getvalue()
    font_css_buffer.close()

    html_buffer.write(
f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{title}</title>
    <style>
        body {{
            margin: 0;
            padding: 20px;
            background-color: #f0f0f0;
        }}
        .page {{
            position: relative;
            margin: 20px auto;
            box-shadow: 0 4px 8px rgba(0,0,0,0.1);
            page-break-after: always;
        }}
        .page-background {{
            width: 100%;
            height: 100%;
            background-repeat: no-repeat;
            background-size: contain;
            background-position: top left;
        }}
        ._l {{
            position: absolute;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
        }}
        ._t {{
            position: absolute;
            white-space: pre;
        }}
        @media print {{
            body {{ margin: 0; padding: 0; background-color: white; }}
            .page {{ margin: 0; box-shadow: none; page-break-after: always; }}
            @page {{ size: A3; margin: 0; }}
        }}
    {font_css}
    </style>
</head>
<body>""")

    for page_data in pages_data:
        html_buffer.write(f"""
    <div class="page" style="width: {page_data['width']}px; height: {page_data['height']}px;">
        <div class="page-background" style="background-image: url(data:image/png;base64,{page_data['image']});"></div>
        <div class="_l">"""
        )

        for text_block in page_data['text_blocks']:
            rotation_style = ""
            if text_block['rotation'] != 0:
                # translateY ?!
                rotation_style = f"transform: rotate({text_block['rotation']}deg) translateY(-75%);transform-origin: left top;"

            html_buffer.write(f"""
            <span class="_t" style="left:{text_block['x']}px;top:{text_block['y']}px;
                font-family:'{text_block['font_name']}';font-size:{text_block['font_size']}px;
                color:{text_block['color']};letter-spacing:{text_block['letter_spacing']}px;
                {rotation_style}
            ">{text_block['text']}</span>"""
            )
        
        html_buffer.write("""
        </div>
    </div>"""
        )
    
    html_buffer.write("""\n</body>\n</html>""")
    
    html_content = html_buffer.getvalue()
    html_buffer.close()
    
    return html_content


def pdf_to_html(pdf_path, output_path=None, scale_factor=2.0):
    """Convert PDF to self-contained HTML file."""
    pymupdf.TOOLS.set_subset_fontnames(True)
    pymupdf.TOOLS.mupdf_display_errors(False)
    
    global TEMPDIR
    temp_dir = tempfile.TemporaryDirectory()
    TEMPDIR = Path(temp_dir.name)
    print(f"Using temporary directory: {TEMPDIR}")

    doc = pymupdf.open(pdf_path)
    print(f"Processing PDF: {pdf_path}")
    print(f"Total pages: {len(doc)}")
    
    print("Extracting fonts...")
    fonts = FontOps.extract_fonts_from_pdf(pdf_path)
    font_buffers = FontOps.woff_to_mupdf()
    print(f"Extracted {len(fonts)} fonts")

    if output_path is None:
        pdf_path = Path(pdf_path)
        name = pdf_path.stem
        output_path = Path(f"{pdf_path.parent/name}.html")
    
    pages_data = []
    for i, page in enumerate(doc):
        print(f"Processing page {i + 1}/{len(doc)}")
        
        text_blocks = Extractors.extract_text_blocks(page, font_buffers, scale_factor)
        image_base64, width, height = Extractors.extract_page_image(page, scale_factor)
        
        pages_data.append({
            "image": image_base64,
            "width": width,
            "height": height,
            "text_blocks": text_blocks
        })
    
    print("Generating HTML...")
    html_content = generate_html(pages_data, fonts, output_path.stem)
    
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    print(f"HTML file saved: {output_path}")
    
    doc.close()
    temp_dir.cleanup()


def main():
    from tkinter.filedialog import askopenfilenames
    pdfs = askopenfilenames(title="Select PDF files", filetypes=[("PDF files", "*.pdf")])

    for pdf_file in pdfs:
        pdf_to_html(pdf_file)


if __name__ == "__main__":
    main()