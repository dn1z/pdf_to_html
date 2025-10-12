import base64
import math
import os
import subprocess
import tempfile
from io import BytesIO, StringIO
from pathlib import Path

import pymupdf
from fontTools.ttLib import TTFont
from natsort import natsorted

CWD = Path.cwd()
BASE = Path(__file__).parent
TEMPDIR = None


class FontOps:
    BUFFERS = {}

    @staticmethod
    def generate_fonts(pdf_path):
        command = [
            'pdftohtml.exe',
            '-overwrite',
            Path(os.path.relpath(pdf_path, CWD)).as_posix(),
            Path(os.path.relpath(TEMPDIR, CWD)).as_posix()
        ]

        subprocess.run(command, check=True)

    @staticmethod
    def extract_fonts_from_pdf(pdf_path, doc):
        """Extract embedded fonts from PDF and return as base64."""
        fonts = {}
        FontOps.generate_fonts(pdf_path)

        data = {}
        for page in doc:
            for font in page.get_fonts():
                xref = font[0]
                font_name = font[3]
                if xref not in data:
                    data[xref] = font_name

        print(f"Found {len(data)} fonts in PDF")

        for ttf_file in natsorted(TEMPDIR.glob('*.ttf')):
            font = TTFont(ttf_file)
            font_family = font["name"].names[1].toStr()

            # sometimes pdftohtml do not include subset font name even though it is subset
            if '+' not in font_family:
                font_family = list(data.values())[0]

            font_xref = None
            for xref, name in data.items():
                if name == font_family:
                    font_xref = xref
                    data.pop(xref)
                    break
            
            FontOps.BUFFERS[font_xref] = pymupdf.Font(fontbuffer=ttf_file.read_bytes())

            font.flavor = "woff"
            woff_buffer = BytesIO()
            font.save(woff_buffer)

            woff_b64 = base64.b64encode(woff_buffer.getvalue()).decode('utf-8')

            fonts[font_xref] = {
                "name": font_family,
                "data": woff_b64,
                "xref": font_xref,
                "format": "woff"
            }

        for otf_file in natsorted(TEMPDIR.glob('*.otf')):
            font = TTFont(otf_file)

            # OTF fonts do not seem to have a name entry
            # TODO make sure otf matches correct extracted font
            font_family = None
            for xref, name in data.items():
                if '+' in name:
                    font_xref = xref
                    font_family = data.pop(xref)
                    break
                    
            FontOps.BUFFERS[font_xref] = pymupdf.Font(fontbuffer=otf_file.read_bytes())
            
            font.flavor = "woff"
            woff_buffer = BytesIO()
            font.save(woff_buffer)

            woff_b64 = base64.b64encode(woff_buffer.getvalue()).decode('utf-8')

            fonts[font_xref] = {
                "name": font_family,
                "data": woff_b64,
                "xref": font_xref,
                "format": "woff",
            }
        
        # Non-embedded fonts
        for xref, name in data.items():
            FontOps.BUFFERS[xref] = pymupdf.Font(fontname=name)
            fonts[xref] = {
                "name": name,
                "data": None,
                "xref": xref,
                "format": None
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

        page_fonts = page.get_fonts()
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

                        for font in page_fonts:
                            if font[3] == font_name:
                                font_xref = font[0]
                                break

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
                                calculated_width = font_buffers[font_xref].text_length(text, font_size)
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
                                "font_xref": font_xref,
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
    
    for _, font_info in fonts.items():
        if font_info["data"] is None:
            font_css_buffer.write(f"""
        @font-face {{
            font-family: 'f_{font_info["xref"]}';
            src: local('{font_info["name"]}');
        }}"""
            )
            continue
        
        font_css_buffer.write(f"""
        @font-face {{
            font-family: 'f_{font_info["xref"]}';
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
        }}{font_css}
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
                font-family:'f_{text_block['font_xref']}';font-size:{text_block['font_size']}px;
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
    fonts = FontOps.extract_fonts_from_pdf(pdf_path, doc)
    font_buffers = FontOps.BUFFERS
    print(f"Extracted {len(fonts)} fonts")

    if output_path is None:
        pdf_path = Path(pdf_path)
        name = pdf_path.stem
        output_path = Path(f"{pdf_path.parent/name}.html")
    
    pages_data = []
    total_pages = len(doc)
    for i, page in enumerate(doc):
        print(f"Processing page {i + 1}/{total_pages} ({(i + 1) / total_pages * 100:.2f}%)")
        
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