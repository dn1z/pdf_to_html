import base64
import math
from io import BytesIO, StringIO
from pathlib import Path

import pymupdf

class Extractors:
    @staticmethod
    def extract_page_image(page, scale_factor):
        """ Convert full page to a WEBP image and return as base64 string along with dimensions """
        mat = pymupdf.Matrix(scale_factor, scale_factor)
        pix = page.get_pixmap(matrix=mat, alpha=False)

        img_buffer = BytesIO()
        pix.pil_save(img_buffer, format="WEBP", optimize=True, quality=90)
        img_base64 = base64.b64encode(img_buffer.getvalue()).decode('utf-8')

        width = int(pix.width)
        height = int(pix.height)
        return img_base64, width, height

    @staticmethod
    def extract_text_blocks(page, scale_factor):
        """ Extract text blocks with coordinates """
        text_lines = []

        text_dict = page.get_text("dict")    
        for block in text_dict["blocks"]:
            if "lines" in block:  # Text block
                for line in block["lines"]:
                    dir_x, dir_y = line["dir"]
                    rotation = line["dir"]
                    degrees = int(math.degrees(math.atan2(*rotation)))
                    html_degrees = 360 - (degrees - 90)
                    if html_degrees >= 360:
                        html_degrees -= 360

                    page.set_rotation(0)
                    
                    if not line["spans"]:
                        continue
                        
                    first_span = line["spans"][0]
                    line_origin_x = first_span["origin"][0] * scale_factor
                    line_origin_y = first_span["origin"][1] * scale_factor

                    line_spans = []
                    for span in line["spans"]:
                        text = span.get("text", "")
                        if text:
                            font_size = span["size"] * scale_factor
                            bbox = span["bbox"]
                            
                            bbox_w = (bbox[2] - bbox[0]) * scale_factor
                            bbox_h = (bbox[3] - bbox[1]) * scale_factor

                            text_length = bbox_w
                            if html_degrees in [90, 270]:
                                text_length = bbox_h
                            elif html_degrees not in [0, 180]:
                                try:
                                    text_length = abs(bbox_w / math.cos(math.radians(html_degrees)))
                                except:
                                    pass
                            
                            if text_length <= 0.1: 
                                text_length = 0.1

                            span_origin_x = span["origin"][0] * scale_factor
                            span_origin_y = span["origin"][1] * scale_factor

                            # Project span origin to unrotated local coordinate space of the line
                            dx = span_origin_x - line_origin_x
                            dy = span_origin_y - line_origin_y

                            proj_x = dx * dir_x + dy * dir_y
                            proj_y = -dx * dir_y + dy * dir_x

                            local_x = line_origin_x + proj_x
                            local_y = line_origin_y + proj_y

                            line_spans.append({
                                "text": text,
                                "font_size": round(font_size, 2),
                                "text_length": round(text_length, 2),
                                "local_x": round(local_x, 2),
                                "local_y": round(local_y, 2)
                            })

                    if line_spans:
                        text_lines.append({
                            "rotation": html_degrees,
                            "origin_x": round(line_origin_x, 2),
                            "origin_y": round(line_origin_y, 2),
                            "spans": line_spans
                        })

        return text_lines


def generate_html(pages_data, title):
    """Generate complete HTML with embedded full-page images and invisible selectable text."""
    
    html_buffer = StringIO()

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
            color: transparent;
            fill: transparent;
            cursor: text;
            white-space: pre;
        }}
        ._t::selection {{
            background: rgba(0, 120, 215, 0.3);
            color: transparent;
            fill: transparent;
        }}
        @media print {{
            body {{ margin: 0; padding: 0; background-color: white; }}
            .page {{ margin: 0; box-shadow: none; page-break-after: always; }}
            @page {{ size: A3; margin: 0; }}
        }}
    </style>
</head>
<body>""")

    for page_data in pages_data:
        html_buffer.write(f"""
    <div class="page" style="width: {page_data['width']}px; height: {page_data['height']}px;transform:rotate({page_data['rotation']}deg)">
        <div class="page-background" style="background-image: url(data:image/webp;base64,{page_data['image']});"></div>
        <svg class="_l" viewBox="0 0 {page_data['width']} {page_data['height']}" preserveAspectRatio="none">"""
        )

        for line_data in page_data['text_lines']:
            transform_attr = ""
            if line_data['rotation'] != 0:
                transform_attr = f"""transform="rotate({line_data['rotation']}, {line_data['origin_x']}, {line_data['origin_y']})" """

            html_buffer.write(f"""\n            <text class="_t" {transform_attr}>""")
            for span in line_data['spans']:
                html_buffer.write(f"""<tspan x="{span['local_x']}" y="{span['local_y']}" font-size="{span['font_size']}px" textLength="{span['text_length']}" lengthAdjust="spacingAndGlyphs">{span['text']}</tspan>""")
            html_buffer.write("</text>")
        
        html_buffer.write("""
        </svg>
    </div>"""
        )
    
    html_buffer.write("""\n</body>\n</html>""")
    
    html_content = html_buffer.getvalue()
    html_buffer.close()
    
    return html_content


def pdf_to_html(pdf_path, output_path=None, scale_factor=2):
    """Convert PDF to self-contained HTML file."""
    pymupdf.TOOLS.mupdf_display_errors(False)

    doc = pymupdf.open(pdf_path)
    print(f"Processing PDF: {pdf_path}")
    print(f"Total pages: {len(doc)}")
    
    if output_path is None:
        pdf_path = Path(pdf_path)
        name = pdf_path.stem
        output_path = Path(f"{pdf_path.parent/name}.html")
    
    pages_data = []
    total_pages = len(doc)
    for i, page in enumerate(doc):
        print(f"Processing page {i + 1}/{total_pages} ({(i + 1) / total_pages * 100:.2f}%)")
        
        rotation = page.rotation
        text_lines = Extractors.extract_text_blocks(page, scale_factor)
        image_base64, width, height = Extractors.extract_page_image(page, scale_factor)
        
        pages_data.append({
            "image": image_base64,
            "width": width,
            "height": height,
            "rotation": rotation,
            "text_lines": text_lines
        })
    
    print("Generating HTML...")
    html_content = generate_html(pages_data, output_path.stem)
    
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    print(f"HTML file saved: {output_path}")
    
    doc.close()


def main():
    import sys
    if len(sys.argv) == 2:
        pdfs = [sys.argv[1]]
    else:
        from tkinter.filedialog import askopenfilenames
        pdfs = askopenfilenames(title="Select PDF files", filetypes=[("PDF files", "*.pdf")])

    for pdf_file in pdfs:
        pdf_to_html(pdf_file)


if __name__ == "__main__":
    main()
