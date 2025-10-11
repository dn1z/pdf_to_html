import fontforge
import sys

pdf_file = sys.argv[1]
output_folder = sys.argv[2]

data = fontforge.fontsInFile(pdf_file)

for i, font in enumerate(data):
    print(f"Extracting font {i+1}/{len(data)}: {font}")
    fontforge.open(f"{pdf_file}({font})").generate(f"{output_folder}/{font}.woff")