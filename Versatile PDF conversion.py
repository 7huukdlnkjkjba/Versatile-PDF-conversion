import argparse
import sys
from pathlib import Path
from PIL import Image
import pytesseract
import subprocess
from pdf2docx import Converter
from pdf2image import convert_from_path
import PyPDF2
import pandas as pd
from pptx import Presentation
from pptx.util import Inches
import fitz  # PyMuPDF


# ========== PDF转换模块 ==========
def pdf_to_word(input_pdf, output_docx):
    """PDF转Word"""
    cv = Converter(input_pdf)
    cv.convert(output_docx)
    cv.close()


def pdf_to_excel(input_pdf, output_xlsx, page=0):
    """PDF转Excel（提取表格）"""
    tables = tabula.read_pdf(input_pdf, pages=page + 1)
    df = tables[0]
    df.to_excel(output_xlsx, index=False)


def pdf_to_ppt(input_pdf, output_pptx, dpi=200):
    """PDF转PPT（每页转为图片插入）"""
    images = convert_from_path(input_pdf, dpi=dpi)
    prs = Presentation()

    for img in images:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        img_path = f"temp_{hash(img)}.png"
        img.save(img_path)
        slide.shapes.add_picture(img_path, 0, 0, prs.slide_width, prs.slide_height)

    prs.save(output_pptx)


def pdf_to_images(input_pdf, output_dir, format='png', dpi=300):
    """PDF转图片"""
    images = convert_from_path(input_pdf, dpi=dpi, output_folder=output_dir,
                               fmt=format, output_file=Path(input_pdf).stem)


def pdf_to_txt(input_pdf, output_txt):
    """PDF转TXT"""
    with open(input_pdf, 'rb') as f:
        reader = PyPDF2.PdfReader(f)
        text = "\n".join([page.extract_text() for page in reader.pages])

    with open(output_txt, 'w', encoding='utf-8') as f:
        f.write(text)


def pdf_to_cad(input_pdf, output_cad, target_format='dwg'):
    """PDF转CAD（需要Inkscape）"""
    subprocess.run([
        "inkscape",
        input_pdf,
        "--export-filename", output_cad,
        "--export-type", target_format
    ], check=True)


# ========== 图片/CAD模块 ==========
def img_to_pdf(input_img, output_pdf):
    """图片转PDF"""
    image = Image.open(input_img)
    image.save(output_pdf, "PDF", resolution=100.0)


def img_to_text(input_img, output_txt):
    """图片文字提取"""
    text = pytesseract.image_to_string(Image.open(input_img))
    with open(output_txt, 'w') as f:
        f.write(text)


def cad_to_pdf(input_cad, output_pdf):
    """CAD转PDF（需要LibreCAD）"""
    subprocess.run([
        "libreCAD",
        "--export-to", "pdf",
        "--output", output_pdf,
        input_cad
    ], check=True)


# ========== 命令行接口 ==========
def main():
    parser = argparse.ArgumentParser(description="超级文件转换工具")
    subparsers = parser.add_subparsers(dest="command", required=True)

    # PDF转换命令
    pdf_cmds = {
        'pdf2word': (pdf_to_word, 'docx'),
        'pdf2excel': (pdf_to_excel, 'xlsx'),
        'pdf2ppt': (pdf_to_ppt, 'pptx'),
        'pdf2img': (pdf_to_images, ''),
        'pdf2txt': (pdf_to_txt, 'txt'),
        'pdf2cad': (pdf_to_cad, 'dwg')
    }

    for cmd, (func, ext) in pdf_cmds.items():
        p = subparsers.add_parser(cmd)
        p.add_argument('-i', '--input', required=True)
        p.add_argument('-o', '--output', required=True)
        if cmd == 'pdf2excel':
            p.add_argument('-p', '--page', type=int, default=0)
        if cmd == 'pdf2img':
            p.add_argument('-f', '--format', default='png')
            p.add_argument('-d', '--dpi', type=int, default=300)

    # 图片/CAD命令
    img = subparsers.add_parser('img2pdf')
    img.add_argument('-i', '--input', required=True)
    img.add_argument('-o', '--output', required=True)

    ocr = subparsers.add_parser('img2txt')
    ocr.add_argument('-i', '--input', required=True)
    ocr.add_argument('-o', '--output', required=True)

    cad = subparsers.add_parser('cad2pdf')
    cad.add_argument('-i', '--input', required=True)
    cad.add_argument('-o', '--output', required=True)

    args = parser.parse_args()

    try:
        if args.command.startswith('pdf2'):
            if args.command == 'pdf2excel':
                pdf_to_excel(args.input, args.output, args.page)
            elif args.command == 'pdf2img':
                pdf_to_images(args.input, args.output, args.format, args.dpi)
            else:
                globals()[pdf_cmds[args.command][0].__name__](args.input, args.output)
        elif args.command == 'img2pdf':
            img_to_pdf(args.input, args.output)
        elif args.command == 'img2txt':
            img_to_text(args.input, args.output)
        elif args.command == 'cad2pdf':
            cad_to_pdf(args.input, args.output)

        print("转换成功！")
    except Exception as e:
        print(f"错误: {str(e)}")
        sys.exit(1)


if __name__ == "__main__":
    main()