"""
Google Docs View-Only Downloader to PDF & High-Fidelity DOCX
Tải toàn bộ tài liệu Google Docs bị chặn tải/in (View-Only / No-Download) sang file PDF và DOCX (Word) chuẩn 100% bản gốc.
"""

import asyncio
import os
import sys
import argparse
import base64
import re
import io
import requests
from bs4 import BeautifulSoup
import docx
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from playwright.async_api import async_playwright
import pymupdf
from PIL import Image

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')


def extract_doc_id(url: str) -> str:
    match = re.search(r"/d/([a-zA-Z0-9-_]+)", url)
    return match.group(1) if match else url


def sanitize_filename(name: str) -> str:
    name = re.sub(r'[\\/*?:"<>|]', "", name)
    name = name.strip()
    return name if name else "GoogleDoc_Export"


def set_cell_border(cell, **kwargs):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = tcPr.first_child_found_in("w:tcBorders")
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)
    
    for edge in ('top', 'left', 'bottom', 'right'):
        edge_data = kwargs.get(edge)
        if edge_data:
            tag = f'w:{edge}'
            element = tcBorders.find(qn(tag))
            if element is None:
                element = OxmlElement(tag)
                tcBorders.append(element)
            for key in ["val", "color", "sz", "space"]:
                if key in edge_data:
                    element.set(qn(f'w:{key}'), str(edge_data[key]))


def set_cell_margins(cell, top=50, bottom=50, left=80, right=80):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for edge, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{edge}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)


def parse_css_style(style_str: str) -> dict:
    styles = {}
    if not style_str:
        return styles
    for part in style_str.split(';'):
        if ':' in part:
            k, v = part.split(':', 1)
            styles[k.strip().lower()] = v.strip().lower()
    return styles


def parse_pt(val_str: str) -> float:
    if not val_str:
        return 0.0
    m = re.search(r'([\d.]+)\s*pt', val_str)
    if m:
        return float(m.group(1))
    m = re.search(r'([\d.]+)\s*px', val_str)
    if m:
        return float(m.group(1)) * 0.75
    return 0.0


def parse_color(color_str: str):
    if not color_str:
        return None
    if color_str.startswith('#'):
        hex_val = color_str.lstrip('#')
        if len(hex_val) == 6:
            return RGBColor(int(hex_val[0:2], 16), int(hex_val[2:4], 16), int(hex_val[4:6], 16))
    elif color_str.startswith('rgb'):
        nums = re.findall(r'\d+', color_str)
        if len(nums) >= 3:
            return RGBColor(int(nums[0]), int(nums[1]), int(nums[2]))
    return None


def download_google_doc_to_docx(url: str, output_path: str = None) -> str:
    """Tải và tạo file .docx chuẩn 100% bản gốc (Times New Roman 13pt, viền bảng đen, căn lề chuẩn)"""
    doc_id = extract_doc_id(url)
    mobile_url = f"https://docs.google.com/document/d/{doc_id}/mobilebasic"

    print("\n" + "=" * 60)
    print("📝 BẮT ĐẦU TẢI DƯỚI ĐỊNH DẠNG WORD (.DOCX) CHUẨN BẢN GỐC")
    print(f"🔗 Mobile Endpoint: {mobile_url}")
    print("=" * 60)

    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    }
    r = requests.get(mobile_url, headers=headers)
    if r.status_code != 200:
        print(f"❌ Không thể truy cập {mobile_url} (HTTP {r.status_code})")
        return None

    soup = BeautifulSoup(r.text, 'html.parser')
    raw_title = soup.title.string if soup.title else "GoogleDoc"
    doc_title = raw_title.replace(" - Google Docs", "").replace(" - Google Tài liệu", "").strip()
    doc_title = sanitize_filename(doc_title)

    if not output_path:
        output_path = f"{doc_title}.docx"
    elif not output_path.lower().endswith(".docx"):
        output_path += ".docx"

    doc_content = soup.find('div', class_='doc-content') or soup.find('body')
    doc = docx.Document()

    # 1. Cấu hình lề trang (Page Margins) chuẩn A4
    for section in doc.sections:
        section.page_width = Cm(21.0)
        section.page_height = Cm(29.7)
        section.top_margin = Cm(2.0)
        section.bottom_margin = Cm(2.0)
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(2.0)

    # 2. Cấu hình Style Normal mặc định
    style_normal = doc.styles['Normal']
    font_normal = style_normal.font
    font_normal.name = 'Times New Roman'
    font_normal.size = Pt(13)
    font_normal.color.rgb = RGBColor(0, 0, 0)
    style_normal.paragraph_format.line_spacing = 1.15
    style_normal.paragraph_format.space_before = Pt(0)
    style_normal.paragraph_format.space_after = Pt(2)

    session = requests.Session()
    session.headers.update({'User-Agent': 'Mozilla/5.0'})

    def process_paragraph(p_tag, p_obj, is_table_cell=False):
        p_style = parse_css_style(p_tag.get('style', ''))
        
        # Căn lề
        text_align = p_style.get('text-align', '')
        if 'center' in text_align:
            p_obj.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif 'right' in text_align:
            p_obj.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        elif 'justify' in text_align:
            p_obj.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        else:
            p_obj.alignment = WD_ALIGN_PARAGRAPH.LEFT

        line_height = p_style.get('line-height', '')
        if line_height:
            try:
                p_obj.paragraph_format.line_spacing = float(line_height)
            except ValueError:
                p_obj.paragraph_format.line_spacing = 1.15
        else:
            p_obj.paragraph_format.line_spacing = 1.15

        p_obj.paragraph_format.space_before = Pt(0)
        p_obj.paragraph_format.space_after = Pt(0 if is_table_cell else 2)

        for child in p_tag.children:
            if child.name == 'span':
                span_style = parse_css_style(child.get('style', ''))
                text = child.get_text()
                if not text:
                    img_tag = child.find('img')
                    if img_tag and img_tag.get('src'):
                        try:
                            img_resp = session.get(img_tag['src'], timeout=10)
                            if img_resp.status_code == 200:
                                p_obj.add_run().add_picture(io.BytesIO(img_resp.content), width=Inches(3.0))
                        except Exception:
                            pass
                    continue

                run = p_obj.add_run(text)
                run.font.name = 'Times New Roman'
                
                fs = span_style.get('font-size') or p_style.get('font-size')
                if fs:
                    pt_size = parse_pt(fs)
                    run.font.size = Pt(pt_size) if pt_size > 0 else Pt(13)
                else:
                    run.font.size = Pt(13)

                fw = span_style.get('font-weight', '')
                if fw in ['bold', '700', '800', '900'] or 'bold' in span_style.get('font-style', ''):
                    run.bold = True
                
                fst = span_style.get('font-style', '')
                if 'italic' in fst:
                    run.italic = True
                    
                td = span_style.get('text-decoration', '')
                if 'underline' in td:
                    run.underline = True

                col = parse_color(span_style.get('color'))
                if col and col != RGBColor(0, 0, 0) and col != RGBColor(0x1f, 0x1f, 0x1f):
                    run.font.color.rgb = col

            elif child.name == 'img':
                src = child.get('src')
                if src:
                    try:
                        img_resp = session.get(src, timeout=10)
                        if img_resp.status_code == 200:
                            p_obj.add_run().add_picture(io.BytesIO(img_resp.content), width=Inches(3.0))
                    except Exception:
                        pass

            elif child.name == 'br':
                p_obj.add_run('\n')
            elif isinstance(child, str):
                if child.strip() or child == ' ':
                    run = p_obj.add_run(child)
                    run.font.name = 'Times New Roman'
                    run.font.size = Pt(13)

    def process_table(t_tag):
        rows = t_tag.find_all('tr')
        if not rows:
            return
        
        num_cols = max(len(r.find_all(['td', 'th'])) for r in rows)
        num_rows = len(rows)
        
        tbl = doc.add_table(rows=num_rows, cols=num_cols)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl.autofit = False

        col_widths = []
        first_row_cells = rows[0].find_all(['td', 'th'])
        for td in first_row_cells:
            td_style = parse_css_style(td.get('style', ''))
            w_pt = parse_pt(td_style.get('width', ''))
            col_widths.append(w_pt if w_pt > 0 else 80.0)

        for r_idx, r_tag in enumerate(rows):
            row_cells = r_tag.find_all(['td', 'th'])
            for c_idx, td in enumerate(row_cells):
                if c_idx >= num_cols:
                    break
                cell = tbl.cell(r_idx, c_idx)
                
                if c_idx < len(col_widths):
                    cell.width = Pt(col_widths[c_idx])

                # Kẻ viền đen cho ô (Solid black border)
                border_spec = {"val": "single", "sz": 4, "color": "000000"}
                set_cell_border(cell, top=border_spec, bottom=border_spec, left=border_spec, right=border_spec)
                set_cell_margins(cell, top=50, bottom=50, left=80, right=80)
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

                p_tags = td.find_all('p')
                if p_tags:
                    for p_i, p_tag in enumerate(p_tags):
                        p_obj = cell.paragraphs[0] if p_i == 0 else cell.add_paragraph()
                        process_paragraph(p_tag, p_obj, is_table_cell=True)
                else:
                    p_obj = cell.paragraphs[0]
                    p_obj.paragraph_format.space_before = Pt(0)
                    p_obj.paragraph_format.space_after = Pt(0)
                    p_obj.paragraph_format.line_spacing = 1.15
                    text = td.get_text()
                    if text.strip():
                        run = p_obj.add_run(text.strip())
                        run.font.name = 'Times New Roman'
                        run.font.size = Pt(12)

        after_p = doc.add_paragraph()
        after_p.paragraph_format.space_before = Pt(0)
        after_p.paragraph_format.space_after = Pt(4)

    for el in doc_content.children:
        if el.name == 'p':
            if not el.get_text().strip() and not el.find('img'):
                continue
            p = doc.add_paragraph()
            process_paragraph(el, p)
        elif el.name == 'table':
            process_table(el)
        elif el.name in ['h1', 'h2', 'h3', 'h4', 'h5', 'h6']:
            p = doc.add_paragraph()
            process_paragraph(el, p)
            if p.runs:
                for r in p.runs:
                    r.bold = True
        elif el.name in ['ul', 'ol']:
            for li in el.find_all('li'):
                p = doc.add_paragraph(style='List Bullet' if el.name == 'ul' else 'List Number')
                process_paragraph(li, p)

    doc.save(output_path)
    abs_path = os.path.abspath(output_path)
    file_size_kb = os.path.getsize(abs_path) / 1024
    print(f"  ✓ Tiêu đề: {doc_title}")
    print(f"  ✓ Số đoạn văn: {len(doc.paragraphs)}")
    print(f"  ✓ Số bảng biểu: {len(doc.tables)}")
    print(f"🎉 Đã lưu tệp Word (.docx) chuẩn bản gốc: {abs_path} ({file_size_kb:.1f} KB)")
    return abs_path


async def download_google_doc_to_pdf(url: str, output_path: str = None, scale: int = 2, quality: int = 92, save_images: bool = True):
    print("\n" + "=" * 60)
    print("📄 BẮT ĐẦU TẢI DƯỚI ĐỊNH DẠNG PDF (CANVAS ENGINE)")
    print(f"🔗 URL: {url}")
    print(f"⚙️  Scale DPI: {scale}x | Chất lượng nén: {quality}%")
    print("=" * 60)

    async with async_playwright() as p:
        print("\n[1/4] Khởi động trình duyệt Chromium...")
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(
            viewport={'width': 1600, 'height': 1200},
            device_scale_factor=scale
        )
        page = await context.new_page()

        print("[2/4] Đang mở liên kết Google Docs...")
        await page.goto(url, wait_until="domcontentloaded")
        
        try:
            await page.wait_for_selector(".kix-appview-editor, canvas", timeout=20000)
            print("  ✓ Đã phát hiện vùng tài liệu Google Docs!")
        except Exception as e:
            print(f"  ⚠️ Cảnh báo thời gian chờ: {e}")

        await page.wait_for_timeout(2500)

        raw_title = await page.title()
        doc_title = raw_title.replace(" - Google Docs", "").replace(" - Google Tài liệu", "").strip()
        doc_title = sanitize_filename(doc_title)
        print(f"  📄 Tiêu đề tài liệu: {doc_title}")

        if not output_path:
            output_path = f"{doc_title}.pdf"
        elif not output_path.lower().endswith(".pdf"):
            output_path += ".pdf"

        print("\n[3/4] Đang quét và chụp từng trang (Canvas Render Engine)...")
        result = await page.evaluate('''async () => {
            const editor = document.querySelector('.kix-appview-editor');
            if (!editor) return { error: "Không tìm thấy vùng soạn thảo Google Docs (.kix-appview-editor)" };

            const firstPage = document.querySelector('.kix-page-paginated');
            const pageHeight = firstPage ? (firstPage.offsetHeight || 1123) : 1123;
            const scrollHeight = editor.scrollHeight;

            const capturedPages = new Map();
            const step = pageHeight;
            const totalSteps = Math.ceil(scrollHeight / step) + 2;

            for (let s = 0; s < totalSteps; s++) {
                editor.scrollTop = s * step;
                await new Promise(r => setTimeout(r, 450));

                const pages = document.querySelectorAll('.kix-page-paginated');
                pages.forEach(p => {
                    const canvas = p.querySelector('canvas.kix-canvas-tile-content');
                    if (canvas && canvas.width > 100 && canvas.height > 100) {
                        const top = p.offsetTop;
                        const pageIdx = Math.round(top / (pageHeight + 10));
                        if (!capturedPages.has(pageIdx)) {
                            try {
                                const dataUrl = canvas.toDataURL('image/png');
                                capturedPages.set(pageIdx, {
                                    index: pageIdx,
                                    width: canvas.width,
                                    height: canvas.height,
                                    data: dataUrl
                                });
                            } catch (e) {
                                console.error('Lỗi canvas:', e);
                            }
                        }
                    }
                });
            }

            const sorted = Array.from(capturedPages.values()).sort((a, b) => a.index - b.index);
            return {
                totalPages: sorted.length,
                pages: sorted
            };
        }''')

        if 'error' in result:
            print(f"❌ Lỗi: {result['error']}")
            await browser.close()
            return None

        total_pages = result['totalPages']
        print(f"  ✓ Tổng số trang đã bắt được: {total_pages} trang")

        print(f"\n[4/4] Đang tối ưu và biên dịch {total_pages} trang thành tệp PDF...")
        pdf_doc = pymupdf.open()
        
        img_dir = f"images_{doc_title}"
        if save_images:
            os.makedirs(img_dir, exist_ok=True)

        for i, p_info in enumerate(result['pages']):
            print(f"  -> Xử lý trang {i+1}/{total_pages} (Kích thước: {p_info['width']}x{p_info['height']}px)...")
            header, encoded = p_info['data'].split(",", 1)
            img_bytes = base64.b64decode(encoded)

            img = Image.open(io.BytesIO(img_bytes)).convert("RGB")
            opt_buffer = io.BytesIO()
            img.save(opt_buffer, format="JPEG", quality=quality, optimize=True)
            opt_bytes = opt_buffer.getvalue()

            if save_images:
                img_path = os.path.join(img_dir, f"page_{i+1:03d}.png")
                with open(img_path, "wb") as f:
                    f.write(img_bytes)

            pt_width = img.width * 72 / (96 * scale)
            pt_height = img.height * 72 / (96 * scale)
            
            rect = pymupdf.Rect(0, 0, pt_width, pt_height)
            pdf_page = pdf_doc.new_page(width=pt_width, height=pt_height)
            pdf_page.insert_image(rect, stream=opt_bytes)

        pdf_doc.save(output_path, deflate=True)
        pdf_doc.close()
        await browser.close()

        abs_pdf_path = os.path.abspath(output_path)
        file_size_mb = os.path.getsize(abs_pdf_path) / (1024 * 1024)
        print("\n" + "=" * 60)
        print("🎉 HOÀN THÀNH XUẤT SẮC TỆP PDF!")
        print(f"📁 Tệp PDF đã tạo: {abs_pdf_path} ({file_size_mb:.2f} MB)")
        if save_images:
            print(f"🖼️  Thư mục ảnh các trang: {os.path.abspath(img_dir)}")
        print("=" * 60)
        return abs_pdf_path


def main():
    parser = argparse.ArgumentParser(description="Tải tài liệu Google Docs, Google Slides & Drive PDF View-Only sang PDF, DOCX hoặc PPTX")
    parser.add_argument("url", nargs="?", default="https://docs.google.com/document/d/1KN5IuYT_D3wzsx1tJ0rdOQxrmpSs3uCX/edit", help="Đường link Google Docs, Google Slides hoặc Drive PDF cần tải")
    parser.add_argument("-f", "--format", choices=["pdf", "docx", "pptx", "all"], default="all", help="Định dạng xuất ra: pdf, docx, pptx hoặc all (mặc định: all)")
    parser.add_argument("-o", "--output", help="Tên file xuất ra", default=None)
    parser.add_argument("-s", "--scale", type=int, default=2, help="Độ phân giải DPI cho PDF/PPTX (2 = 2x Retina)")
    parser.add_argument("-q", "--quality", type=int, default=92, help="Chất lượng nén ảnh (1-100, mặc định 92)")
    parser.add_argument("--no-images", action="store_true", help="Không lưu ảnh PNG riêng lẻ")

    args = parser.parse_args()

    # 1. Google Slides / Presentation
    if "docs.google.com/presentation" in args.url:
        from downloader_core import download_single_presentation_pptx, download_single_presentation_pdf
        out_base = args.output or "GoogleSlides_Export"
        if out_base.lower().endswith(".pptx") or out_base.lower().endswith(".pdf"):
            out_base = os.path.splitext(out_base)[0]

        if args.format in ["pptx", "all"]:
            pptx_file = f"{out_base}.pptx"
            download_single_presentation_pptx(args.url, pptx_file, scale=args.scale, quality=args.quality, log_cb=print)
        if args.format in ["pdf", "all"]:
            pdf_file = f"{out_base}.pdf"
            download_single_presentation_pdf(args.url, pdf_file, scale=args.scale, quality=args.quality, log_cb=print)
        return

    # 2. Check if URL is Google Drive PDF
    if "drive.google.com/file" in args.url:
        from downloader_core import download_drive_pdf
        out = args.output or "Drive_Export.pdf"
        asyncio.run(download_drive_pdf(args.url, out, quality=args.quality, log_cb=print))
        return

    # 3. Google Docs
    if args.format in ["docx", "all"]:
        download_google_doc_to_docx(args.url, args.output)

    if args.format in ["pdf", "all"]:
        asyncio.run(download_google_doc_to_pdf(
            url=args.url,
            output_path=args.output,
            scale=args.scale,
            quality=args.quality,
            save_images=not args.no_images
        ))


if __name__ == "__main__":
    main()

