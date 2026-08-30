"""
Core Module for Google Drive & Google Docs View-Only Downloader
Hỗ trợ:
1. Quét thư mục Google Drive (Folder Scanning)
2. Tải tài liệu Google Docs sang .DOCX và .PDF
3. Tải tệp PDF View-Only từ Google Drive Viewer sang .PDF chuẩn 100% (hỗ trợ mọi số lượng trang)
"""

import os
import sys
import re
import io
import time
import base64
import asyncio
from urllib.parse import urlparse, parse_qs
import requests
from bs4 import BeautifulSoup
import docx
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from playwright.async_api import async_playwright
import pymupdf
from PIL import Image

if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8')


def extract_id_and_type(url: str) -> tuple:
    """Xác định ID và loại liên kết (folder, doc, pdf, file)"""
    url = url.strip()
    
    # 1. Google Drive Folder
    folder_match = re.search(r"drive\.google\.com/drive/(?:u/\d+/)?folders/([a-zA-Z0-9-_]+)", url)
    if folder_match:
        return folder_match.group(1), "folder"
    
    # 2. Google Docs
    doc_match = re.search(r"docs\.google\.com/document/(?:u/\d+/)?d/([a-zA-Z0-9-_]+)", url)
    if doc_match:
        return doc_match.group(1), "doc"
        
    # 3. Google Drive File (PDF hoặc Docx)
    file_match = re.search(r"drive\.google\.com/file/(?:u/\d+/)?d/([a-zA-Z0-9-_]+)", url)
    if file_match:
        file_id = file_match.group(1)
        return file_id, "drive_file"
        
    # Raw ID
    if re.match(r"^[a-zA-Z0-9-_]{15,}$", url):
        return url, "unknown"
        
    return url, "unknown"


def sanitize_filename(name: str) -> str:
    name = re.sub(r'[\\/*?:"<>|]', "", name)
    name = name.strip()
    return name if name else "Document"


async def scan_google_drive_folder(folder_url: str, log_cb=print) -> list:
    """Quét toàn bộ danh sách tệp trong thư mục Google Drive"""
    log_cb(f"🔍 Bắt đầu quét thư mục Google Drive: {folder_url}")
    
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        page = await browser.new_page(viewport={'width': 1600, 'height': 1200})
        
        await page.goto(folder_url, wait_until="domcontentloaded")
        await page.wait_for_timeout(3500)
        
        items_map = await page.evaluate('''async () => {
            const foundItems = new Map();
            
            function scanCurrent() {
                const elements = document.querySelectorAll('[data-id]');
                elements.forEach(el => {
                    const id = el.getAttribute('data-id');
                    if (id && id.length >= 10 && !id.startsWith('drive-')) {
                        let name = '';
                        const label = el.getAttribute('aria-label') || '';
                        if (label) {
                            name = label;
                        } else {
                            const textLines = (el.innerText || '').split('\\n').map(s => s.trim()).filter(Boolean);
                            if (textLines.length > 0) {
                                name = textLines[0];
                            }
                        }
                        
                        if (name && !foundItems.has(id)) {
                            const isFolder = el.querySelector('[data-is-folder="true"]') !== null || 
                                             label.toLowerCase().includes('thư mục') || 
                                             label.toLowerCase().includes('folder');
                            
                            const isPdf = name.toLowerCase().endsWith('.pdf') || label.toLowerCase().includes('.pdf');
                            
                            let fileUrl = `https://docs.google.com/document/d/${id}/edit`;
                            if (isFolder) {
                                fileUrl = `https://drive.google.com/drive/folders/${id}`;
                            } else if (isPdf) {
                                fileUrl = `https://drive.google.com/file/d/${id}/view`;
                            }
                            
                            foundItems.set(id, {
                                id: id,
                                name: name,
                                isFolder: isFolder,
                                isPdf: isPdf,
                                url: fileUrl
                            });
                        }
                    }
                });
            }
            
            const scrollContainer = document.querySelector('[role="main"]') || document.body;
            let lastCount = 0;
            let sameCountTries = 0;
            
            for (let i = 0; i < 25; i++) {
                scanCurrent();
                window.scrollBy(0, 1200);
                if (scrollContainer) scrollContainer.scrollTop += 1200;
                await new Promise(r => setTimeout(r, 450));
                
                if (foundItems.size === lastCount) {
                    sameCountTries++;
                    if (sameCountTries >= 3) break;
                } else {
                    sameCountTries = 0;
                    lastCount = foundItems.size;
                }
            }
            
            return Array.from(foundItems.values());
        }''')
        
        await browser.close()
        log_cb(f"✅ Quét thành công: Phát hiện {len(items_map)} tệp trong thư mục.")
        return items_map


# --- DRIVE PDF DOWNLOADER (NETWORK INTERCEPTION & ADAPTIVE SCROLLING) ---
async def download_drive_pdf(file_url: str, output_path: str, quality: int = 92, log_cb=print) -> str:
    """Tải tệp PDF View-Only từ Google Drive Viewer đầy đủ 100% các trang"""
    doc_id, _ = extract_id_and_type(file_url)
    preview_url = f"https://drive.google.com/file/d/{doc_id}/preview"
    log_cb(f"  [PDF Viewer] Đang kết nối Google Drive PDF Viewer: {doc_id}")

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(viewport={'width': 1800, 'height': 1200}, device_scale_factor=2)
        page = await context.new_page()

        pages_captured = {} # page_idx (int) -> bytes

        async def handle_response(res):
            url = res.url
            if 'viewer/img' in url and 'auditContext=forDisplay' in url:
                parsed = urlparse(url)
                qs = parse_qs(parsed.query)
                page_param = qs.get('page', [''])[0]
                if page_param.isdigit():
                    page_idx = int(page_param)
                    try:
                        body = await res.body()
                        if len(body) > 1000 and page_idx not in pages_captured:
                            pages_captured[page_idx] = body
                            log_cb(f"    ✓ Đã nạp Trang {page_idx + 1} ({len(body)/1024:.1f} KB)")
                    except Exception:
                        pass

        page.on("response", handle_response)
        
        await page.goto(preview_url, wait_until="domcontentloaded")
        await page.wait_for_timeout(3500)

        # Focus container
        await page.mouse.move(900, 600)
        await page.mouse.click(900, 600)

        # Lấy thông tin chiều cao container
        container_info = await page.evaluate('''() => {
            const container = document.querySelector('.ndfHFb-c4YZDc-cYSp0e-s2gQvd') || 
                              document.querySelector('.drive-viewer-paginated-scrollable') || 
                              document.querySelector('[role="main"]') || 
                              Array.from(document.querySelectorAll('div')).find(d => d.scrollHeight > d.clientHeight + 400) || 
                              document.documentElement;
            return {
                scrollHeight: container.scrollHeight,
                clientHeight: container.clientHeight
            };
        }''')

        scroll_height = container_info.get('scrollHeight', 10000)
        step = 500
        total_steps = max(int(scroll_height / step) + 6, 30)

        log_cb(f"  [PDF Scan] Chiều cao tài liệu: {scroll_height}px (~{total_steps} bước quét)...")

        # Quét lần 1: Từ trên xuống dưới
        for s in range(total_steps):
            await page.evaluate(f'''(stepVal) => {{
                const container = document.querySelector('.ndfHFb-c4YZDc-cYSp0e-s2gQvd') || 
                                  document.querySelector('.drive-viewer-paginated-scrollable') || 
                                  document.querySelector('[role="main"]') || 
                                  Array.from(document.querySelectorAll('div')).find(d => d.scrollHeight > d.clientHeight + 400) || 
                                  document.documentElement;
                container.scrollTop = {s} * stepVal;
                window.scrollBy(0, stepVal);
            }}''', step)

            await page.mouse.wheel(0, step)
            await page.keyboard.press("PageDown")
            await page.wait_for_timeout(280)

        await page.wait_for_timeout(1000)

        # Kiểm tra trang bị thiếu và quét bổ sung
        if pages_captured:
            max_page = max(pages_captured.keys())
            missing_pages = [p for p in range(max_page + 1) if p not in pages_captured]

            if missing_pages:
                log_cb(f"  [PDF Retry] Đang quét bổ sung {len(missing_pages)} trang còn thiếu: {[p+1 for p in missing_pages]}...")
                for miss_p in missing_pages:
                    approx_scroll = int((miss_p / (max_page + 1)) * scroll_height)
                    await page.evaluate(f'''(scrollTarget) => {{
                        const container = document.querySelector('.ndfHFb-c4YZDc-cYSp0e-s2gQvd') || 
                                          document.querySelector('.drive-viewer-paginated-scrollable') || 
                                          document.querySelector('[role="main"]') || 
                                          Array.from(document.querySelectorAll('div')).find(d => d.scrollHeight > d.clientHeight + 400) || 
                                          document.documentElement;
                        container.scrollTop = scrollTarget;
                    }}''', approx_scroll)
                    await page.mouse.wheel(0, 100)
                    await page.wait_for_timeout(600)

        # Scroll to end
        await page.evaluate('''() => {
            const container = document.querySelector('.ndfHFb-c4YZDc-cYSp0e-s2gQvd') || 
                              document.querySelector('.drive-viewer-paginated-scrollable') || 
                              document.querySelector('[role="main"]') || 
                              Array.from(document.querySelectorAll('div')).find(d => d.scrollHeight > d.clientHeight + 400) || 
                              document.documentElement;
            container.scrollTop = container.scrollHeight;
        }''')
        await page.wait_for_timeout(1500)
        await browser.close()

        total_pages = len(pages_captured)
        if total_pages == 0:
            raise Exception("Không trích xuất được trang PDF nào từ Google Drive.")

        pdf_doc = pymupdf.open()
        sorted_indices = sorted(pages_captured.keys())

        for idx in sorted_indices:
            img_bytes = pages_captured[idx]
            img = Image.open(io.BytesIO(img_bytes)).convert("RGB")

            opt_buf = io.BytesIO()
            img.save(opt_buf, format="JPEG", quality=quality, optimize=True)
            opt_bytes = opt_buf.getvalue()

            pt_w = img.width * 72 / 96
            pt_h = img.height * 72 / 96

            rect = pymupdf.Rect(0, 0, pt_w, pt_h)
            pdf_page = pdf_doc.new_page(width=pt_w, height=pt_h)
            pdf_page.insert_image(rect, stream=opt_bytes)

        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        pdf_doc.save(output_path, deflate=True)
        pdf_doc.close()

        log_cb(f"  ✓ Đã xuất tệp PDF hoàn chỉnh: {os.path.basename(output_path)} ({total_pages} trang)")
        return output_path


# --- DOCX CONVERSION UTILS ---
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


def download_single_docx(url: str, output_path: str, log_cb=print) -> str:
    """Tải và xuất file Word .docx chuẩn 100% bản gốc từ Google Docs"""
    doc_id, _ = extract_id_and_type(url)
    mobile_url = f"https://docs.google.com/document/d/{doc_id}/mobilebasic"
    
    log_cb(f"  [DOCX] Đang lấy nội dung cấu trúc: {doc_id}")
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    }
    r = requests.get(mobile_url, headers=headers, timeout=20)
    if r.status_code != 200:
        raise Exception(f"Không thể truy cập tài liệu qua mobile endpoint (HTTP {r.status_code})")

    soup = BeautifulSoup(r.text, 'html.parser')
    doc_content = soup.find('div', class_='doc-content') or soup.find('body')
    doc = docx.Document()

    # Cấu hình A4
    for section in doc.sections:
        section.page_width = Cm(21.0)
        section.page_height = Cm(29.7)
        section.top_margin = Cm(2.0)
        section.bottom_margin = Cm(2.0)
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(2.0)

    # Style Normal
    style_normal = doc.styles['Normal']
    style_normal.font.name = 'Times New Roman'
    style_normal.font.size = Pt(13)
    style_normal.font.color.rgb = RGBColor(0, 0, 0)
    style_normal.paragraph_format.line_spacing = 1.15
    style_normal.paragraph_format.space_before = Pt(0)
    style_normal.paragraph_format.space_after = Pt(2)

    session = requests.Session()
    session.headers.update({'User-Agent': 'Mozilla/5.0'})

    def process_paragraph(p_tag, p_obj, is_table_cell=False):
        p_style = parse_css_style(p_tag.get('style', ''))
        
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

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    doc.save(output_path)
    log_cb(f"  ✓ Đã xuất tệp Word: {os.path.basename(output_path)}")
    return output_path


async def download_single_pdf(url: str, output_path: str, scale: int = 2, quality: int = 92, log_cb=print) -> str:
    """Tải PDF từ Google Docs Canvas Engine hoặc Google Drive PDF Viewer"""
    if "drive.google.com/file" in url:
        return await download_drive_pdf(url, output_path, quality=quality, log_cb=log_cb)

    doc_id, _ = extract_id_and_type(url)
    doc_url = f"https://docs.google.com/document/d/{doc_id}/edit"
    log_cb(f"  [PDF] Đang kết xuất Canvas tài liệu: {doc_id}")

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(
            viewport={'width': 1600, 'height': 1200},
            device_scale_factor=scale
        )
        page = await context.new_page()

        await page.goto(doc_url, wait_until="domcontentloaded")
        try:
            await page.wait_for_selector(".kix-appview-editor, canvas", timeout=18000)
        except Exception:
            pass
        await page.wait_for_timeout(2000)

        result = await page.evaluate('''async () => {
            const editor = document.querySelector('.kix-appview-editor');
            if (!editor) return { error: "Không tìm thấy editor" };

            const firstPage = document.querySelector('.kix-page-paginated');
            const pageHeight = firstPage ? (firstPage.offsetHeight || 1123) : 1123;
            const scrollHeight = editor.scrollHeight;

            const capturedPages = new Map();
            const step = pageHeight;
            const totalSteps = Math.ceil(scrollHeight / step) + 2;

            for (let s = 0; s < totalSteps; s++) {
                editor.scrollTop = s * step;
                await new Promise(r => setTimeout(r, 400));

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
                            } catch (e) {}
                        }
                    }
                });
            }

            const sorted = Array.from(capturedPages.values()).sort((a, b) => a.index - b.index);
            return { totalPages: sorted.length, pages: sorted };
        }''')

        if 'error' in result or not result.get('pages'):
            await browser.close()
            raise Exception("Không thể chụp canvas của tài liệu")

        pdf_doc = pymupdf.open()
        for i, p_info in enumerate(result['pages']):
            header, encoded = p_info['data'].split(",", 1)
            img_bytes = base64.b64decode(encoded)

            img = Image.open(io.BytesIO(img_bytes)).convert("RGB")
            opt_buffer = io.BytesIO()
            img.save(opt_buffer, format="JPEG", quality=quality, optimize=True)
            opt_bytes = opt_buffer.getvalue()

            pt_width = img.width * 72 / (96 * scale)
            pt_height = img.height * 72 / (96 * scale)
            
            rect = pymupdf.Rect(0, 0, pt_width, pt_height)
            pdf_page = pdf_doc.new_page(width=pt_width, height=pt_height)
            pdf_page.insert_image(rect, stream=opt_bytes)

        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        pdf_doc.save(output_path, deflate=True)
        pdf_doc.close()
        await browser.close()
        
        log_cb(f"  ✓ Đã xuất tệp PDF: {os.path.basename(output_path)} ({result['totalPages']} trang)")
        return output_path
