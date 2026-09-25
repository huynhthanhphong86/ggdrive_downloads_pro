"""
Core Module for Google Drive, Google Docs & Google Slides View-Only Downloader
Hỗ trợ:
1. Quét thư mục Google Drive (Folder Scanning)
2. Tải tài liệu Google Docs sang .DOCX và .PDF
3. Tải bài thuyết trình Google Slides / PPTX sang .PPTX và .PDF
4. Tải tệp PDF View-Only từ Google Drive Viewer sang .PDF chuẩn 100% (hỗ trợ mọi số lượng trang)
"""

import os
import sys
import json
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
from pptx import Presentation
from pptx.util import Inches as PptxInches, Pt as PptxPt
from pptx.dml.color import RGBColor as PptxRGB
from pptx.oxml.ns import qn as pptx_qn
from lxml import etree

if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8')


def extract_id_and_type(url: str) -> tuple:
    """Xác định ID và loại liên kết (folder, doc, presentation, pdf, drive_file)"""
    url = url.strip()
    
    # 1. Google Drive Folder
    folder_match = re.search(r"drive\.google\.com/drive/(?:u/\d+/)?folders/([a-zA-Z0-9-_]+)", url)
    if folder_match:
        return folder_match.group(1), "folder"
    
    # 2. Google Slides / Presentation
    presentation_match = re.search(r"docs\.google\.com/presentation/(?:u/\d+/)?d/([a-zA-Z0-9-_]+)", url)
    if presentation_match:
        return presentation_match.group(1), "presentation"

    # 3. Google Docs
    doc_match = re.search(r"docs\.google\.com/document/(?:u/\d+/)?d/([a-zA-Z0-9-_]+)", url)
    if doc_match:
        return doc_match.group(1), "doc"
        
    # 4. Google Drive File (PDF, PPTX hoặc Docx)
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
                            const nameLower = name.toLowerCase();
                            const labelLower = label.toLowerCase();

                            const isFolder = el.querySelector('[data-is-folder="true"]') !== null || 
                                             labelLower.includes('thư mục') || 
                                             labelLower.includes('folder');
                            
                            const isPdf = nameLower.endsWith('.pdf') || labelLower.includes('.pdf');
                            const isPresentation = nameLower.endsWith('.pptx') || 
                                                   nameLower.endsWith('.ppt') || 
                                                   labelLower.includes('.pptx') || 
                                                   labelLower.includes('.ppt') || 
                                                   labelLower.includes('trang trình bày') || 
                                                   labelLower.includes('presentation') || 
                                                   labelLower.includes('slide');

                            const isDriveFile = !isPresentation && !isPdf && !isFolder && (
                                nameLower.endsWith('.jpg') || 
                                nameLower.endsWith('.jpeg') || 
                                nameLower.endsWith('.png') || 
                                nameLower.endsWith('.webp') || 
                                nameLower.endsWith('.gif') || 
                                nameLower.endsWith('.bmp') || 
                                nameLower.endsWith('.zip') || 
                                nameLower.endsWith('.rar') || 
                                nameLower.endsWith('.mp4') || 
                                nameLower.endsWith('.mp3')
                            );
                            
                            let fileUrl = `https://docs.google.com/document/d/${id}/edit`;
                            if (isFolder) {
                                fileUrl = `https://drive.google.com/drive/folders/${id}`;
                            } else if (isPresentation) {
                                fileUrl = `https://docs.google.com/presentation/d/${id}/edit`;
                            } else if (isPdf || isDriveFile) {
                                fileUrl = `https://drive.google.com/file/d/${id}/view`;
                            }
                            
                            foundItems.set(id, {
                                id: id,
                                name: name,
                                isFolder: isFolder,
                                isPdf: isPdf,
                                isPresentation: isPresentation,
                                isDriveFile: isDriveFile,
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
def get_system_ttf_font() -> str:
    """Tìm font TrueType hỗ trợ đầy đủ Unicode / Tiếng Việt trên hệ thống"""
    candidates = [
        "C:/Windows/Fonts/arial.ttf",
        "C:/Windows/Fonts/times.ttf",
        "C:/Windows/Fonts/segoeui.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
        "/System/Library/Fonts/Helvetica.ttc",
        "/System/Library/Fonts/Supplemental/Arial.ttf"
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    return None


def group_words_into_segments(words_list):
    """Gom các từ liền kề trên cùng một dòng thành các cụm từ (phrase segments) chuẩn theo từng cột"""
    if not words_list:
        return []

    # Sắp xếp các từ theo thứ tự: từ trên xuống dưới, từ trái sang phải
    sorted_words = sorted(words_list, key=lambda w: (round(float(w[0][0]) / 3.0), float(w[0][1])))

    segments = []
    current_seg = None

    for box, text in sorted_words:
        y, x, h, w = float(box[0]), float(box[1]), float(box[2]), float(box[3])
        if not text.strip():
            continue

        if current_seg is None:
            current_seg = {
                'y': y,
                'x': x,
                'h': h,
                'w': w,
                'words': [text.strip()],
                'last_x_end': x + w
            }
        else:
            # Kiểm tra xem từ tiếp theo có nằm trên cùng dòng và gần từ trước (cùng ô / cùng cột) không
            same_line = abs(y - current_seg['y']) < 5.0
            gap = x - current_seg['last_x_end']
            
            # Khoảng cách giữa 2 từ trong cùng một cụm < 20 pt. Nếu gap >= 20 pt -> Sang cột khác
            if same_line and -5.0 <= gap < 20.0:
                current_seg['words'].append(text.strip())
                current_seg['w'] = (x + w) - current_seg['x']
                current_seg['h'] = max(current_seg['h'], h)
                current_seg['last_x_end'] = x + w
            else:
                segments.append({
                    'box': [current_seg['y'], current_seg['x'], current_seg['h'], current_seg['w']],
                    'text': ' '.join(current_seg['words'])
                })
                current_seg = {
                    'y': y,
                    'x': x,
                    'h': h,
                    'w': w,
                    'words': [text.strip()],
                    'last_x_end': x + w
                }

    if current_seg:
        segments.append({
            'box': [current_seg['y'], current_seg['x'], current_seg['h'], current_seg['w']],
            'text': ' '.join(current_seg['words'])
        })

    return segments


# --- DRIVE PDF DOWNLOADER (NETWORK INTERCEPTION & ADAPTIVE SCROLLING & TEXT LAYER) ---
async def download_drive_pdf(file_url: str, output_path: str, quality: int = 92, log_cb=print) -> str:
    """Tải tệp PDF View-Only từ Google Drive Viewer với đầy đủ hình ảnh siêu nét và Lớp Text Layer (Searchable/Selectable)"""
    doc_id, _ = extract_id_and_type(file_url)
    preview_url = f"https://drive.google.com/file/d/{doc_id}/preview"
    log_cb(f"  [PDF Viewer] Đang kết nối Google Drive PDF Viewer: {doc_id}")

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(viewport={'width': 1800, 'height': 1200}, device_scale_factor=2)
        page = await context.new_page()

        pages_captured = {}   # page_idx (int) -> image bytes
        text_layers = {}      # page_idx (int) -> list of line dicts
        page_dimensions = {}  # page_idx (int) -> (dpi, pt_height, pt_width)

        async def handle_response(res):
            url = res.url
            # 1. Bắt hình ảnh hiển thị từng trang
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

            # 2. Bắt Lớp Text Layer cấu trúc từ Google Drive Viewer
            elif 'viewer/presspage' in url:
                parsed = urlparse(url)
                qs = parse_qs(parsed.query)
                page_param = qs.get('page', [''])[0]
                if page_param.isdigit():
                    page_idx = int(page_param)
                    try:
                        raw = await res.text()
                        start = raw.find('[')
                        end = raw.rfind(']')
                        if start != -1 and end != -1:
                            data = json.loads(raw[start:end+1])
                            # data structure: [dpi, height, width, lines_blocks]
                            if len(data) >= 4:
                                page_dimensions[page_idx] = (data[0], data[1], data[2])
                                raw_words = []
                                for block in data[3]:
                                    if len(block) >= 2:
                                        def extract_words(node):
                                            if isinstance(node, list):
                                                if len(node) == 2 and isinstance(node[0], list) and isinstance(node[1], str):
                                                    raw_words.append((node[0], node[1]))
                                                else:
                                                    for item in node:
                                                        extract_words(item)
                                        extract_words(block[1])
                                text_layers[page_idx] = group_words_into_segments(raw_words)
                                log_cb(f"    ✓ Đã nạp Text Layer Trang {page_idx + 1} ({len(text_layers[page_idx])} phân đoạn)")
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
        step = 400
        total_steps = max(int(scroll_height / step) + 8, 30)

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

        log_cb(f"  [PDF Compiler] Đang tối ưu hình ảnh và tích hợp Lớp Text Layer cho {total_pages} trang...")
        pdf_doc = pymupdf.open()
        font_path = get_system_ttf_font()
        sorted_indices = sorted(pages_captured.keys())

        text_layer_count = 0
        for idx in sorted_indices:
            img_bytes = pages_captured[idx]
            img = Image.open(io.BytesIO(img_bytes)).convert("RGB")

            opt_buf = io.BytesIO()
            img.save(opt_buf, format="JPEG", quality=quality, optimize=True)
            opt_bytes = opt_buf.getvalue()

            # Nhận diện chính xác hướng trang (Landscape vs Portrait)
            is_landscape = img.width > img.height

            if idx in page_dimensions:
                dpi, d1, d2 = page_dimensions[idx]
                if is_landscape:
                    pt_w = max(float(d1), float(d2))
                    pt_h = min(float(d1), float(d2))
                else:
                    pt_w = min(float(d1), float(d2))
                    pt_h = max(float(d1), float(d2))
            else:
                pt_w = img.width * 72 / 96
                pt_h = img.height * 72 / 96

            rect = pymupdf.Rect(0, 0, pt_w, pt_h)
            pdf_page = pdf_doc.new_page(width=pt_w, height=pt_h)
            # 1. Gắn ảnh trang siêu nét
            pdf_page.insert_image(rect, stream=opt_bytes)

            # 2. Gắn lớp văn bản vô hình (Invisible Text Layer) chuẩn xác từng cụm cột
            if idx in text_layers and text_layers[idx]:
                text_layer_count += 1
                font_id = 'f0'
                if font_path:
                    try:
                        pdf_page.insert_font(fontname=font_id, fontfile=font_path)
                    except Exception:
                        font_id = 'helv'
                else:
                    font_id = 'helv'

                for seg in text_layers[idx]:
                    y, x, h, w = seg['box']
                    pt = pymupdf.Point(float(x), float(y) + float(h) * 0.82)
                    try:
                        pdf_page.insert_text(
                            pt,
                            seg['text'],
                            fontname=font_id,
                            fontsize=max(float(h) * 0.78, 6.0),
                            render_mode=3 # Lớp text ẩn hỗ trợ bôi đen & tìm kiếm
                        )
                    except Exception:
                        pass

        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        pdf_doc.save(output_path, deflate=True)
        pdf_doc.close()

        if text_layer_count > 0:
            log_cb(f"  ✓ Đã xuất PDF hoàn tất: {os.path.basename(output_path)} ({total_pages} trang, kèm Lớp Text Layer bôi đen/tìm kiếm)")
        else:
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
    """Tải PDF từ Google Docs Native Vector Engine hoặc Google Drive PDF Viewer với Text Layer"""
    if "drive.google.com/file" in url:
        return await download_drive_pdf(url, output_path, quality=quality, log_cb=log_cb)

    doc_id, _ = extract_id_and_type(url)

    # 1. Thử xuất Native Vector PDF từ Semantic HTML của Google Docs (Chuẩn 100% Text, Font, Layout)
    mobile_url = f"https://docs.google.com/document/d/{doc_id}/mobilebasic"
    try:
        log_cb(f"  [PDF Vector] Đang trích xuất cấu trúc văn bản: {doc_id}")
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        }
        r = requests.get(mobile_url, headers=headers, timeout=20)
        if r.status_code == 200 and len(r.text) > 500:
            async with async_playwright() as p:
                browser = await p.chromium.launch(headless=True)
                page = await browser.new_page()
                await page.set_content(r.text, wait_until="load")
                await page.add_style_tag(content="""
                    @page { size: A4; margin: 20mm 15mm 20mm 15mm; }
                    body { font-family: 'Times New Roman', 'Arial', sans-serif !important; color: #000 !important; background: #fff !important; }
                    #header, #footer, .mobile-header, .mobile-footer { display: none !important; }
                    .doc-content { width: 100% !important; margin: 0 !important; padding: 0 !important; }
                """)
                pdf_bytes = await page.pdf(
                    format="A4",
                    print_background=True,
                    margin={"top": "20mm", "bottom": "20mm", "left": "15mm", "right": "15mm"}
                )
                await browser.close()
                os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
                with open(output_path, "wb") as f:
                    f.write(pdf_bytes)
                log_cb(f"  ✓ Đã xuất PDF Vector chuẩn 100% Text: {os.path.basename(output_path)}")
                return output_path
    except Exception as e:
        log_cb(f"  ⚠️ Chuyển sang Canvas Engine: {e}")

    # 2. Dự phòng: Canvas Engine
    doc_url = f"https://docs.google.com/document/d/{doc_id}/edit"
    log_cb(f"  [PDF Canvas] Đang kết xuất Canvas tài liệu: {doc_id}")

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


# --- GOOGLE SLIDES / PPTX DOWNLOADER (VECTOR & RETINA SLIDE CAPTURE) ---
async def extract_slides_from_google_presentation(url: str, scale: int = 2, log_cb=print) -> list:
    """Trích xuất hình ảnh toàn bộ các slide chất lượng cao (Retina 4K) từ Google Slides / Presentation (không dính menu/toolbar)"""
    doc_id, _ = extract_id_and_type(url)
    log_cb(f"  [Google Slides] Đang kết nối tới bài thuyết trình: {doc_id}")

    captured_slides = {} # pos (int) -> image bytes

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(
            viewport={'width': 1920, 'height': 1080},
            device_scale_factor=scale
        )
        page = await context.new_page()

        embed_url = f"https://docs.google.com/presentation/d/{doc_id}/embed"
        await page.goto(embed_url, wait_until="domcontentloaded")
        await page.wait_for_timeout(3500)

        # Ẩn hoàn toàn thanh công cụ điều khiển phía dưới để đảm bảo 100% hình ảnh slide sạch và nét
        await page.add_style_tag(content="""
            .punch-viewer-navbar, .punch-viewer-navbar-container {
                display: none !important;
            }
            .punch-viewer-content {
                bottom: 0 !important;
            }
        """)
        await page.wait_for_timeout(400)

        # Lấy tổng số lượng slide trong tài liệu
        total_slides = await page.evaluate('''() => {
            const opt = document.querySelector('.punch-viewer-navbar-page [role="option"]');
            return opt ? parseInt(opt.getAttribute('aria-setsize') || '0') : 0;
        }''')

        if total_slides > 0:
            log_cb(f"  [Google Slides] Phát hiện tổng cộng {total_slides} slide. Bắt đầu trích xuất...")
        else:
            log_cb(f"  [Google Slides] Bắt đầu quét và trích xuất slide...")

        last_logged_pos = 0
        step_count = 0
        max_steps = max(total_slides * 4 + 20, 250) if total_slides > 0 else 250

        while step_count < max_steps:
            step_count += 1

            status = await page.evaluate('''() => {
                const opt = document.querySelector('.punch-viewer-navbar-page [role="option"]');
                const nextBtn = document.querySelector('.punch-viewer-navbar-next');
                const isDisabled = nextBtn ? (nextBtn.getAttribute('aria-disabled') === 'true' || nextBtn.classList.contains('goog-flat-button-disabled')) : false;
                return {
                    pos: opt ? parseInt(opt.getAttribute('aria-posinset') || '1') : 1,
                    total: opt ? parseInt(opt.getAttribute('aria-setsize') || '1') : 1,
                    nextDisabled: isDisabled
                };
            }''')

            pos = status['pos']
            total = status['total'] or total_slides or 1

            if pos != last_logged_pos:
                log_cb(f"    ✓ Đã nạp Slide {pos}/{total}")
                last_logged_pos = pos

            # Chụp vùng hiển thị slide chính
            container = page.locator('.punch-viewer-content, .punch-viewer-svgpage, svg').first
            if await container.count() > 0:
                img_bytes = await container.screenshot(type="png")
            else:
                img_bytes = await page.screenshot(type="png")

            # Ghi đè trạng thái cuối cùng của từng slide (đảm bảo hiển thị đầy đủ mọi hiệu ứng/nội dung hoàn chỉnh)
            captured_slides[pos] = img_bytes

            if status['nextDisabled']:
                log_cb(f"  [Google Slides] Đã duyệt đến slide cuối cùng ({pos}/{total}). Hoàn tất trích xuất!")
                break

            if pos >= total and status['nextDisabled']:
                break

            # Bấm sang bước tiếp theo
            await page.keyboard.press("ArrowRight")
            await page.wait_for_timeout(250)

        await browser.close()

    sorted_indices = sorted(captured_slides.keys())
    if not sorted_indices:
        raise Exception("Không thể trích xuất slide nào từ bài thuyết trình Google Slides.")

    # Trả về danh sách hình ảnh theo thứ tự từ slide 1 đến slide cuối cùng
    return [captured_slides[idx] for idx in sorted_indices]


def extract_slide_texts_from_presentation(url: str, log_cb=print) -> list:
    """Trích xuất text từng slide từ Google Slides qua endpoint htmlpresent (không cần login)"""
    doc_id, _ = extract_id_and_type(url)
    html_url = f"https://docs.google.com/presentation/d/{doc_id}/htmlpresent"
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
    try:
        r = requests.get(html_url, headers=headers, timeout=20)
        if r.status_code != 200:
            log_cb(f"  [Text Extract] Không thể lấy text htmlpresent (HTTP {r.status_code})")
            return []
    except Exception as e:
        log_cb(f"  [Text Extract] Lỗi kết nối: {e}")
        return []

    soup = BeautifulSoup(r.text, 'html.parser')
    full_text = soup.get_text(separator='\n')
    lines = [l.strip() for l in full_text.split('\n') if l.strip()]

    slides = []
    current_slide = []
    total = None
    slide_pattern = re.compile(r'^(\d+)/(\d+)$')
    num_only = re.compile(r'^\d+$')

    for line in lines:
        m = slide_pattern.match(line)
        if m:
            total = int(m.group(2))
            if current_slide:
                slides.append(current_slide)
            current_slide = []
        elif total is not None:
            # Lọc bỏ các dòng chỉ là số (số thứ tự mục)
            if not num_only.match(line) and len(line) > 1:
                current_slide.append(line)

    if current_slide:
        slides.append(current_slide)

    log_cb(f"  [Text Extract] Trích xuất text xong: {len(slides)}/{total or '?'} slide")
    return slides


def _pptx_set_transparent_fill(shape):
    """Xóa fill của shape (trong suốt hoàn toàn) và bỏ border"""
    sp = shape._element
    spPr = sp.find('.//' + pptx_qn('p:spPr'))
    if spPr is None:
        return
    # Xóa fill cũ
    for tag in [pptx_qn('a:noFill'), pptx_qn('a:solidFill'), pptx_qn('a:gradFill'), pptx_qn('a:pattFill'), pptx_qn('a:blipFill')]:
        for el in spPr.findall('.//' + tag):
            el.getparent().remove(el)
    # Thêm noFill
    etree.SubElement(spPr, pptx_qn('a:noFill'))
    # Thêm/cập nhật border: noFill
    ln = spPr.find(pptx_qn('a:ln'))
    if ln is None:
        ln = etree.SubElement(spPr, pptx_qn('a:ln'))
    for child in list(ln):
        ln.remove(child)
    etree.SubElement(ln, pptx_qn('a:noFill'))


def download_single_presentation_pptx(url: str, output_path: str, scale: int = 2, quality: int = 92, log_cb=print) -> str:
    """Tải và đóng gói toàn bộ bài thuyết trình Google Slides / PPTX sang định dạng PowerPoint (.pptx) chuẩn 100%"""
    doc_id, _ = extract_id_and_type(url)
    log_cb(f"  [PPTX Compiler] Đang khởi tạo tệp PowerPoint cho: {doc_id}")

    slide_images = asyncio.run(extract_slides_from_google_presentation(url, scale=scale, log_cb=log_cb))
    total_slides = len(slide_images)

    log_cb(f"  [PPTX Compiler] Đang đóng gói {total_slides} slide sang tệp .pptx...")
    prs = Presentation()

    # Kiểm tra kích thước hình ảnh đầu tiên để nhận diện tỷ lệ (16:9 Widescreen vs 4:3 Standard)
    first_img = Image.open(io.BytesIO(slide_images[0]))
    aspect = first_img.width / first_img.height if first_img.height > 0 else 1.777

    if aspect > 1.5:
        # Tỷ lệ 16:9 Widescreen (13.333 x 7.5 inches)
        prs.slide_width = PptxInches(13.333)
        prs.slide_height = PptxInches(7.5)
    else:
        # Tỷ lệ 4:3 Standard (10.0 x 7.5 inches)
        prs.slide_width = PptxInches(10.0)
        prs.slide_height = PptxInches(7.5)

    blank_layout = prs.slide_layouts[6] # Layout hoàn toàn trống

    for idx, img_b in enumerate(slide_images):
        slide = prs.slides.add_slide(blank_layout)
        slide.shapes.add_picture(
            io.BytesIO(img_b),
            0,
            0,
            width=prs.slide_width,
            height=prs.slide_height
        )

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    prs.save(output_path)
    log_cb(f"  ✓ Đã xuất tệp PowerPoint: {os.path.basename(output_path)} ({total_slides} slide, {os.path.getsize(output_path)/1024:.1f} KB)")
    return output_path


def _parse_css_color(css: str):
    """Chuyển đổi CSS color (rgb/rgba) sang tuple (r, g, b) hoặc None"""
    import re
    m = re.match(r'rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)', css or '')
    return (int(m.group(1)), int(m.group(2)), int(m.group(3))) if m else None


async def _extract_slides_full_layout(url: str, log_cb=print) -> list:
    """
    Dùng Playwright trích xuất layout đầy đủ từng slide:
    - Text: vị trí chính xác, font size, màu, bold/italic, alignment
    - Ảnh: tải qua blob URL trong browser context
    - Màu nền từ CSS computed style
    """
    doc_id, _ = extract_id_and_type(url)
    html_url = f"https://docs.google.com/presentation/d/{doc_id}/htmlpresent"
    log_cb(f"  [Rebuild] Mở htmlpresent để trích xuất layout đầy đủ...")

    JS_EXTRACT = """() => {
        const vw = window.innerWidth, vh = window.innerHeight;

        // Tìm container slide (div 16:9 hoặc 4:3 lớn nhất, không phải body)
        let cr = {left: 0, top: 0, width: vw, height: vh};
        const divs = Array.from(document.querySelectorAll('div, section'));
        for (const d of divs) {
            const r = d.getBoundingClientRect();
            if (r.width < 400 || r.height < 250) continue;
            const rat = r.width / r.height;
            if (Math.abs(rat - 16/9) < 0.25 || Math.abs(rat - 4/3) < 0.25) {
                cr = r; break;
            }
        }
        const SW = cr.width || vw, SH = cr.height || vh;

        // Background color
        let bgEl = document.elementFromPoint(cr.left + 5, cr.top + 5) || document.body;
        let bgColor = '';
        for (let el = bgEl; el && el !== document.documentElement; el = el.parentElement) {
            const c = window.getComputedStyle(el).backgroundColor;
            if (c && c !== 'rgba(0, 0, 0, 0)' && c !== 'transparent') { bgColor = c; break; }
        }

        const texts = [], images = [];
        const seen = new Set();

        // Walk DOM lấy leaf text nodes + images
        function walk(el, depth) {
            if (!el || depth > 20) return;
            const tag = (el.tagName || '').toLowerCase();
            if (['script','style','head','noscript','meta','link','svg'].includes(tag)) return;

            const style = window.getComputedStyle(el);
            if (style.display === 'none' || style.visibility === 'hidden') return;
            if (parseFloat(style.opacity) < 0.05) return;

            // Xử lý ảnh
            if (tag === 'img') {
                const src = el.currentSrc || el.src || '';
                if (src) {
                    const r = el.getBoundingClientRect();
                    if (r.width >= 20 && r.height >= 20 &&
                        r.left >= cr.left - 30 && r.right <= cr.right + 30 &&
                        r.top >= cr.top - 30 && r.bottom <= cr.bottom + 30) {
                        const key = 'I' + Math.round(r.left) + '_' + Math.round(r.top) + '_' + Math.round(r.width);
                        if (!seen.has(key)) {
                            seen.add(key);
                            images.push({
                                src: src,
                                x: (r.left - cr.left) / SW,
                                y: (r.top - cr.top) / SH,
                                w: r.width / SW,
                                h: r.height / SH
                            });
                        }
                    }
                }
                return;
            }

            // Lấy text trực tiếp từ text nodes (không qua children)
            let directText = '';
            for (const node of el.childNodes) {
                if (node.nodeType === 3) directText += node.textContent;
            }
            directText = directText.trim();

            if (directText.length > 0) {
                const r = el.getBoundingClientRect();
                if (r.width > 3 && r.height > 3 &&
                    r.left >= cr.left - 20 && r.right <= cr.right + 20 &&
                    r.top >= cr.top - 5 && r.bottom <= cr.bottom + 5) {
                    const key = 'T' + Math.round(r.left) + '_' + Math.round(r.top) + '_' + directText.substring(0, 25);
                    if (!seen.has(key)) {
                        seen.add(key);
                        texts.push({
                            text: directText,
                            x: (r.left - cr.left) / SW,
                            y: (r.top - cr.top) / SH,
                            w: Math.max(r.width, 20) / SW,
                            h: Math.max(r.height, 10) / SH,
                            fontSize: parseFloat(style.fontSize) || 12,
                            fontWeight: style.fontWeight,
                            fontStyle: style.fontStyle,
                            color: style.color,
                            textAlign: style.textAlign,
                            letterSpacing: style.letterSpacing
                        });
                    }
                }
            }

            for (const child of el.children) walk(child, depth + 1);
        }

        walk(document.body, 0);

        // Sắp xếp text theo y rồi x
        texts.sort((a, b) => Math.round(a.y * 100) - Math.round(b.y * 100) || a.x - b.x);

        return { bgColor, texts, images, SW, SH };
    }"""

    JS_FETCH_IMAGE = """(src) => {
        return fetch(src, {cache: 'force-cache', credentials: 'include'})
            .then(r => r.ok ? r.blob() : Promise.reject('HTTP ' + r.status))
            .then(blob => new Promise((res, rej) => {
                const fr = new FileReader();
                fr.onloadend = () => res(fr.result);
                fr.onerror = () => res(null);
                fr.readAsDataURL(blob);
            }))
            .catch(() => null);
    }"""

    slides_data = []

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=True)
        page = await browser.new_page(viewport={'width': 1280, 'height': 720})
        await page.goto(html_url, wait_until='networkidle', timeout=30000)
        await page.wait_for_timeout(3000)

        # Đếm tổng slide
        total = await page.evaluate("""() => {
            const m = document.body.innerText.match(/(\\d+)\\s*\\/\\s*(\\d+)/);
            return m ? parseInt(m[2]) : 0;
        }""")
        if not total or total > 500:
            total = 200
        log_cb(f"  [Rebuild] Phát hiện {total} slide, bắt đầu trích xuất layout...")

        for i in range(total):
            await page.wait_for_timeout(600)

            slide_info = await page.evaluate(JS_EXTRACT)

            # Tải ảnh qua blob URL trong browser context
            enhanced_images = []
            for img in slide_info.get('images', []):
                src = img.get('src', '')
                if not src:
                    continue
                try:
                    b64 = await page.evaluate(JS_FETCH_IMAGE, src)
                    if b64 and isinstance(b64, str) and ',' in b64:
                        import base64 as _b64
                        img['image_bytes'] = _b64.b64decode(b64.split(',', 1)[1])
                except Exception:
                    pass
                enhanced_images.append(img)

            slide_info['images'] = enhanced_images
            slides_data.append(slide_info)
            n_txt = len(slide_info.get('texts', []))
            n_img = sum(1 for im in enhanced_images if 'image_bytes' in im)
            log_cb(f"    ✓ Slide {i+1}/{total}: {n_txt} text, {n_img} ảnh")

            if i < total - 1:
                await page.keyboard.press('ArrowRight')

        await browser.close()

    return slides_data


def download_single_presentation_pptx_editable(url: str, output_path: str, scale: int = 2, quality: int = 92, log_cb=print) -> str:
    """Tái tạo PPTX chỉnh sửa được: vị trí text chính xác + ảnh minh họa + màu nền + font style đầy đủ"""
    from pptx.enum.text import PP_ALIGN

    doc_id, _ = extract_id_and_type(url)
    log_cb(f"  [PPTX Rebuild] Bắt đầu tái tạo PPTX layout đầy đủ cho: {doc_id}")

    slides_data = asyncio.run(_extract_slides_full_layout(url, log_cb=log_cb))
    total = len(slides_data)

    prs = Presentation()
    # Mặc định 16:9 — sẽ điều chỉnh sau nếu cần
    prs.slide_width = PptxInches(13.333)
    prs.slide_height = PptxInches(7.5)
    SW = 13.333  # inches
    SH = 7.5

    blank_layout = prs.slide_layouts[6]
    align_map = {'center': PP_ALIGN.CENTER, 'right': PP_ALIGN.RIGHT,
                 'justify': PP_ALIGN.JUSTIFY, 'left': PP_ALIGN.LEFT}

    log_cb(f"  [PPTX Rebuild] Đang xây dựng {total} slide...")

    for idx, sd in enumerate(slides_data):
        slide = prs.slides.add_slide(blank_layout)

        # 1. Màu nền
        bg = _parse_css_color(sd.get('bgColor', ''))
        if bg and max(bg) > 0:  # Bỏ qua nếu đen hoàn toàn (chưa detect được)
            slide.background.fill.solid()
            slide.background.fill.fore_color.rgb = PptxRGB(*bg)

        # 2. Ảnh (layer dưới)
        for img in sd.get('images', []):
            if 'image_bytes' not in img:
                continue
            try:
                l = PptxInches(max(0.0, img['x'] * SW))
                t = PptxInches(max(0.0, img['y'] * SH))
                w = PptxInches(max(0.1, img['w'] * SW))
                h = PptxInches(max(0.1, img['h'] * SH))
                slide.shapes.add_picture(io.BytesIO(img['image_bytes']), l, t, w, h)
            except Exception as e:
                log_cb(f"    ⚠ Ảnh slide {idx+1}: {e}")

        # 3. Text boxes (layer trên, có thể chỉnh sửa)
        for txt in sd.get('texts', []):
            try:
                l = PptxInches(max(0.0, txt['x'] * SW))
                t = PptxInches(max(0.0, txt['y'] * SH))
                w = PptxInches(max(0.15, txt['w'] * SW))
                h = PptxInches(max(0.1, txt['h'] * SH))

                tb = slide.shapes.add_textbox(l, t, w, h)
                _pptx_set_transparent_fill(tb)

                tf = tb.text_frame
                tf.word_wrap = False
                tf.auto_size = None

                p = tf.paragraphs[0]
                p.alignment = align_map.get(txt.get('textAlign', 'left'), PP_ALIGN.LEFT)

                run = p.add_run()
                run.text = txt['text']

                # Font size (giới hạn hợp lý)
                fs = max(6, min(120, txt.get('fontSize', 12)))
                run.font.size = PptxPt(fs)

                # Bold / Italic
                fw = str(txt.get('fontWeight', '400')).lower().strip()
                run.font.bold = (fw in ('700', '800', '900', 'bold', 'bolder') or
                                 (fw.isdigit() and int(fw) >= 700))
                run.font.italic = txt.get('fontStyle', '') == 'italic'

                # Màu chữ từ CSS
                col = _parse_css_color(txt.get('color', ''))
                if col:
                    run.font.color.rgb = PptxRGB(*col)

            except Exception as e:
                log_cb(f"    ⚠ Text slide {idx+1}: {e}")

        # 4. Speaker Notes: toàn bộ text đầy đủ để tham khảo
        try:
            all_text = '\n'.join(t['text'] for t in sd.get('texts', []))
            if all_text:
                notes_tf = slide.notes_slide.notes_text_frame
                notes_tf.clear()
                notes_tf.text = f"[Slide {idx + 1}]\n{all_text}"
                for para in notes_tf.paragraphs:
                    for run in para.runs:
                        run.font.size = PptxPt(11)
        except Exception:
            pass

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    prs.save(output_path)
    log_cb(f"  ✓ PPTX Rebuild: {os.path.basename(output_path)} ({total} slide, {os.path.getsize(output_path)/1024:.1f} KB)")
    log_cb(f"  💡 Click vào bất kỳ text nào để chỉnh sửa. Ảnh và màu nền đã được tái tạo từ slide gốc.")
    return output_path


    """Dat hinh anh lam NEN slide (khong phai shape). Text box tren do la shape duy nhat, de click/edit."""
    try:
        img_part, rId = slide.part.get_or_add_image_part(io.BytesIO(image_bytes))
        sld = slide._element
        cSld = sld.find(pptx_qn('p:cSld'))
        if cSld is None:
            return False
        for old in list(cSld.findall(pptx_qn('p:bg'))):
            cSld.remove(old)
        bg = etree.Element(pptx_qn('p:bg'))
        bgPr = etree.SubElement(bg, pptx_qn('p:bgPr'))
        blipFill = etree.SubElement(bgPr, pptx_qn('a:blipFill'))
        blip = etree.SubElement(blipFill, pptx_qn('a:blip'))
        blip.set(pptx_qn('r:embed'), rId)
        stretch = etree.SubElement(blipFill, pptx_qn('a:stretch'))
        etree.SubElement(stretch, pptx_qn('a:fillRect'))
        etree.SubElement(bgPr, pptx_qn('a:effectLst'))
        spTree = cSld.find(pptx_qn('p:spTree'))
        idx_insert = list(cSld).index(spTree) if spTree is not None else len(list(cSld))
        cSld.insert(idx_insert, bg)
        return True
    except Exception:
        return False


def download_single_presentation_pptx_editable(url: str, output_path: str, scale: int = 2, quality: int = 92, log_cb=print) -> str:
    """Tao PPTX chinh sua duoc thuc su: KHONG co anh slide, dung mau nen trich xuat + text box ro rang"""
    doc_id, _ = extract_id_and_type(url)
    log_cb(f"  [PPTX Editable] Bat dau trich xuat text chinh sua duoc cho: {doc_id}")

    # Buoc 1: Lay anh slide o do phan giai thap (chi de lay mau nen, khong luu vao PPTX)
    log_cb(f"  [PPTX Editable] Dang phan tich slide de trich xuat mau nen va text...")
    slide_images = asyncio.run(extract_slides_from_google_presentation(url, scale=1, log_cb=log_cb))
    total_slides = len(slide_images)

    # Buoc 2: Lay text tung slide tu htmlpresent
    slide_texts = extract_slide_texts_from_presentation(url, log_cb=log_cb)

    log_cb(f"  [PPTX Editable] Dang tao PPTX voi {total_slides} slide co text chinh sua duoc...")
    prs = Presentation()

    first_img = Image.open(io.BytesIO(slide_images[0]))
    aspect = first_img.width / first_img.height if first_img.height > 0 else 1.777
    if aspect > 1.5:
        prs.slide_width = PptxInches(13.333)
        prs.slide_height = PptxInches(7.5)
    else:
        prs.slide_width = PptxInches(10.0)
        prs.slide_height = PptxInches(7.5)

    blank_layout = prs.slide_layouts[6]

    for idx, img_b in enumerate(slide_images):
        slide = prs.slides.add_slide(blank_layout)
        texts = slide_texts[idx] if idx < len(slide_texts) else []
        meaningful = [t.strip() for t in texts if len(t.strip()) > 1]

        # Trich xuat mau nen tu goc slide (khong luu anh, chi lay mau)
        try:
            img_rgb = Image.open(io.BytesIO(img_b)).convert('RGB')
            w, h = img_rgb.size
            corners = [
                img_rgb.getpixel((3, 3)),
                img_rgb.getpixel((w - 3, 3)),
                img_rgb.getpixel((3, h - 3)),
                img_rgb.getpixel((w - 3, h - 3)),
                img_rgb.getpixel((w // 2, 3)),
            ]
            bg_r = sum(c[0] for c in corners) // len(corners)
            bg_g = sum(c[1] for c in corners) // len(corners)
            bg_b = sum(c[2] for c in corners) // len(corners)
            luminance = bg_r * 0.299 + bg_g * 0.587 + bg_b * 0.114
            is_dark = luminance < 140
        except Exception:
            bg_r, bg_g, bg_b = 12, 34, 51  # Xanh dam mac dinh
            is_dark = True

        # Dat mau nen solid (KHONG dung anh - chi mau thuan)
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = PptxRGB(bg_r, bg_g, bg_b)

        # Chon mau chu phu hop voi nen
        if is_dark:
            title_color = PptxRGB(0xFF, 0xFF, 0xFF)   # Trang
            body_color  = PptxRGB(0xCC, 0xDD, 0xEE)   # Xanh nhat
        else:
            title_color = PptxRGB(0x0C, 0x22, 0x33)   # Xanh dam
            body_color  = PptxRGB(0x22, 0x33, 0x44)   # Xam dam

        if not meaningful:
            continue

        margin = PptxInches(0.5)
        sw, sh = prs.slide_width, prs.slide_height

        # Phan chia: dong dau = tieu de chinh, dong tiep = tieu de phu, con lai = noi dung
        title_main  = meaningful[0] if len(meaningful) > 0 else ""
        title_sub   = meaningful[1] if len(meaningful) > 1 else ""
        body_lines  = meaningful[2:]

        # === Text box 1: Tieu de chinh ===
        tb_title = slide.shapes.add_textbox(margin, margin, sw - 2 * margin, PptxInches(1.3))
        _pptx_set_transparent_fill(tb_title)
        tf = tb_title.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = title_main
        run.font.size = PptxPt(36)
        run.font.bold = True
        run.font.color.rgb = title_color

        # === Text box 2: Tieu de phu (neu co) ===
        if title_sub:
            tb_sub = slide.shapes.add_textbox(margin, PptxInches(1.6), sw - 2 * margin, PptxInches(0.7))
            _pptx_set_transparent_fill(tb_sub)
            tf = tb_sub.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            run = p.add_run()
            run.text = title_sub
            run.font.size = PptxPt(22)
            run.font.bold = False
            run.font.color.rgb = body_color

        # === Text box 3: Noi dung ===
        if body_lines:
            body_top = PptxInches(2.5)
            body_h = sh - body_top - margin
            bfp = max(12, min(20, 600 // max(sum(len(l) for l in body_lines), 20)))
            tb_body = slide.shapes.add_textbox(margin, body_top, sw - 2 * margin, body_h)
            _pptx_set_transparent_fill(tb_body)
            tf = tb_body.text_frame
            tf.word_wrap = True
            for i, line in enumerate(body_lines):
                p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                run = p.add_run()
                run.text = line
                run.font.size = PptxPt(bfp)
                run.font.color.rgb = body_color

        # === Speaker Notes: toan bo text day du, de copy ===
        try:
            notes_tf = slide.notes_slide.notes_text_frame
            notes_tf.clear()
            notes_tf.text = f"[Slide {idx + 1}]\n" + '\n'.join(meaningful)
            for para in notes_tf.paragraphs:
                for run in para.runs:
                    run.font.size = PptxPt(12)
        except Exception:
            pass

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    prs.save(output_path)
    log_cb(f"  Da xuat PPTX Editable: {os.path.basename(output_path)} ({total_slides} slide, {os.path.getsize(output_path)/1024:.1f} KB)")
    log_cb(f"  TIPS: Click vao tieu de/noi dung tren slide de chinh sua. Tab 'Notes' co toan bo text day du.")
    return output_path


def download_single_presentation_pdf(url: str, output_path: str, scale: int = 2, quality: int = 92, log_cb=print) -> str:
    """Tải và đóng gói bài thuyết trình Google Slides / PPTX sang tệp PDF (.pdf) siêu nét"""
    doc_id, _ = extract_id_and_type(url)
    log_cb(f"  [PDF Slides] Đang khởi tạo tệp PDF trình chiếu cho: {doc_id}")

    slide_images = asyncio.run(extract_slides_from_google_presentation(url, scale=scale, log_cb=log_cb))
    total_slides = len(slide_images)

    log_cb(f"  [PDF Slides] Đang tối ưu hình ảnh cho {total_slides} slide...")
    pdf_doc = pymupdf.open()

    first_img = Image.open(io.BytesIO(slide_images[0]))
    aspect = first_img.width / first_img.height if first_img.height > 0 else 1.777

    if aspect > 1.5:
        pt_w, pt_h = 960.0, 540.0  # 16:9
    else:
        pt_w, pt_h = 720.0, 540.0  # 4:3

    rect = pymupdf.Rect(0, 0, pt_w, pt_h)

    for idx, img_b in enumerate(slide_images):
        img = Image.open(io.BytesIO(img_b)).convert("RGB")
        opt_buf = io.BytesIO()
        img.save(opt_buf, format="JPEG", quality=quality, optimize=True)
        opt_bytes = opt_buf.getvalue()

        pdf_page = pdf_doc.new_page(width=pt_w, height=pt_h)
        pdf_page.insert_image(rect, stream=opt_bytes)

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    pdf_doc.save(output_path, deflate=True)
    pdf_doc.close()
    log_cb(f"  ✓ Đã xuất tệp PDF bài thuyết trình: {os.path.basename(output_path)} ({total_slides} slide, {os.path.getsize(output_path)/1024:.1f} KB)")
    return output_path

