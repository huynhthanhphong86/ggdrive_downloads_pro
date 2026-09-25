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
from concurrent.futures import ThreadPoolExecutor
import requests
from bs4 import BeautifulSoup, NavigableString, Tag
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
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml import parse_xml
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


def _download_single_presentation_pptx_images(url: str, output_path: str, scale: int = 2, quality: int = 92, log_cb=print) -> str:
    """Chế độ dự phòng: Chụp ảnh toàn bộ các slide bằng Playwright và đóng gói sang PPTX"""
    doc_id, _ = extract_id_and_type(url)
    log_cb(f"  [PPTX Fallback] Đang chụp ảnh slide qua Playwright cho: {doc_id}")

    slide_images = asyncio.run(extract_slides_from_google_presentation(url, scale=scale, log_cb=log_cb))
    total_slides = len(slide_images)

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


def download_single_presentation_pptx(url: str, output_path: str, scale: int = 2, quality: int = 92, log_cb=print) -> str:
    """Tải và đóng gói toàn bộ bài thuyết trình Google Slides / PPTX sang PowerPoint (.pptx) chuẩn đồ họa + text chỉnh sửa được"""
    return download_single_presentation_pptx_editable(url, output_path, scale=scale, quality=quality, log_cb=log_cb)


def _parse_px(val_str, default=0.0):
    if not val_str:
        return default
    m = re.search(r'([\d.]+)\s*px', str(val_str))
    if m:
        return float(m.group(1))
    m = re.search(r'([\d.]+)\s*pt', str(val_str))
    if m:
        return float(m.group(1)) * 1.3333
    return default


def _parse_pt(val_str, default=14.0):
    if not val_str:
        return default
    m = re.search(r'([\d.]+)\s*pt', str(val_str))
    if m:
        return float(m.group(1))
    m = re.search(r'([\d.]+)\s*px', str(val_str))
    if m:
        return float(m.group(1)) * 0.75
    return default


def _parse_css_dict(style_str):
    styles = {}
    if not style_str:
        return styles
    for item in style_str.split(';'):
        if ':' in item:
            k, v = item.split(':', 1)
            styles[k.strip().lower()] = v.strip()
    return styles


def _parse_css_color_rgb(c_str):
    if not c_str:
        return None
    c_str = c_str.strip()
    if c_str.startswith('#'):
        hex_c = c_str.lstrip('#')
        if len(hex_c) == 3:
            hex_c = ''.join([c*2 for c in hex_c])
        if len(hex_c) == 6:
            try:
                return PptxRGB(int(hex_c[0:2], 16), int(hex_c[2:4], 16), int(hex_c[4:6], 16))
            except Exception:
                return None
    m = re.match(r'rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)', c_str)
    if m:
        try:
            return PptxRGB(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except Exception:
            return None
    return None


def _hex_to_pptx_rgb(hex_str: str) -> PptxRGB:
    if not hex_str or not hex_str.startswith('#'):
        return PptxRGB(0, 0, 0)
    h = hex_str.lstrip('#')
    if len(h) == 3:
        h = ''.join([c*2 for c in h])
    if len(h) == 6:
        try:
            return PptxRGB(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
        except Exception:
            return PptxRGB(0, 0, 0)
    return PptxRGB(0, 0, 0)


def _apply_shape_opacity(shape, opacity: float):
    """Gán độ trong suốt (opacity/alpha) cho vector shape trong python-pptx"""
    if opacity >= 0.99:
        return
    try:
        alpha_val = int(max(0.0, min(1.0, opacity)) * 100000)
        sp_elem = shape._element
        spPr = sp_elem.spPr
        solidFill = spPr.find('{http://schemas.openxmlformats.org/drawingml/2006/main}solidFill')
        if solidFill is not None:
            srgb = solidFill.find('{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr')
            if srgb is not None:
                alpha = parse_xml(f'<a:alpha xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" val="{alpha_val}"/>')
                srgb.append(alpha)
    except Exception:
        pass


def download_single_presentation_pptx_editable(url: str, output_path: str, scale: int = 2, quality: int = 92, log_cb=print) -> str:
    """
    Tái tạo tệp PowerPoint PPTX hoàn chỉnh với các đối tượng tách biệt tối đa:
    - Nền slide (Background): Nền riêng biệt (màu chuẩn hoặc background fill)
    - Khung Shapes: Vector shapes riêng biệt (Hình Ôvan, Hình chữ nhật bo góc, Card, Badge)
      với đầy đủ màu fill, độ mờ (alpha/transparency), viền
    - Hình ảnh (Images): Các ảnh minh họa/đồ họa dạng PNG trong suốt / JPEG độc lập
      có thể phóng to, thu nhỏ, di chuyển, thay thế
    - Văn bản (Text layer): Từng textbox riêng biệt với font, cỡ chữ, màu sắc, bold, italic
    - Ghi chú: Speaker Notes đầy đủ
    - Chế độ dự phòng (Fallback): Nếu tài liệu không đọc được /edit cấu trúc chi tiết,
      tự động dùng cơ chế showText=0 + Text layer hoặc Playwright
    """
    doc_id, _ = extract_id_and_type(url)
    log_cb(f"  [PPTX Tách đối tượng] Bắt đầu xử lý bài thuyết trình: {doc_id}")

    edit_url = f"https://docs.google.com/presentation/d/{doc_id}/edit"
    html_url = f"https://docs.google.com/presentation/d/{doc_id}/htmlpresent"
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}

    # -------------------------------------------------------------
    # PHƯƠNG PHÁP 1: TÁCH BIỆT HOÀN TOÀN CÁC ĐỐI TƯỢNG (SEPARATED OBJECTS ENGINE)
    # Nền riêng + Khung shape vector riêng + Ảnh minh họa riêng + Text layer chuẩn
    # -------------------------------------------------------------
    try:
        log_cb(f"  [PPTX Tách đối tượng] Bước 1/3: Đang phân tích mô hình tài liệu (/edit & /htmlpresent)...")
        r_edit = requests.get(edit_url, headers=headers, timeout=25)
        r_html = requests.get(html_url, headers=headers, timeout=25)

        if r_edit.status_code == 200 and r_html.status_code == 200:
            m_img = re.search(r'SK_config\[\'documentImageUrls\'\]\s*=\s*(\{.*?\});', r_edit.text)
            image_urls = json.loads(m_img.group(1)) if m_img else {}

            model_chunks_str = re.findall(r'var modelChunk = (\{.*?\});\s*var modelChunkParseEnd', r_edit.text, re.DOTALL)
            soup = BeautifulSoup(r_html.text, 'html.parser')
            slides_html = soup.find_all(class_='slide')
            total = min(len(model_chunks_str), len(slides_html))

            if total > 0:
                log_cb(f"  [PPTX Tách đối tượng] Phát hiện {total} slide và {len(image_urls)} hình ảnh minh họa độc lập.")

                # Tải song song tất cả các hình ảnh độc lập (high-res transparent PNG)
                downloaded_images = {}
                if image_urls:
                    log_cb(f"  [PPTX Tách đối tượng] Bước 2/3: Đang tải song song {len(image_urls)} ảnh minh họa PNG...")
                    def _fetch_img(item):
                        b_id, u = item
                        try:
                            r = requests.get(u, headers=headers, timeout=20)
                            if r.status_code == 200 and len(r.content) > 100:
                                return b_id, r.content
                        except Exception:
                            pass
                        return b_id, None

                    with ThreadPoolExecutor(max_workers=15) as ex:
                        for b_id, content in ex.map(_fetch_img, image_urls.items()):
                            if content:
                                downloaded_images[b_id] = content
                    log_cb(f"  [PPTX Tách đối tượng] Đã nạp thành công {len(downloaded_images)}/{len(image_urls)} ảnh minh họa.")
                else:
                    log_cb(f"  [PPTX Tách đối tượng] Bước 2/3: Chuẩn bị dựng các khung shape & text...")

                log_cb(f"  [PPTX Tách đối tượng] Bước 3/3: Đang dựng PowerPoint với các đối tượng tách biệt...")
                prs = Presentation()
                prs.slide_width = PptxInches(13.333)
                prs.slide_height = PptxInches(7.5)
                SW_in = 13.333
                SH_in = 7.5
                blank_layout = prs.slide_layouts[6]

                for idx in range(total):
                    slide = prs.slides.add_slide(blank_layout)
                    s_el = slides_html[idx]
                    chunk_data = json.loads(model_chunks_str[idx])
                    chunk = chunk_data.get('chunk', [])

                    # A. Nền Slide riêng biệt (Slide Background Color)
                    bg_color_hex = '#FFFFFF'
                    for op in chunk:
                        if op[0] == 9 and len(op) > 2:
                            for p in op[2]:
                                if isinstance(p, str) and p.startswith('#'):
                                    bg_color_hex = p
                                    break

                    background = slide.background
                    fill = background.fill
                    fill.solid()
                    fill.fore_color.rgb = _hex_to_pptx_rgb(bg_color_hex)

                    # B. Trích xuất phần tử đồ họa từ modelChunk
                    model_elements = []
                    for op in chunk:
                        if op[0] == 3 and len(op) > 2:
                            el_id = op[1]
                            t_id = op[2]
                            trans = op[3]
                            props = op[4]
                            if t_id in [3, 7, 8]:
                                model_elements.append({
                                    'id': el_id,
                                    'type': t_id,
                                    'trans': trans,
                                    'props': props
                                })

                    # C. Phân loại đối tượng từ htmlpresent
                    sc = s_el.find(class_='slide-content')
                    base_w, base_h = 1280.0, 720.0
                    if sc:
                        sc_style = _parse_css_dict(sc.get('style', ''))
                        base_w = _parse_px(sc_style.get('width'), 1280.0)
                        base_h = _parse_px(sc_style.get('height'), 720.0)
                    scale_x = SW_in / base_w
                    scale_y = SH_in / base_h

                    hp_shapes = []
                    for sh in s_el.find_all(class_='shape'):
                        role = sh.get('role', '')
                        title = sh.get('title', '')
                        txt = sh.get_text().strip()
                        st = _parse_css_dict(sh.get('style', ''))
                        w = _parse_px(st.get('width', 0))
                        h = _parse_px(st.get('height', 0))
                        left = _parse_px(st.get('left', 0))
                        top = _parse_px(st.get('top', 0))
                        hp_shapes.append({
                            'role': role,
                            'title': title,
                            'text': txt,
                            'w': w,
                            'h': h,
                            'left': left,
                            'top': top,
                            'el': sh
                        })

                    # D. Thêm các Khung Shape vector riêng biệt (Ovals, Rounded Rectangles, Cards, Badges)
                    used_shape_m_ids = set()
                    for sh in hp_shapes:
                        if sh['role'] != 'img' and not sh['text'] and sh['w'] > 5 and sh['h'] > 5:
                            x_in = sh['left'] * scale_x
                            y_in = sh['top'] * scale_y
                            w_in = sh['w'] * scale_x
                            h_in = sh['h'] * scale_y

                            matched_m = None
                            best_dist = 999999
                            for m in model_elements:
                                if m['id'] in used_shape_m_ids:
                                    continue
                                is_oval = ('ôvan' in sh['title'].lower() and m['type'] == 8)
                                is_round = ('tròn' in sh['title'].lower() and m['type'] == 7)
                                if is_oval or is_round or m['type'] in [7, 8]:
                                    m_x = m['trans'][4] / 36576.0
                                    m_y = m['trans'][5] / 36576.0
                                    dist = (x_in - m_x)**2 + (y_in - m_y)**2
                                    if dist < best_dist:
                                        best_dist = dist
                                        matched_m = m

                            fill_hex = '#EAF4F6' if bg_color_hex == '#FFFFFF' else '#0E7C8C'
                            opacity = 1.0
                            if matched_m:
                                used_shape_m_ids.add(matched_m['id'])
                                props = matched_m['props']
                                for idx_p, p in enumerate(props):
                                    if isinstance(p, str) and p.startswith('#'):
                                        fill_hex = p
                                    if p == 16 and idx_p + 1 < len(props) and isinstance(props[idx_p + 1], (int, float)):
                                        opacity = float(props[idx_p + 1])

                            mso_shape = MSO_SHAPE.OVAL if 'ôvan' in sh['title'].lower() else MSO_SHAPE.ROUNDED_RECTANGLE
                            try:
                                new_shape = slide.shapes.add_shape(
                                    mso_shape,
                                    PptxInches(x_in), PptxInches(y_in),
                                    PptxInches(w_in), PptxInches(h_in)
                                )
                                new_shape.line.fill.background()
                                new_shape.fill.solid()
                                new_shape.fill.fore_color.rgb = _hex_to_pptx_rgb(fill_hex)
                                _apply_shape_opacity(new_shape, opacity)
                            except Exception:
                                pass

                    # E. Thêm các Hình ảnh minh họa riêng biệt (Picture shape PNG trong suốt)
                    used_img_m_ids = set()
                    for sh in hp_shapes:
                        if sh['role'] == 'img':
                            x_in = sh['left'] * scale_x
                            y_in = sh['top'] * scale_y
                            w_in = sh['w'] * scale_x
                            h_in = sh['h'] * scale_y

                            matched_img_m = None
                            best_dist = 999999
                            for m in model_elements:
                                if m['type'] == 3 and m['id'] not in used_img_m_ids:
                                    m_x = m['trans'][4] / 36576.0
                                    m_y = m['trans'][5] / 36576.0
                                    dist = (x_in - m_x)**2 + (y_in - m_y)**2
                                    if dist < best_dist:
                                        best_dist = dist
                                        matched_img_m = m

                            img_bytes = None
                            if matched_img_m:
                                used_img_m_ids.add(matched_img_m['id'])
                                props = matched_img_m['props']
                                for idx_p, p in enumerate(props):
                                    if p == 49 and idx_p + 1 < len(props):
                                        b_id = props[idx_p + 1]
                                        img_bytes = downloaded_images.get(b_id)
                                        break

                            if img_bytes:
                                try:
                                    slide.shapes.add_picture(
                                        io.BytesIO(img_bytes),
                                        PptxInches(x_in), PptxInches(y_in),
                                        PptxInches(w_in), PptxInches(h_in)
                                    )
                                except Exception:
                                    pass

                    # F. Thêm các Hộp Text Box riêng biệt (văn bản giữ đúng font, màu, kích thước, bold/italic)
                    slide_texts = []
                    for sh in hp_shapes:
                        if not sh['text']:
                            continue
                        slide_texts.append(sh['text'])

                        x_in = max(0.0, sh['left'] * scale_x)
                        y_in = max(0.0, sh['top'] * scale_y)
                        w_in = max(0.2, sh['w'] * scale_x)
                        h_in = max(0.2, sh['h'] * scale_y)

                        tb = slide.shapes.add_textbox(PptxInches(x_in), PptxInches(y_in), PptxInches(w_in), PptxInches(h_in))
                        _pptx_set_transparent_fill(tb)
                        tf = tb.text_frame
                        tf.word_wrap = True
                        tf.margin_left = PptxInches(0.02)
                        tf.margin_right = PptxInches(0.02)
                        tf.margin_top = PptxInches(0.01)
                        tf.margin_bottom = PptxInches(0.01)

                        shape_el = sh['el']
                        paragraphs = shape_el.find_all('p') or [shape_el]
                        para_index = 0

                        for p_el in paragraphs:
                            p_style = _parse_css_dict(p_el.get('style', ''))
                            ta = p_style.get('text-align', 'left')
                            p_fs = _parse_pt(p_style.get('font-size'), 14.0)
                            p_fw = p_style.get('font-weight', '400')
                            p_bold = (p_fw in ['bold', '700', '800', '900'] or (p_fw.isdigit() and int(p_fw) >= 700))
                            p_italic = (p_style.get('font-style') == 'italic')
                            p_color = _parse_css_color_rgb(p_style.get('color'))
                            p_ff = p_style.get('font-family', 'Arial').strip('\'"')

                            runs_data = []
                            children = list(p_el.children) if hasattr(p_el, 'children') else [p_el]
                            if not children:
                                children = [p_el]

                            for child in children:
                                if isinstance(child, NavigableString):
                                    raw_text = str(child).replace('\ufffd', '\n').replace('\u000b', '\n').replace('\r', '')
                                    if raw_text:
                                        runs_data.append({
                                            'text': raw_text,
                                            'size': p_fs,
                                            'bold': p_bold,
                                            'italic': p_italic,
                                            'color': p_color,
                                            'font': p_ff
                                        })
                                elif isinstance(child, Tag):
                                    raw_text = child.get_text().replace('\ufffd', '\n').replace('\u000b', '\n').replace('\r', '')
                                    if not raw_text:
                                        continue
                                    c_style = _parse_css_dict(child.get('style', ''))
                                    c_fs = _parse_pt(c_style.get('font-size'), p_fs)
                                    c_fw = c_style.get('font-weight', p_fw)
                                    c_bold = (c_fw in ['bold', '700', '800', '900'] or (c_fw.isdigit() and int(c_fw) >= 700))
                                    c_italic = (c_style.get('font-style') == 'italic') if 'font-style' in c_style else p_italic
                                    c_color = _parse_css_color_rgb(c_style.get('color')) or p_color
                                    c_ff = c_style.get('font-family', p_ff).strip('\'"')

                                    runs_data.append({
                                        'text': raw_text,
                                        'size': c_fs,
                                        'bold': c_bold,
                                        'italic': c_italic,
                                        'color': c_color,
                                        'font': c_ff
                                    })

                            if not runs_data:
                                continue

                            p_para = tf.paragraphs[0] if para_index == 0 else tf.add_paragraph()
                            para_index += 1

                            if ta == 'center':
                                p_para.alignment = PP_ALIGN.CENTER
                            elif ta == 'right':
                                p_para.alignment = PP_ALIGN.RIGHT
                            elif ta == 'justify':
                                p_para.alignment = PP_ALIGN.JUSTIFY
                            else:
                                p_para.alignment = PP_ALIGN.LEFT

                            for r_info in runs_data:
                                run = p_para.add_run()
                                run.text = r_info['text']
                                run.font.size = PptxPt(r_info['size'])
                                run.font.bold = r_info['bold']
                                run.font.italic = r_info['italic']
                                if r_info['font']:
                                    run.font.name = r_info['font']
                                if r_info['color']:
                                    run.font.color.rgb = r_info['color']

                    # G. Ghi chú Speaker Notes
                    if slide_texts:
                        try:
                            notes_tf = slide.notes_slide.notes_text_frame
                            notes_tf.clear()
                            notes_tf.text = f"[Slide {idx + 1}]\n" + "\n\n".join(slide_texts)
                        except Exception:
                            pass

                    if (idx + 1) % 15 == 0 or (idx + 1) == total:
                        log_cb(f"    ✓ Đã hoàn thiện {idx+1}/{total} slide tách đối tượng")

                os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
                prs.save(output_path)
                file_size_mb = os.path.getsize(output_path) / 1024 / 1024
                log_cb(f"  ✓ Hoàn tất PPTX Tách đối tượng: {os.path.basename(output_path)} ({total} slide, {file_size_mb:.2f} MB)")
                log_cb(f"  🌟 TÁCH BIỆT HOÀN TOÀN: Nền slide riêng + Khung shape vector riêng + Hình ảnh riêng + Text layer chỉnh sửa 100%!")
                return output_path
    except Exception as e:
        log_cb(f"  ⚠ Lỗi phương pháp tách đối tượng: {e}. Đang chuyển sang phương thức dự phòng...")

    # -------------------------------------------------------------
    # PHƯƠNG PHÁP 2: DỰ PHÒNG DÙNG ẢNH NỀN SẠCH (showText=0) + TEXT LAYER
    # -------------------------------------------------------------
    try:
        log_cb(f"  [PPTX Dự phòng] Đang tải cấu trúc slide & hình nền sạch...")
        resp = requests.get(html_url, headers=headers, timeout=25)
        if resp.status_code == 200 and len(resp.text) > 1000:
            soup = BeautifulSoup(resp.text, 'html.parser')
            slide_elements = soup.find_all(class_='slide')
            total = len(slide_elements)

            if total > 0:
                bg_urls = {}
                for idx, slide_el in enumerate(slide_elements):
                    sc = slide_el.find(class_='slide-content')
                    if sc:
                        st = _parse_css_dict(sc.get('style', ''))
                        m = re.search(r'url\((.*?)\)', st.get('background-image', ''))
                        if m:
                            bg_urls[idx] = m.group(1).strip('\'"')

                bg_images = {}
                def _fetch_bg(item):
                    s_idx, u = item
                    try:
                        r = requests.get(u, headers=headers, timeout=25)
                        if r.status_code == 200 and len(r.content) > 100:
                            return s_idx, r.content
                    except Exception:
                        pass
                    return s_idx, None

                with ThreadPoolExecutor(max_workers=10) as executor:
                    for s_idx, content in executor.map(_fetch_bg, bg_urls.items()):
                        if content:
                            bg_images[s_idx] = content

                prs = Presentation()
                prs.slide_width = PptxInches(13.333)
                prs.slide_height = PptxInches(7.5)
                SW_in = 13.333
                SH_in = 7.5
                blank_layout = prs.slide_layouts[6]

                for idx, slide_el in enumerate(slide_elements):
                    sc = slide_el.find(class_='slide-content')
                    base_w, base_h = 1280.0, 720.0
                    if sc:
                        sc_style = _parse_css_dict(sc.get('style', ''))
                        base_w = _parse_px(sc_style.get('width'), 1280.0)
                        base_h = _parse_px(sc_style.get('height'), 720.0)

                    scale_x = SW_in / base_w
                    scale_y = SH_in / base_h

                    slide = prs.slides.add_slide(blank_layout)

                    img_bytes = bg_images.get(idx)
                    if img_bytes:
                        try:
                            bg_pic = slide.shapes.add_picture(
                                io.BytesIO(img_bytes), 0, 0,
                                width=prs.slide_width,
                                height=prs.slide_height
                            )
                            spTree = slide.shapes._spTree
                            sp_elem = bg_pic._element
                            spTree.remove(sp_elem)
                            spTree.insert(2, sp_elem)
                        except Exception:
                            pass

                    shapes = slide_el.find_all(class_='shape')
                    slide_texts = []

                    for shape_el in shapes:
                        text_content = shape_el.get_text().strip()
                        if not text_content:
                            continue
                        slide_texts.append(text_content)

                        st = _parse_css_dict(shape_el.get('style', ''))
                        l_px = _parse_px(st.get('left'), 0)
                        t_px = _parse_px(st.get('top'), 0)
                        w_px = _parse_px(st.get('width'), 200)
                        h_px = _parse_px(st.get('height'), 50)

                        l_in = PptxInches(max(0.0, l_px * scale_x))
                        t_in = PptxInches(max(0.0, t_px * scale_y))
                        w_in = PptxInches(max(0.2, w_px * scale_x))
                        h_in = PptxInches(max(0.2, h_px * scale_y))

                        tb = slide.shapes.add_textbox(l_in, t_in, w_in, h_in)
                        _pptx_set_transparent_fill(tb)
                        tf = tb.text_frame
                        tf.word_wrap = True
                        tf.margin_left = PptxInches(0.02)
                        tf.margin_right = PptxInches(0.02)
                        tf.margin_top = PptxInches(0.01)
                        tf.margin_bottom = PptxInches(0.01)

                        paragraphs = shape_el.find_all('p') or [shape_el]
                        para_index = 0

                        for p_el in paragraphs:
                            p_style = _parse_css_dict(p_el.get('style', ''))
                            ta = p_style.get('text-align', 'left')
                            p_fs = _parse_pt(p_style.get('font-size'), 14.0)
                            p_fw = p_style.get('font-weight', '400')
                            p_bold = (p_fw in ['bold', '700', '800', '900'] or (p_fw.isdigit() and int(p_fw) >= 700))
                            p_italic = (p_style.get('font-style') == 'italic')
                            p_color = _parse_css_color_rgb(p_style.get('color'))
                            p_ff = p_style.get('font-family', 'Arial').strip('\'"')

                            runs_data = []
                            children = list(p_el.children) if hasattr(p_el, 'children') else [p_el]
                            if not children:
                                children = [p_el]

                            for child in children:
                                if isinstance(child, NavigableString):
                                    raw_text = str(child).replace('\ufffd', '\n').replace('\u000b', '\n').replace('\r', '')
                                    if raw_text:
                                        runs_data.append({
                                            'text': raw_text,
                                            'size': p_fs,
                                            'bold': p_bold,
                                            'italic': p_italic,
                                            'color': p_color,
                                            'font': p_ff
                                        })
                                elif isinstance(child, Tag):
                                    raw_text = child.get_text().replace('\ufffd', '\n').replace('\u000b', '\n').replace('\r', '')
                                    if not raw_text:
                                        continue
                                    c_style = _parse_css_dict(child.get('style', ''))
                                    c_fs = _parse_pt(c_style.get('font-size'), p_fs)
                                    c_fw = c_style.get('font-weight', p_fw)
                                    c_bold = (c_fw in ['bold', '700', '800', '900'] or (c_fw.isdigit() and int(c_fw) >= 700))
                                    c_italic = (c_style.get('font-style') == 'italic') if 'font-style' in c_style else p_italic
                                    c_color = _parse_css_color_rgb(c_style.get('color')) or p_color
                                    c_ff = c_style.get('font-family', p_ff).strip('\'"')

                                    runs_data.append({
                                        'text': raw_text,
                                        'size': c_fs,
                                        'bold': c_bold,
                                        'italic': c_italic,
                                        'color': c_color,
                                        'font': c_ff
                                    })

                            if not runs_data:
                                continue

                            p_para = tf.paragraphs[0] if para_index == 0 else tf.add_paragraph()
                            para_index += 1

                            if ta == 'center':
                                p_para.alignment = PP_ALIGN.CENTER
                            elif ta == 'right':
                                p_para.alignment = PP_ALIGN.RIGHT
                            elif ta == 'justify':
                                p_para.alignment = PP_ALIGN.JUSTIFY
                            else:
                                p_para.alignment = PP_ALIGN.LEFT

                            for r_info in runs_data:
                                run = p_para.add_run()
                                run.text = r_info['text']
                                run.font.size = PptxPt(r_info['size'])
                                run.font.bold = r_info['bold']
                                run.font.italic = r_info['italic']
                                if r_info['font']:
                                    run.font.name = r_info['font']
                                if r_info['color']:
                                    run.font.color.rgb = r_info['color']

                    if slide_texts:
                        try:
                            notes_tf = slide.notes_slide.notes_text_frame
                            notes_tf.clear()
                            notes_tf.text = f"[Slide {idx + 1}]\n" + "\n\n".join(slide_texts)
                        except Exception:
                            pass

                os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
                prs.save(output_path)
                file_size_mb = os.path.getsize(output_path) / 1024 / 1024
                log_cb(f"  ✓ Hoàn tất PPTX Editable: {os.path.basename(output_path)} ({total} slide, {file_size_mb:.2f} MB)")
                return output_path
    except Exception as e:
        log_cb(f"  ⚠ Lỗi xử lý htmlpresent: {e}. Đang chuyển sang Playwright...")

    # -------------------------------------------------------------
    # PHƯƠNG PHÁP 3: DỰ PHÒNG CUỐI CÙNG BẰNG PLAYWRIGHT SCREENSHOTS
    # -------------------------------------------------------------
    log_cb("  [PPTX Editable] Sử dụng chế độ dự phòng bằng Playwright...")
    return _download_single_presentation_pptx_images(url, output_path, scale=scale, quality=quality, log_cb=log_cb)


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

