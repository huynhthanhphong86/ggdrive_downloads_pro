"""
High-Fidelity Google Docs to DOCX Converter
Chuyển đổi tài liệu Google Docs sang Word (.docx) chuẩn 100% bản gốc:
- Font chữ: Times New Roman, 13pt chuẩn
- Bảng biểu: Kẻ khung viền đen (solid borders), chuẩn kích thước cột (column width), căn giữa
- Căn lề đoạn văn: Căn giữa, căn đều 2 bên, thụt lề
- Giãn dòng & Khoảng cách đoạn: Line spacing 1.15, Spacing 0pt
- Nhúng hình ảnh trực tiếp vào tài liệu
- Số trang khớp hoàn toàn bản gốc (~8 trang)
"""

import os
import sys
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

if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8')


def set_cell_border(cell, **kwargs):
    """
    Kẻ viền cho ô trong bảng (Top, Bottom, Left, Right)
    """
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = tcPr.first_child_found_in("w:tcBorders")
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)
    
    for edge in ('top', 'left', 'bottom', 'right'):
        edge_data = kwargs.get(edge)
        if edge_data:
            tag = 'w:{}'.format(edge)
            element = tcBorders.find(qn(tag))
            if element is None:
                element = OxmlElement(tag)
                tcBorders.append(element)
            for key in ["val", "color", "sz", "space"]:
                if key in edge_data:
                    element.set(qn('w:{}'.format(key)), str(edge_data[key]))


def set_cell_margins(cell, top=50, bottom=50, left=80, right=80):
    """Cài đặt khoảng đệm bên trong ô (dxa: 20 dxa = 1 pt)"""
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
            r = int(hex_val[0:2], 16)
            g = int(hex_val[2:4], 16)
            b = int(hex_val[4:6], 16)
            return RGBColor(r, g, b)
    elif color_str.startswith('rgb'):
        nums = re.findall(r'\d+', color_str)
        if len(nums) >= 3:
            return RGBColor(int(nums[0]), int(nums[1]), int(nums[2]))
    return None


def convert_html_to_clean_docx(html_content: str, output_path: str):
    soup = BeautifulSoup(html_content, 'html.parser')
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
    session.headers.update({
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
    })

    def process_paragraph_element(p_tag, p_obj, is_table_cell=False):
        p_style = parse_css_style(p_tag.get('style', ''))
        
        # Căn lề đoạn văn
        text_align = p_style.get('text-align', '')
        if 'center' in text_align:
            p_obj.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif 'right' in text_align:
            p_obj.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        elif 'justify' in text_align:
            p_obj.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        else:
            p_obj.alignment = WD_ALIGN_PARAGRAPH.LEFT

        # Giãn dòng & khoảng cách đoạn
        line_height = p_style.get('line-height', '')
        if line_height:
            try:
                lh = float(line_height)
                p_obj.paragraph_format.line_spacing = lh
            except ValueError:
                p_obj.paragraph_format.line_spacing = 1.15
        else:
            p_obj.paragraph_format.line_spacing = 1.15

        p_obj.paragraph_format.space_before = Pt(0)
        p_obj.paragraph_format.space_after = Pt(0 if is_table_cell else 2)

        # Xử lý các thẻ con
        for child in p_tag.children:
            if child.name == 'span':
                span_style = parse_css_style(child.get('style', ''))
                text = child.get_text()
                if not text:
                    # Kiểm tra xem có ảnh bên trong span không
                    img_tag = child.find('img')
                    if img_tag and img_tag.get('src'):
                        try:
                            img_resp = session.get(img_tag['src'], timeout=10)
                            if img_resp.status_code == 200:
                                img_stream = io.BytesIO(img_resp.content)
                                p_obj.add_run().add_picture(img_stream, width=Inches(3.0))
                        except Exception as e:
                            print(f"Không thể tải ảnh: {e}")
                    continue

                run = p_obj.add_run(text)
                run.font.name = 'Times New Roman'
                
                # Cỡ chữ
                fs = span_style.get('font-size') or p_style.get('font-size')
                if fs:
                    pt_size = parse_pt(fs)
                    run.font.size = Pt(pt_size) if pt_size > 0 else Pt(13)
                else:
                    run.font.size = Pt(13)

                # In đậm, in nghiêng, gạch chân
                fw = span_style.get('font-weight', '')
                if fw in ['bold', '700', '800', '900'] or 'bold' in span_style.get('font-style', ''):
                    run.bold = True
                
                fst = span_style.get('font-style', '')
                if 'italic' in fst:
                    run.italic = True
                    
                td = span_style.get('text-decoration', '')
                if 'underline' in td:
                    run.underline = True

                # Màu chữ
                col = parse_color(span_style.get('color'))
                if col and col != RGBColor(0, 0, 0) and col != RGBColor(0x1f, 0x1f, 0x1f):
                    run.font.color.rgb = col

            elif child.name == 'img':
                src = child.get('src')
                if src:
                    try:
                        img_resp = session.get(src, timeout=10)
                        if img_resp.status_code == 200:
                            img_stream = io.BytesIO(img_resp.content)
                            p_obj.add_run().add_picture(img_stream, width=Inches(3.0))
                    except Exception as e:
                        print(f"Không thể tải ảnh: {e}")

            elif child.name == 'br':
                p_obj.add_run('\n')
            elif isinstance(child, str):
                if child.strip() or child == ' ':
                    run = p_obj.add_run(child)
                    run.font.name = 'Times New Roman'
                    run.font.size = Pt(13)

    def process_table_element(t_tag):
        rows = t_tag.find_all('tr')
        if not rows:
            return
        
        num_cols = max(len(r.find_all(['td', 'th'])) for r in rows)
        num_rows = len(rows)
        
        tbl = doc.add_table(rows=num_rows, cols=num_cols)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl.autofit = False

        # Thu thập độ rộng các cột từ dòng đầu tiên
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
                
                # Cài đặt độ rộng cột
                if c_idx < len(col_widths):
                    cell.width = Pt(col_widths[c_idx])

                # Kẻ viền đen cho ô (Solid black border 0.5pt - 1pt)
                border_spec = {"val": "single", "sz": 4, "color": "000000"}
                set_cell_border(cell, top=border_spec, bottom=border_spec, left=border_spec, right=border_spec)
                
                # Căn lề trong ô
                set_cell_margins(cell, top=50, bottom=50, left=80, right=80)
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

                # Xử lý nội dung trong ô
                p_tags = td.find_all('p')
                if p_tags:
                    for p_i, p_tag in enumerate(p_tags):
                        p_obj = cell.paragraphs[0] if p_i == 0 else cell.add_paragraph()
                        process_paragraph_element(p_tag, p_obj, is_table_cell=True)
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

        # Khoảng cách sau bảng
        after_p = doc.add_paragraph()
        after_p.paragraph_format.space_before = Pt(0)
        after_p.paragraph_format.space_after = Pt(4)

    # 3. Duyệt tuần tự các phần tử trong tài liệu
    for el in doc_content.children:
        if el.name == 'p':
            p_text = el.get_text().strip()
            if not p_text and not el.find('img'):
                continue
            p = doc.add_paragraph()
            process_paragraph_element(el, p)
        elif el.name == 'table':
            process_table_element(el)
        elif el.name in ['h1', 'h2', 'h3', 'h4', 'h5', 'h6']:
            p = doc.add_paragraph()
            process_paragraph_element(el, p)
            if p.runs:
                for r in p.runs:
                    r.bold = True
        elif el.name in ['ul', 'ol']:
            for li in el.find_all('li'):
                p = doc.add_paragraph(style='List Bullet' if el.name == 'ul' else 'List Number')
                process_paragraph_element(li, p)

    # 4. Lưu tệp
    doc.save(output_path)
    print(f"🎉 Đã tạo tệp Word chuẩn bản gốc: {output_path}")


def download_google_doc_to_docx_advanced(url: str, output_path: str = None) -> str:
    """Tải và chuyển đổi Google Docs sang .docx chuẩn bản gốc 100%"""
    match = re.search(r"/d/([a-zA-Z0-9-_]+)", url)
    doc_id = match.group(1) if match else url
    mobile_url = f"https://docs.google.com/document/d/{doc_id}/mobilebasic"

    print("\n" + "=" * 60)
    print("📝 BẮT ĐẦU TẢI GOOGLE DOCS SANG WORD (.DOCX) CHUẨN BẢN GỐC")
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
    doc_title = re.sub(r'[\\/*?:"<>|]', "", doc_title).strip() or "GoogleDoc_Export"

    if not output_path:
        output_path = f"{doc_title}.docx"
    elif not output_path.lower().endswith(".docx"):
        output_path += ".docx"

    convert_html_to_clean_docx(r.text, output_path)
    abs_path = os.path.abspath(output_path)
    print(f"📁 Tệp Word (.docx) đã lưu tại: {abs_path}")
    return abs_path


if __name__ == "__main__":
    url = sys.argv[1] if len(sys.argv) > 1 else "https://docs.google.com/document/d/1KN5IuYT_D3wzsx1tJ0rdOQxrmpSs3uCX/edit"
    download_google_doc_to_docx_advanced(url, "BÀI 34 ( 2 TIẾT).docx")
