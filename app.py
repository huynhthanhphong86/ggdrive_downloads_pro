"""
Google Drive Downloads Pro - Web Application
Hỗ trợ tải tài liệu Docs, quét toàn bộ Folder Drive và trích xuất PDF View-Only từ Google Drive (đầy đủ 100% trang).
"""

import os
import sys
import json
import time
import queue
import threading
import subprocess
import webbrowser
import asyncio
import requests
from bs4 import BeautifulSoup
from flask import Flask, render_template, request, jsonify, Response

from downloader_core import (
    extract_id_and_type,
    sanitize_filename,
    scan_google_drive_folder,
    download_single_docx,
    download_single_pdf,
    download_drive_pdf,
    download_single_presentation_pptx,
    download_single_presentation_pptx_editable,
    download_single_presentation_pdf
)

if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8')

app = Flask(__name__, template_folder="templates", static_folder="static")

# Thư mục lưu trữ mặc định
USER_HOME = os.path.expanduser("~")
DEFAULT_DOWNLOAD_DIR = os.path.join(USER_HOME, "Downloads", "GoogleDrive_Downloads")
DESKTOP_DIR = os.path.join(USER_HOME, "Desktop", "GoogleDrive_Downloads")
WORKSPACE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Downloads")

os.makedirs(DEFAULT_DOWNLOAD_DIR, exist_ok=True)

state_lock = threading.Lock()
current_task = {
    "is_running": False,
    "total": 0,
    "current": 0,
    "percent": 0,
    "current_file": "",
    "status": "idle",
    "item_status": {},
}

event_queues = []


def broadcast_event(event_type: str, data: dict):
    payload = f"event: {event_type}\ndata: {json.dumps(data, ensure_ascii=False)}\n\n"
    for q in list(event_queues):
        try:
            q.put_nowait(payload)
        except Exception:
            if q in event_queues:
                event_queues.remove(q)


def log_message(msg: str, level: str = "info"):
    timestamp = time.strftime("%H:%M:%S")
    formatted = f"[{timestamp}] {msg}"
    print(formatted)
    broadcast_event("log", {
        "timestamp": timestamp,
        "message": msg,
        "level": level
    })


@app.route("/")
def index():
    presets = {
        "downloads": DEFAULT_DOWNLOAD_DIR,
        "desktop": DESKTOP_DIR,
        "workspace": WORKSPACE_DIR
    }
    return render_template("index.html", default_dir=DEFAULT_DOWNLOAD_DIR, presets=presets)


@app.route("/api/scan", methods=["POST"])
def api_scan():
    data = request.json or {}
    url = data.get("url", "").strip()
    if not url:
        return jsonify({"error": "Vui lòng nhập đường link Google Drive / Docs / Slides"}), 400

    target_id, link_type = extract_id_and_type(url)
    log_message(f"Nhận yêu cầu quét: {url} (Loại: {link_type})")

    try:
        # 1. Thư mục Google Drive
        if link_type == "folder":
            items = asyncio.run(scan_google_drive_folder(url, log_cb=lambda m: log_message(m, "info")))
            folder_title = f"Thư mục Drive ({len(items)} tệp)"
            return jsonify({
                "type": "folder",
                "folderId": target_id,
                "title": folder_title,
                "items": items
            })

        # 2. Google Slides / Bài thuyết trình
        elif link_type == "presentation":
            doc_name = "Bài thuyết trình Google Slides"
            try:
                r = requests.get(f"https://docs.google.com/presentation/d/{target_id}/htmlpresent", headers={'User-Agent': 'Mozilla/5.0'}, timeout=6)
                soup = BeautifulSoup(r.text, 'html.parser')
                if soup.title and soup.title.string:
                    raw = soup.title.string.replace(" - Google Trang trình bày", "").replace(" - Google Slides", "").strip()
                    if raw:
                        doc_name = raw
            except Exception:
                pass

            item = {
                "id": target_id,
                "name": doc_name,
                "isFolder": False,
                "isPdf": False,
                "isPresentation": True,
                "url": f"https://docs.google.com/presentation/d/{target_id}/edit"
            }
            log_message(f"Phát hiện Google Slides: {doc_name} (ID: {target_id})", "success")
            return jsonify({
                "type": "presentation",
                "folderId": None,
                "title": doc_name,
                "items": [item]
            })

        # 3. Tệp Google Drive File (PDF, PPTX, DOCX)
        elif link_type == "drive_file" or "drive.google.com/file" in url:
            doc_name = "Tệp Google Drive"
            is_pdf = False
            is_presentation = False
            try:
                r = requests.get(f"https://drive.google.com/file/d/{target_id}/preview", headers={'User-Agent': 'Mozilla/5.0'}, timeout=6)
                soup = BeautifulSoup(r.text, 'html.parser')
                if soup.title and soup.title.string:
                    raw = soup.title.string.replace(" - Google Drive", "").strip()
                    if raw:
                        doc_name = raw
                        raw_lower = raw.lower()
                        is_presentation = raw_lower.endswith('.pptx') or raw_lower.endswith('.ppt') or 'pptx' in raw_lower or 'trình bày' in raw_lower
                        is_pdf = raw_lower.endswith('.pdf') or 'pdf' in raw_lower
            except Exception:
                pass

            item = {
                "id": target_id,
                "name": doc_name,
                "isFolder": False,
                "isPdf": is_pdf,
                "isPresentation": is_presentation,
                "url": f"https://drive.google.com/file/d/{target_id}/view"
            }
            item_type = "presentation" if is_presentation else ("pdf" if is_pdf else "file")
            log_message(f"Phát hiện tệp Google Drive: {doc_name} (Loại: {item_type}, ID: {target_id})", "success")
            return jsonify({
                "type": item_type,
                "folderId": None,
                "title": doc_name,
                "items": [item]
            })

        # 4. Google Docs
        else:
            doc_name = "Tài liệu Google Docs"
            try:
                r = requests.get(f"https://docs.google.com/document/d/{target_id}/mobilebasic", headers={'User-Agent': 'Mozilla/5.0'}, timeout=5)
                soup = BeautifulSoup(r.text, 'html.parser')
                if soup.title and soup.title.string:
                    raw = soup.title.string.replace(" - Google Docs", "").replace(" - Google Tài liệu", "").strip()
                    if raw:
                        doc_name = raw
            except Exception:
                pass

            item = {
                "id": target_id,
                "name": doc_name,
                "isFolder": False,
                "isPdf": False,
                "isPresentation": False,
                "url": f"https://docs.google.com/document/d/{target_id}/edit"
            }
            log_message(f"Phát hiện tệp Google Docs: {doc_name} (ID: {target_id})", "success")
            return jsonify({
                "type": "doc",
                "folderId": None,
                "title": doc_name,
                "items": [item]
            })
    except Exception as e:
        log_message(f"Lỗi khi quét liên kết: {str(e)}", "error")
        return jsonify({"error": f"Không thể quét liên kết: {str(e)}"}), 500


def background_downloader(items, output_dir, fmt, scale, quality):
    global current_task
    try:
        with state_lock:
            current_task["is_running"] = True
            current_task["total"] = len(items)
            current_task["current"] = 0
            current_task["percent"] = 0
            current_task["status"] = "running"
            current_task["item_status"] = {it["id"]: {"status": "pending", "error_msg": ""} for it in items}

        broadcast_event("task_start", {
            "total": len(items),
            "outputDir": output_dir,
            "format": fmt
        })

        log_message(f"🚀 Bắt đầu quá trình tải {len(items)} tệp vào: {output_dir}", "info")

        os.makedirs(output_dir, exist_ok=True)
        success_count = 0
        fail_count = 0

        for idx, it in enumerate(items):
            item_id = it["id"]
            raw_name = it.get("name", "Document")
            clean_name = sanitize_filename(raw_name)
            
            # Skip subfolders if any
            if it.get("isFolder", False) or "/drive/folders/" in it.get("url", ""):
                log_message(f"⚠️ Bỏ qua thư mục con: {raw_name}", "warning")
                with state_lock:
                    current_task["item_status"][item_id] = {"status": "completed", "error_msg": "Bỏ qua thư mục con"}
                broadcast_event("item_update", {"id": item_id, "status": "completed"})
                continue

            clean_lower = clean_name.lower()
            is_presentation = it.get("isPresentation", False) or clean_lower.endswith(".pptx") or clean_lower.endswith(".ppt") or "presentation" in it.get("url", "")
            is_pdf_file = (not is_presentation) and (it.get("isPdf", False) or clean_lower.endswith(".pdf"))
            is_drive_file = (not is_presentation) and (not is_pdf_file) and (
                it.get("isDriveFile", False) or 
                "drive.google.com/file" in it.get("url", "") or 
                any(clean_lower.endswith(ext) for ext in [".jpg", ".jpeg", ".png", ".gif", ".webp", ".bmp", ".svg", ".zip", ".rar", ".7z", ".mp4", ".mp3", ".wav"])
            )
            
            if clean_name.lower().endswith(".docx"):
                base_name = clean_name[:-5]
            elif clean_name.lower().endswith(".pptx"):
                base_name = clean_name[:-5]
            elif clean_name.lower().endswith(".ppt"):
                base_name = clean_name[:-4]
            elif clean_name.lower().endswith(".pdf"):
                base_name = clean_name[:-4]
            else:
                base_name = clean_name

            with state_lock:
                current_task["current"] = idx + 1
                current_task["current_file"] = raw_name
                current_task["percent"] = int(((idx) / len(items)) * 100)
                current_task["item_status"][item_id]["status"] = "downloading"

            broadcast_event("item_update", {
                "id": item_id,
                "status": "downloading",
                "current": idx + 1,
                "total": len(items),
                "percent": current_task["percent"],
                "currentFile": raw_name
            })

            log_message(f"[{idx+1}/{len(items)}] Đang xử lý: {raw_name}...", "info")

            item_failed = False
            error_details = ""

            # A. Xử lý Google Slides / PowerPoint (.pptx)
            if is_presentation:
                # 1. Tải PPTX Editable (hình ảnh nền + text layer chỉnh sửa được)
                if fmt == "pptx_text":
                    editable_path = os.path.join(output_dir, f"{base_name}_editable.pptx")
                    try:
                        download_single_presentation_pptx_editable(it["url"], editable_path, scale=scale, quality=quality, log_cb=lambda m: log_message(m, "info"))
                    except Exception as e:
                        item_failed = True
                        error_details += f"PPTX Editable lỗi: {str(e)}; "
                        log_message(f"❌ Lỗi tạo PPTX Editable cho '{raw_name}': {e}", "error")

                # 2. Tải PPTX hình ảnh (khi chọn PPTX, Word DOCX, hoặc All)
                elif fmt in ["pptx", "docx", "all"]:
                    pptx_path = os.path.join(output_dir, f"{base_name}.pptx")
                    try:
                        download_single_presentation_pptx(it["url"], pptx_path, scale=scale, quality=quality, log_cb=lambda m: log_message(m, "info"))
                    except Exception as e:
                        item_failed = True
                        error_details += f"PPTX lỗi: {str(e)}; "
                        log_message(f"❌ Lỗi tải PPTX cho '{raw_name}': {e}", "error")

                # 3. Tải PDF cho slide nếu chọn PDF hoặc All
                if fmt in ["pdf", "all"] and not item_failed:
                    pdf_path = os.path.join(output_dir, f"{base_name}.pdf")
                    try:
                        download_single_presentation_pdf(it["url"], pdf_path, scale=scale, quality=quality, log_cb=lambda m: log_message(m, "info"))
                    except Exception as e:
                        item_failed = True
                        error_details += f"PDF lỗi: {str(e)}; "
                        log_message(f"❌ Lỗi tải PDF Slide cho '{raw_name}': {e}", "error")

            # B. Xử lý tệp PDF từ Drive Viewer
            elif is_pdf_file:
                pdf_path = os.path.join(output_dir, f"{base_name}.pdf")
                try:
                    asyncio.run(download_drive_pdf(it["url"], pdf_path, quality=quality, log_cb=lambda m: log_message(m, "info")))
                except Exception as e:
                    item_failed = True
                    error_details = f"Lỗi tải Drive PDF: {str(e)}"
                    log_message(f"❌ {error_details}", "error")

            # C. Xử lý tệp Drive / Hình ảnh / File nhị phân
            elif is_drive_file:
                file_path = os.path.join(output_dir, clean_name)
                try:
                    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
                    dl_url = f"https://drive.usercontent.google.com/download?id={item_id}&export=download"
                    r = requests.get(dl_url, headers=headers, stream=True, timeout=30)
                    if r.status_code == 200:
                        with open(file_path, "wb") as f:
                            for chunk in r.iter_content(chunk_size=65536):
                                if chunk:
                                    f.write(chunk)
                        log_message(f"  ✓ Đã tải tệp Drive: {clean_name}", "success")
                    else:
                        alt_url = f"https://drive.google.com/uc?export=download&id={item_id}"
                        r2 = requests.get(alt_url, headers=headers, stream=True, timeout=30)
                        if r2.status_code == 200 and "html" not in r2.headers.get("content-type", "").lower():
                            with open(file_path, "wb") as f:
                                for chunk in r2.iter_content(chunk_size=65536):
                                    if chunk:
                                        f.write(chunk)
                            log_message(f"  ✓ Đã tải tệp Drive: {clean_name}", "success")
                        else:
                            raise Exception(f"Không thể tải tệp Drive (HTTP {r.status_code})")
                except Exception as e:
                    item_failed = True
                    error_details = f"Lỗi tải tệp: {str(e)}"
                    log_message(f"❌ {error_details}", "error")

            # D. Xử lý tệp Google Docs
            else:
                # 1. Tải DOCX (khi chọn DOCX, PPTX, hoặc All để đảm bảo tải đúng định dạng của tệp trong thư mục)
                if fmt in ["docx", "pptx", "all"]:
                    docx_path = os.path.join(output_dir, f"{base_name}.docx")
                    try:
                        download_single_docx(it["url"], docx_path, log_cb=lambda m: log_message(m, "info"))
                    except Exception as e:
                        item_failed = True
                        error_details += f"DOCX lỗi: {str(e)}; "
                        log_message(f"❌ Lỗi tải DOCX cho '{raw_name}': {e}", "error")

                # 2. Tải PDF nếu chọn PDF hoặc All
                if fmt in ["pdf", "all"] and not item_failed:
                    pdf_path = os.path.join(output_dir, f"{base_name}.pdf")
                    try:
                        asyncio.run(download_single_pdf(it["url"], pdf_path, scale=scale, quality=quality, log_cb=lambda m: log_message(m, "info")))
                    except Exception as e:
                        item_failed = True
                        error_details += f"PDF lỗi: {str(e)}; "
                        log_message(f"❌ Lỗi tải PDF cho '{raw_name}': {e}", "error")

            if item_failed:
                fail_count += 1
                with state_lock:
                    current_task["item_status"][item_id] = {"status": "error", "error_msg": error_details}
                broadcast_event("item_update", {
                    "id": item_id,
                    "status": "error",
                    "error_msg": error_details
                })
            else:
                success_count += 1
                with state_lock:
                    current_task["item_status"][item_id] = {"status": "completed", "error_msg": ""}
                broadcast_event("item_update", {
                    "id": item_id,
                    "status": "completed"
                })
                log_message(f"  ✓ Hoàn thành: {raw_name}", "success")

        with state_lock:
            current_task["status"] = "finished"
            current_task["percent"] = 100

        broadcast_event("task_complete", {
            "successCount": success_count,
            "failCount": fail_count,
            "total": len(items),
            "outputDir": output_dir
        })
        log_message(f"🎉 TẤT CẢ ĐÃ HOÀN TẤT! Thành công: {success_count}/{len(items)}, Thất bại: {fail_count}", "success")
    finally:
        with state_lock:
            current_task["is_running"] = False


@app.route("/api/download", methods=["POST"])
def api_download():
    global current_task
    with state_lock:
        if current_task["is_running"]:
            return jsonify({"error": "Đang có tiến trình tải khác đang chạy. Vui lòng chờ!"}), 400

    data = request.json or {}
    items = data.get("items", [])
    output_dir = data.get("outputDir", "").strip() or DEFAULT_DOWNLOAD_DIR
    fmt = data.get("format", "docx")
    scale = int(data.get("scale", 2))
    quality = int(data.get("quality", 92))

    if not items:
        return jsonify({"error": "Danh sách tệp cần tải trống!"}), 400

    thread = threading.Thread(
        target=background_downloader,
        args=(items, output_dir, fmt, scale, quality),
        daemon=True
    )
    thread.start()

    return jsonify({"success": True, "message": f"Đã bắt đầu tải {len(items)} tệp."})


@app.route("/api/events")
def api_events():
    def event_stream():
        q = queue.Queue()
        event_queues.append(q)
        try:
            with state_lock:
                initial = {
                    "is_running": current_task["is_running"],
                    "total": current_task["total"],
                    "current": current_task["current"],
                    "percent": current_task["percent"],
                    "current_file": current_task["current_file"],
                    "status": current_task["status"],
                    "item_status": current_task["item_status"]
                }
            yield f"event: state\ndata: {json.dumps(initial, ensure_ascii=False)}\n\n"

            while True:
                msg = q.get()
                yield msg
        except GeneratorExit:
            if q in event_queues:
                event_queues.remove(q)

    return Response(event_stream(), mimetype="text/event-stream")


@app.route("/api/browse-directory", methods=["POST"])
def api_browse_directory():
    """Mở hộp thoại chọn thư mục Windows với Win32 Native Shell / Subprocess"""
    data = request.json or {}
    initial_dir = data.get("initialDir", "").strip() or DEFAULT_DOWNLOAD_DIR
    picker_script = os.path.join(os.path.dirname(os.path.abspath(__file__)), "win32_folder_picker.py")
    try:
        cmd = [sys.executable, picker_script]
        if initial_dir and os.path.exists(initial_dir):
            cmd.append(initial_dir)
        res = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            encoding="utf-8",
            timeout=60
        )
        selected_path = res.stdout.strip()
        if selected_path and os.path.isdir(selected_path):
            norm_path = os.path.normpath(selected_path)
            return jsonify({"success": True, "path": norm_path})
        return jsonify({"success": False, "message": "Người dùng đã hủy hoặc không chọn"})
    except subprocess.TimeoutExpired:
        return jsonify({"success": False, "message": "Hết thời gian chờ chọn thư mục"}), 408
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500


@app.route("/api/open-folder", methods=["POST"])
def api_open_folder():
    data = request.json or {}
    path = data.get("path", "").strip() or DEFAULT_DOWNLOAD_DIR
    try:
        norm_path = os.path.abspath(os.path.normpath(path))
        os.makedirs(norm_path, exist_ok=True)
        if sys.platform == 'win32':
            try:
                os.startfile(norm_path)
            except Exception:
                subprocess.Popen(['explorer', norm_path])
        elif sys.platform == 'darwin':
            subprocess.Popen(['open', norm_path])
        else:
            subprocess.Popen(['xdg-open', norm_path])
        return jsonify({"success": True, "path": norm_path})
    except Exception as e:
        return jsonify({"error": str(e)}), 500



def open_browser_delayed():
    time.sleep(1.2)
    webbrowser.open("http://127.0.0.1:5000")


if __name__ == "__main__":
    port = 5000
    print("=" * 60)
    print("🚀 GOOGLE DRIVE DOWNLOADS PRO - GIAO DIỆN WEB ĐANG KHỞI CHẠY")
    print(f"🔗 Mở trình duyệt tại: http://127.0.0.1:{port}")
    print("=" * 60)
    
    threading.Thread(target=open_browser_delayed, daemon=True).start()
    app.run(host="0.0.0.0", port=port, debug=False, use_reloader=False)
