/**
 * GOOGLE DOCS VIEW-ONLY TO PDF EXTRACTOR (CONSOLE SCRIPT)
 * 
 * Hướng dẫn sử dụng:
 * 1. Mở trang Google Docs cần tải trên trình duyệt (Chrome, Edge, Brave, Cốc Cốc,...).
 * 2. Nhấn F12 (hoặc Chuột phải -> Inspect / Kiểm tra) -> Chọn tab "Console".
 * 3. Dán toàn bộ đoạn mã bên dưới vào và nhấn Enter.
 * 4. Tiện ích sẽ tự động cuộn qua tất cả các trang, chụp ảnh Canvas độ nét cao và tải về file PDF.
 */

(async function () {
    console.log("%c[GDoc PDF Extractor] Bắt đầu khởi tạo...", "color: #1a73e8; font-size: 14px; font-weight: bold;");

    // 1. Nạp thư viện jsPDF nếu chưa có
    if (!window.jspdf) {
        console.log("[GDoc PDF Extractor] Đang tải thư viện jsPDF...");
        await new Promise((resolve, reject) => {
            const script = document.createElement('script');
            script.src = 'https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js';
            script.onload = resolve;
            script.onerror = () => reject(new Error("Không thể tải thư viện jsPDF từ CDN"));
            document.head.appendChild(script);
        });
    }

    const { jsPDF } = window.jspdf;

    // 2. Tạo giao diện thông báo nổi (UI Overlay)
    let overlay = document.getElementById('gdoc-extractor-ui');
    if (overlay) overlay.remove();

    overlay = document.createElement('div');
    overlay.id = 'gdoc-extractor-ui';
    overlay.innerHTML = `
        <div style="position: fixed; bottom: 24px; right: 24px; z-index: 999999; background: #1e1e2e; color: #ffffff; padding: 18px 24px; border-radius: 12px; box-shadow: 0 10px 30px rgba(0,0,0,0.5); font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; min-width: 320px; border: 1px solid rgba(255,255,255,0.1); backdrop-filter: blur(10px);">
            <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 12px;">
                <div style="font-weight: 600; font-size: 15px; display: flex; align-items: center; gap: 8px;">
                    <span style="font-size: 18px;">📄</span> GDoc PDF Extractor
                </div>
                <div id="gdoc-status-badge" style="font-size: 12px; background: #2563eb; padding: 2px 8px; border-radius: 20px; font-weight: 500;">Đang quét...</div>
            </div>
            <div id="gdoc-status-text" style="font-size: 13px; color: #a6adc8; margin-bottom: 10px;">Chuẩn bị quét các trang tài liệu...</div>
            <div style="width: 100%; height: 6px; background: #313244; border-radius: 3px; overflow: hidden; margin-bottom: 8px;">
                <div id="gdoc-progress-bar" style="width: 0%; height: 100%; background: linear-gradient(90deg, #3b82f6, #06b6d4); transition: width 0.3s ease;"></div>
            </div>
            <div id="gdoc-progress-num" style="font-size: 11px; color: #6c7086; text-align: right;">0%</div>
        </div>
    `;
    document.body.appendChild(overlay);

    const updateStatus = (text, percent, badge = "Đang quét...") => {
        const textEl = document.getElementById('gdoc-status-text');
        const barEl = document.getElementById('gdoc-progress-bar');
        const numEl = document.getElementById('gdoc-progress-num');
        const badgeEl = document.getElementById('gdoc-status-badge');
        if (textEl) textEl.textContent = text;
        if (barEl) barEl.style.width = `${percent}%`;
        if (numEl) numEl.textContent = `${Math.round(percent)}%`;
        if (badgeEl) badgeEl.textContent = badge;
    };

    // 3. Tìm vùng cuộn của Google Docs
    const editor = document.querySelector('.kix-appview-editor');
    if (!editor) {
        alert("Lỗi: Không tìm thấy vùng hiển thị tài liệu (.kix-appview-editor). Hãy đảm bảo bạn đang mở một Google Docs.");
        overlay.remove();
        return;
    }

    const firstPage = document.querySelector('.kix-page-paginated');
    const pageHeight = firstPage ? (firstPage.offsetHeight || 1123) : 1123;
    const scrollHeight = editor.scrollHeight;
    const step = pageHeight;
    const totalSteps = Math.ceil(scrollHeight / step) + 2;

    const capturedPages = new Map();
    const originalScrollTop = editor.scrollTop;

    console.log(`[GDoc PDF Extractor] Bắt đầu cuộn tài liệu (Tổng chiều cao: ${scrollHeight}px)...`);

    // 4. Cuộn từng bước và trích xuất canvas
    for (let s = 0; s < totalSteps; s++) {
        editor.scrollTop = s * step;
        const progress = Math.min(90, Math.round(((s + 1) / totalSteps) * 90));
        updateStatus(`Đang đọc dữ liệu trang (bước ${s + 1}/${totalSteps})...`, progress, "Đang quét");
        
        await new Promise(r => setTimeout(r, 450)); // Chờ Google Docs vẽ canvas tile

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
                    } catch (err) {
                        console.warn('Lỗi đọc canvas:', err);
                    }
                }
            }
        });
    }

    // Khôi phục vị trí cuộn ban đầu
    editor.scrollTop = originalScrollTop;

    const sortedPages = Array.from(capturedPages.values()).sort((a, b) => a.index - b.index);
    const totalPages = sortedPages.length;

    if (totalPages === 0) {
        alert("Không tìm thấy trang nào để xuất. Vui lòng thử lại!");
        overlay.remove();
        return;
    }

    console.log(`[GDoc PDF Extractor] Đã bắt thành công ${totalPages} trang. Bắt đầu đóng gói PDF...`);
    updateStatus(`Đang đóng gói ${totalPages} trang vào tệp PDF...`, 92, "Đang tạo PDF");

    // 5. Khởi tạo PDF và thêm từng trang
    const firstImg = sortedPages[0];
    const isLandscape = firstImg.width > firstImg.height;
    
    // Đơn vị tính: pt (point)
    const pdf = new jsPDF({
        orientation: isLandscape ? 'landscape' : 'portrait',
        unit: 'pt',
        format: [firstImg.width * 0.75, firstImg.height * 0.75]
    });

    for (let i = 0; i < sortedPages.length; i++) {
        const p = sortedPages[i];
        const percent = 92 + Math.round(((i + 1) / totalPages) * 7);
        updateStatus(`Đang ghép trang ${i + 1} / ${totalPages}...`, percent, "Đang tạo PDF");

        const w = p.width * 0.75;
        const h = p.height * 0.75;

        if (i > 0) {
            pdf.addPage([w, h], w > h ? 'landscape' : 'portrait');
        }

        pdf.addImage(p.data, 'PNG', 0, 0, w, h, undefined, 'FAST');
    }

    // 6. Lấy tiêu đề và tải về
    let docTitle = document.title.replace(" - Google Docs", "").replace(" - Google Tài liệu", "").trim();
    docTitle = docTitle.replace(/[\/\\?%*:|"<>]/g, '').trim() || "GoogleDoc_Download";
    const filename = `${docTitle}.pdf`;

    updateStatus(`Hoàn tất! Đang tải về file: ${filename}`, 100, "Xong!");
    const badgeEl = document.getElementById('gdoc-status-badge');
    if (badgeEl) badgeEl.style.background = '#10b981';

    pdf.save(filename);
    console.log(`%c[GDoc PDF Extractor] Thành công! Đã lưu file: ${filename}`, "color: #10b981; font-size: 14px; font-weight: bold;");

    // Tự đóng bảng thông báo sau 4 giây
    setTimeout(() => {
        if (overlay) overlay.remove();
    }, 4500);
})();
