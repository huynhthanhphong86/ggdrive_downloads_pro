// Google Drive Downloads Pro - Frontend Reactive Controller

document.addEventListener('DOMContentLoaded', () => {
    // Elements
    const urlInput = document.getElementById('urlInput');
    const btnClearUrl = document.getElementById('btnClearUrl');
    const btnScan = document.getElementById('btnScan');
    const btnDemoSlides = document.getElementById('btnDemoSlides');
    const btnDemoFolder = document.getElementById('btnDemoFolder');
    const btnDemoPdf = document.getElementById('btnDemoPdf');
    const btnDemoFile = document.getElementById('btnDemoFile');

    const outputDirInput = document.getElementById('outputDirInput');
    const btnBrowseDir = document.getElementById('btnBrowseDir');
    const btnOpenDir = document.getElementById('btnOpenDir');
    const presetButtons = document.querySelectorAll('.btn-preset');

    const formatRadios = document.querySelectorAll('.format-radio');
    
    const itemsSection = document.getElementById('itemsSection');
    const folderTitleText = document.getElementById('folderTitleText');
    const selectedCountBadge = document.getElementById('selectedCountBadge');
    const filterInput = document.getElementById('filterInput');
    const btnSelectAll = document.getElementById('btnSelectAll');
    const btnDeselectAll = document.getElementById('btnDeselectAll');
    const btnStartDownload = document.getElementById('btnStartDownload');
    const headerCheckbox = document.getElementById('headerCheckbox');
    const itemsTableBody = document.getElementById('itemsTableBody');

    const progressSection = document.getElementById('progressSection');
    const progressStatusTitle = document.getElementById('progressStatusTitle');
    const currentFileName = document.getElementById('currentFileName');
    const progressPercentText = document.getElementById('progressPercentText');
    const progressBarFill = document.getElementById('progressBarFill');

    const logConsole = document.getElementById('logConsole');
    const btnClearLogs = document.getElementById('btnClearLogs');
    const systemStatus = document.getElementById('systemStatus');
    const systemStatusText = document.getElementById('systemStatusText');

    let scannedItems = [];
    let selectedFormat = "docx";

    // 1. QUICK DEMO BUTTONS
    if (btnDemoSlides) {
        btnDemoSlides.addEventListener('click', () => {
            urlInput.value = "https://docs.google.com/presentation/d/1rmO-DqqLrK86lyS4GZ37cLTb_gNlYqaf/edit";
            urlInput.dispatchEvent(new Event('input'));
            btnScan.click();
        });
    }

    btnDemoFolder.addEventListener('click', () => {
        urlInput.value = "https://drive.google.com/drive/folders/1UAVZMjk0v-f-LY00KzTzCfiLwxwB3fUA";
        urlInput.dispatchEvent(new Event('input'));
        btnScan.click();
    });

    if (btnDemoPdf) {
        btnDemoPdf.addEventListener('click', () => {
            urlInput.value = "https://drive.google.com/file/d/1s-6MsJbKV_i3MQOa3mGbY03YUUKU5Ad_/view";
            urlInput.dispatchEvent(new Event('input'));
            btnScan.click();
        });
    }

    btnDemoFile.addEventListener('click', () => {
        urlInput.value = "https://docs.google.com/document/d/1KN5IuYT_D3wzsx1tJ0rdOQxrmpSs3uCX/edit";
        urlInput.dispatchEvent(new Event('input'));
        btnScan.click();
    });

    // 2. INPUT HANDLERS
    urlInput.addEventListener('input', () => {
        btnClearUrl.style.display = urlInput.value ? 'block' : 'none';
    });

    btnClearUrl.addEventListener('click', () => {
        urlInput.value = '';
        btnClearUrl.style.display = 'none';
        urlInput.focus();
    });

    // 3. FORMAT SELECTION
    function setFormat(fmt) {
        formatRadios.forEach(radio => {
            const val = radio.getAttribute('data-value');
            if (val === fmt) {
                radio.classList.add('active');
                const inp = radio.querySelector('input');
                if (inp) inp.checked = true;
                selectedFormat = fmt;
            } else {
                radio.classList.remove('active');
            }
        });
    }

    formatRadios.forEach(radio => {
        radio.addEventListener('click', () => {
            formatRadios.forEach(r => r.classList.remove('active'));
            radio.classList.add('active');
            selectedFormat = radio.getAttribute('data-value');
        });
    });

    // 4. DIRECTORY PICKER & PRESETS
    presetButtons.forEach(btn => {
        btn.addEventListener('click', () => {
            presetButtons.forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            const targetPath = btn.getAttribute('data-path');
            if (targetPath) {
                outputDirInput.value = targetPath;
                appendLog(new Date().toLocaleTimeString(), `Đã chọn thư mục lưu: ${targetPath}`, 'info');
            }
        });
    });

    btnBrowseDir.addEventListener('click', async () => {
        const controller = new AbortController();
        const timeoutId = setTimeout(() => controller.abort(), 45000);
        
        try {
            btnBrowseDir.disabled = true;
            btnBrowseDir.innerHTML = '<span>⏳ Đang chọn...</span>';
            appendLog(new Date().toLocaleTimeString(), 'Đang mở hộp thoại chọn thư mục Windows...', 'info');

            const resp = await fetch('/api/browse-directory', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ initialDir: outputDirInput.value.trim() }),
                signal: controller.signal
            });
            clearTimeout(timeoutId);
            const data = await resp.json();
            if (data.success && data.path) {
                outputDirInput.value = data.path;
                presetButtons.forEach(b => b.classList.remove('active'));
                appendLog(new Date().toLocaleTimeString(), `✓ Đã chọn thư mục: ${data.path}`, 'success');
            } else if (data.message) {
                appendLog(new Date().toLocaleTimeString(), data.message, 'info');
            }
        } catch (err) {
            console.warn('Duyệt thư mục hoàn tất hoặc đã hủy:', err);
        } finally {
            clearTimeout(timeoutId);
            btnBrowseDir.disabled = false;
            btnBrowseDir.innerHTML = '<span>📁 Duyệt...</span>';
        }
    });

    btnOpenDir.addEventListener('click', async () => {
        const path = outputDirInput.value.trim();
        try {
            btnOpenDir.disabled = true;
            btnOpenDir.innerHTML = '<span>⏳ Đang mở...</span>';
            appendLog(new Date().toLocaleTimeString(), `Đang mở thư mục lưu trữ: ${path}`, 'info');

            const resp = await fetch('/api/open-folder', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ path: path })
            });
            const data = await resp.json();
            if (data.success) {
                appendLog(new Date().toLocaleTimeString(), `✓ Đã mở thư mục trong File Explorer: ${data.path || path}`, 'success');
            } else {
                appendLog(new Date().toLocaleTimeString(), `Không thể mở thư mục: ${data.error || 'Lỗi không xác định'}`, 'error');
                alert('Không thể mở thư mục: ' + (data.error || 'Lỗi'));
            }
        } catch (err) {
            appendLog(new Date().toLocaleTimeString(), `Lỗi khi yêu cầu mở thư mục: ${err.message}`, 'error');
            alert('Không thể mở thư mục: ' + err.message);
        } finally {
            btnOpenDir.disabled = false;
            btnOpenDir.innerHTML = '<span>📂 Mở thư mục</span>';
        }
    });

    const btnOpenDoneDir = document.getElementById('btnOpenDoneDir');
    const progressActions = document.getElementById('progressActions');
    if (btnOpenDoneDir) {
        btnOpenDoneDir.addEventListener('click', () => {
            btnOpenDir.click();
        });
    }

    // 5. SCANNING LOGIC
    btnScan.addEventListener('click', async () => {
        const url = urlInput.value.trim();
        if (!url) {
            alert('Vui lòng nhập liên kết Google Drive, Google Docs hoặc Google Slides!');
            urlInput.focus();
            return;
        }

        btnScan.disabled = true;
        btnScan.querySelector('.btn-text').style.display = 'none';
        btnScan.querySelector('.spinner').style.display = 'inline-block';
        systemStatusText.textContent = 'Đang quét liên kết...';

        try {
            const resp = await fetch('/api/scan', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ url: url })
            });

            const data = await resp.json();
            if (!resp.ok || data.error) {
                throw new Error(data.error || 'Quét thất bại');
            }

            scannedItems = data.items || [];
            
            // Auto-switch format radio based on scanned item type
            if (data.type === 'presentation') {
                setFormat('pptx_text'); // Mặc định: PPTX Editable cho Google Slides
            } else if (data.type === 'doc') {
                setFormat('docx');
            } else if (data.type === 'pdf') {
                setFormat('pdf');
            } else if (data.type === 'folder' && scannedItems.length > 0) {
                const hasPres = scannedItems.some(it => it.isPresentation || (it.name || '').toLowerCase().endsWith('.pptx'));
                const hasDoc = scannedItems.some(it => !it.isPresentation && !it.isPdf && !it.isFolder && ((it.name || '').toLowerCase().endsWith('.docx') || (it.url || '').includes('document')));
                if (hasPres && hasDoc) {
                    setFormat('all');
                } else if (hasPres) {
                    setFormat('pptx_text'); // Folder toàn slides -> mặc định PPTX Editable
                } else if (hasDoc) {
                    setFormat('docx');
                }
            }

            renderItemsTable(scannedItems, data.title || 'Danh sách tệp');
            itemsSection.style.display = 'block';
            itemsSection.scrollIntoView({ behavior: 'smooth' });

            systemStatusText.textContent = `Đã tìm thấy ${scannedItems.length} tệp`;
        } catch (err) {
            alert('Lỗi: ' + err.message);
            systemStatusText.textContent = 'Lỗi quét liên kết';
        } finally {
            btnScan.disabled = false;
            btnScan.querySelector('.btn-text').style.display = 'inline-flex';
            btnScan.querySelector('.spinner').style.display = 'none';
        }
    });

    // 6. RENDER TABLE
    function renderItemsTable(items, title) {
        folderTitleText.textContent = title;
        itemsTableBody.innerHTML = '';

        if (!items || items.length === 0) {
            itemsTableBody.innerHTML = '<tr><td colspan="5" style="text-align: center; color: var(--text-muted); padding: 24px;">Không tìm thấy tệp nào trong liên kết.</td></tr>';
            updateSelectedCount();
            return;
        }

        items.forEach((it, idx) => {
            const name = (it.name || '').toLowerCase();
            const url = it.url || '';
            const isPres = Boolean(!it.isFolder && (it.isPresentation || name.endsWith('.pptx') || name.endsWith('.ppt') || url.includes('presentation')));
            const isPdf = Boolean(!it.isFolder && (!isPres) && (it.isPdf || name.endsWith('.pdf') || url.includes('drive.google.com/file')));
            const isImg = Boolean(!it.isFolder && (!isPres) && (!isPdf) && (name.endsWith('.jpg') || name.endsWith('.jpeg') || name.endsWith('.png') || name.endsWith('.webp') || name.endsWith('.gif')));
            
            let icon = '📄';
            if (it.isFolder) {
                icon = '📁';
            } else if (isPres) {
                icon = '📊';
            } else if (isPdf) {
                icon = '📕';
            } else if (isImg) {
                icon = '🖼️';
            }
            
            const tr = document.createElement('tr');
            tr.setAttribute('data-id', it.id);
            tr.innerHTML = `
                <td><input type="checkbox" class="row-checkbox" value="${it.id}" checked></td>
                <td style="color: var(--text-muted); font-family: var(--font-mono);">${idx + 1}</td>
                <td>
                    <div class="file-name-cell">
                        <span style="font-size: 16px;">${icon}</span>
                        <span>${escapeHtml(it.name || 'Không có tiêu đề')}</span>
                    </div>
                </td>
                <td><span class="file-id-text">${it.id || ''}</span></td>
                <td><span class="status-pill status-pending" id="status-${it.id}">Chờ tải</span></td>
            `;
            itemsTableBody.appendChild(tr);
        });

        document.querySelectorAll('.row-checkbox').forEach(cb => {
            cb.addEventListener('change', updateSelectedCount);
        });

        updateSelectedCount();
    }

    function updateSelectedCount() {
        const checkboxes = document.querySelectorAll('.row-checkbox');
        const checked = document.querySelectorAll('.row-checkbox:checked');
        selectedCountBadge.textContent = `Đã chọn: ${checked.length}/${checkboxes.length} tệp`;
        headerCheckbox.checked = checkboxes.length > 0 && checked.length === checkboxes.length;
        btnStartDownload.disabled = checked.length === 0;
    }

    headerCheckbox.addEventListener('change', () => {
        const checked = headerCheckbox.checked;
        document.querySelectorAll('.row-checkbox').forEach(cb => {
            const tr = cb.closest('tr');
            if (tr && tr.style.display !== 'none') {
                cb.checked = checked;
            }
        });
        updateSelectedCount();
    });

    btnSelectAll.addEventListener('click', () => {
        document.querySelectorAll('.row-checkbox').forEach(cb => cb.checked = true);
        updateSelectedCount();
    });

    btnDeselectAll.addEventListener('click', () => {
        document.querySelectorAll('.row-checkbox').forEach(cb => cb.checked = false);
        updateSelectedCount();
    });

    // 7. SEARCH / FILTER TABLE
    filterInput.addEventListener('input', () => {
        const query = filterInput.value.toLowerCase().trim();
        const rows = itemsTableBody.querySelectorAll('tr');
        rows.forEach(r => {
            const text = r.innerText.toLowerCase();
            r.style.display = text.includes(query) ? '' : 'none';
        });
    });

    // 8. START DOWNLOAD
    btnStartDownload.addEventListener('click', async () => {
        const checkedBoxes = Array.from(document.querySelectorAll('.row-checkbox:checked'));
        if (checkedBoxes.length === 0) {
            alert('Vui lòng chọn ít nhất 1 tệp để tải!');
            return;
        }

        const selectedIds = new Set(checkedBoxes.map(cb => cb.value));
        const selectedItems = scannedItems.filter(it => selectedIds.has(it.id));
        const outputDir = outputDirInput.value.trim();

        btnStartDownload.disabled = true;
        btnScan.disabled = true;

        progressSection.style.display = 'block';
        if (progressActions) progressActions.style.display = 'none';
        progressBarFill.style.width = '0%';
        progressPercentText.textContent = '0%';
        progressStatusTitle.textContent = `Đang bắt đầu tải ${selectedItems.length} tệp...`;
        currentFileName.textContent = 'Đang khởi động tiến trình...';

        selectedItems.forEach(it => {
            const badge = document.getElementById(`status-${it.id}`);
            if (badge) {
                badge.className = 'status-pill status-pending';
                badge.textContent = 'Chờ tải';
            }
        });

        try {
            const resp = await fetch('/api/download', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    items: selectedItems,
                    outputDir: outputDir,
                    format: selectedFormat
                })
            });
            const data = await resp.json();
            if (!resp.ok || data.error) {
                throw new Error(data.error || 'Không thể bắt đầu tải');
            }
        } catch (err) {
            alert('Lỗi: ' + err.message);
            btnStartDownload.disabled = false;
            btnScan.disabled = false;
        }
    });

    // 9. REAL-TIME SERVER-SENT EVENTS (SSE)
    const evtSource = new EventSource('/api/events');

    evtSource.addEventListener('log', (e) => {
        const data = JSON.parse(e.data);
        appendLog(data.timestamp, data.message, data.level);
    });

    evtSource.addEventListener('item_update', (e) => {
        const data = JSON.parse(e.data);
        const badge = document.getElementById(`status-${data.id}`);
        if (badge) {
            if (data.status === 'downloading') {
                badge.className = 'status-pill status-downloading';
                badge.textContent = 'Đang tải...';
                badge.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
            } else if (data.status === 'completed') {
                badge.className = 'status-pill status-completed';
                badge.textContent = '✓ Xong';
            } else if (data.status === 'error') {
                badge.className = 'status-pill status-error';
                badge.textContent = '✕ Lỗi';
                badge.title = data.error_msg || 'Lỗi tải tệp';
            }
        }

        if (data.percent !== undefined) {
            progressBarFill.style.width = `${data.percent}%`;
            progressPercentText.textContent = `${data.percent}%`;
        }
        if (data.current && data.total) {
            progressStatusTitle.textContent = `Đang xử lý: ${data.current} / ${data.total} tệp`;
        }
        if (data.currentFile) {
            currentFileName.textContent = `Tệp: ${data.currentFile}`;
        }
    });

    evtSource.addEventListener('task_complete', (e) => {
        const data = JSON.parse(e.data);
        progressBarFill.style.width = '100%';
        progressPercentText.textContent = '100%';
        progressStatusTitle.textContent = `🎉 Hoàn tất toàn bộ ${data.total} tệp!`;
        currentFileName.textContent = `Thành công: ${data.successCount}, Thất bại: ${data.failCount}`;
        
        if (progressActions) {
            progressActions.style.display = 'block';
        }

        btnStartDownload.disabled = false;
        btnScan.disabled = false;
        systemStatusText.textContent = 'Hoàn thành tất cả tệp';
    });

    function appendLog(timestamp, message, level = 'info') {
        const div = document.createElement('div');
        div.className = `log-entry log-${level}`;
        div.textContent = `[${timestamp}] ${message}`;
        logConsole.appendChild(div);
        logConsole.scrollTop = logConsole.scrollHeight;
    }

    btnClearLogs.addEventListener('click', () => {
        logConsole.innerHTML = '';
    });

    function escapeHtml(str) {
        return (str || '').replace(/[&<>"']/g, m => ({
            '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
        })[m]);
    }
});
