import sys
import os

if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass


def pick_folder_tkinter(initial_dir=None):
    """Sử dụng Tkinter native Windows Folder Picker - Hiện đại, hỗ trợ gõ đường dẫn, tạo folder"""
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        root.focus_force()

        kwargs = {
            "title": "Chọn thư mục lưu tệp - Google Drive Downloads Pro",
            "parent": root
        }
        if initial_dir and os.path.exists(initial_dir):
            kwargs["initialdir"] = initial_dir

        folder = filedialog.askdirectory(**kwargs)
        root.destroy()
        if folder:
            return os.path.normpath(folder)
    except Exception:
        pass
    return None


def pick_folder_powershell():
    """Sử dụng PowerShell System.Windows.Forms.FolderBrowserDialog làm fallback"""
    try:
        import subprocess
        ps_code = (
            "Add-Type -AssemblyName System.Windows.Forms; "
            "$dialog = New-Object System.Windows.Forms.FolderBrowserDialog; "
            "$dialog.Description = 'Chọn thư mục lưu tệp - Google Drive Downloads Pro'; "
            "$dialog.ShowNewFolderButton = $true; "
            "if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { "
            "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8; "
            "[Console]::WriteLine($dialog.SelectedPath) }"
        )
        res = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps_code],
            capture_output=True,
            text=True,
            encoding="utf-8",
            timeout=50
        )
        out = res.stdout.strip()
        if out and os.path.isdir(out):
            return os.path.normpath(out)
    except Exception:
        pass
    return None


def pick_folder_ctypes():
    """Sử dụng Win32 SHBrowseForFolderW với ctypes 64-bit chuẩn xác"""
    try:
        import ctypes
        from ctypes import wintypes

        ctypes.windll.ole32.CoInitialize(None)

        class BROWSEINFOW(ctypes.Structure):
            _fields_ = [
                ("hwndOwner", wintypes.HWND),
                ("pidlRoot", ctypes.c_void_p),
                ("pszDisplayName", wintypes.LPWSTR),
                ("lpszTitle", wintypes.LPCWSTR),
                ("ulFlags", wintypes.UINT),
                ("lpfn", ctypes.c_void_p),
                ("lParam", ctypes.c_void_p),
                ("iImage", ctypes.c_int)
            ]

        BIF_RETURNONLYFSDIRS = 0x00000001
        BIF_NEWDIALOGSTYLE = 0x00000040
        BIF_EDITBOX = 0x00000010

        bi = BROWSEINFOW()
        bi.hwndOwner = ctypes.windll.user32.GetForegroundWindow()
        bi.lpszTitle = "Chọn thư mục lưu tệp - Google Drive Downloads Pro"
        bi.ulFlags = BIF_RETURNONLYFSDIRS | BIF_NEWDIALOGSTYLE | BIF_EDITBOX
        bi.pidlRoot = None
        bi.pszDisplayName = None
        bi.lpfn = None
        bi.lParam = None
        bi.iImage = 0

        SHBrowseForFolderW = ctypes.windll.shell32.SHBrowseForFolderW
        SHBrowseForFolderW.restype = ctypes.c_void_p
        SHBrowseForFolderW.argtypes = [ctypes.POINTER(BROWSEINFOW)]

        SHGetPathFromIDListW = ctypes.windll.shell32.SHGetPathFromIDListW
        SHGetPathFromIDListW.restype = wintypes.BOOL
        SHGetPathFromIDListW.argtypes = [ctypes.c_void_p, wintypes.LPWSTR]

        CoTaskMemFree = ctypes.windll.ole32.CoTaskMemFree
        CoTaskMemFree.restype = None
        CoTaskMemFree.argtypes = [ctypes.c_void_p]

        pidl = SHBrowseForFolderW(ctypes.byref(bi))
        if pidl:
            buf = ctypes.create_unicode_buffer(wintypes.MAX_PATH)
            success = SHGetPathFromIDListW(pidl, buf)
            CoTaskMemFree(pidl)
            ctypes.windll.ole32.CoUninitialize()
            if success and buf.value:
                return os.path.normpath(buf.value)
        ctypes.windll.ole32.CoUninitialize()
    except Exception:
        pass
    return None


def pick_folder(initial_dir=None):
    # 1. Thử Tkinter (Modern Windows Native File Dialog, đáng tin cậy nhất)
    folder = pick_folder_tkinter(initial_dir)
    if folder:
        return folder

    # 2. Thử ctypes SHBrowseForFolder 64-bit
    folder = pick_folder_ctypes()
    if folder:
        return folder

    # 3. Thử PowerShell FolderBrowserDialog
    folder = pick_folder_powershell()
    if folder:
        return folder

    return None


if __name__ == "__main__":
    init_dir = sys.argv[1] if len(sys.argv) > 1 else None
    selected = pick_folder(init_dir)
    if selected:
        print(selected)

