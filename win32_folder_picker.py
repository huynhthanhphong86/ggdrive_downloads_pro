import ctypes
from ctypes import wintypes
import sys
import os

if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8')

def pick_folder_win32():
    ctypes.windll.ole32.CoInitialize(None)
    
    class BROWSEINFO(ctypes.Structure):
        _fields_ = [
            ("hwndOwner", wintypes.HWND),
            ("pidlRoot", wintypes.LPARAM),
            ("pszDisplayName", wintypes.LPWSTR),
            ("lpszTitle", wintypes.LPCWSTR),
            ("ulFlags", wintypes.UINT),
            ("lpfn", wintypes.LPARAM),
            ("lParam", wintypes.LPARAM),
            ("iImage", wintypes.INT)
        ]
        
    BIF_RETURNONLYFSDIRS = 0x00000001
    BIF_NEWDIALOGSTYLE = 0x00000040
    BIF_EDITBOX = 0x00000010
    
    bi = BROWSEINFO()
    # Attach to active foreground window
    bi.hwndOwner = ctypes.windll.user32.GetForegroundWindow()
    bi.lpszTitle = "Chọn thư mục lưu tệp - Google Drive Downloads Pro"
    bi.ulFlags = BIF_RETURNONLYFSDIRS | BIF_NEWDIALOGSTYLE | BIF_EDITBOX
    
    pidl = ctypes.windll.shell32.SHBrowseForFolderW(ctypes.byref(bi))
    if pidl:
        path = ctypes.create_unicode_buffer(wintypes.MAX_PATH)
        if ctypes.windll.shell32.SHGetPathFromIDListW(pidl, path):
            val = path.value
            ctypes.windll.ole32.CoTaskMemFree(pidl)
            ctypes.windll.ole32.CoUninitialize()
            return val
        ctypes.windll.ole32.CoTaskMemFree(pidl)
    ctypes.windll.ole32.CoUninitialize()
    return None

if __name__ == "__main__":
    folder = pick_folder_win32()
    if folder:
        print(folder)
