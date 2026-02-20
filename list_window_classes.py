import win32gui

def list_windows():
    print(f"{'HWND':<10} {'Class':<30} {'Title'}")
    print("-" * 60)
    
    def enum_handler(hwnd, ctx):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            cls = win32gui.GetClassName(hwnd)
            try:
                print(f"{hwnd:<10} {cls:<30} {title}")
            except UnicodeEncodeError:
                safe_title = title.encode('ascii', 'ignore').decode('ascii')
                safe_cls = cls.encode('ascii', 'ignore').decode('ascii')
                print(f"{hwnd:<10} {safe_cls:<30} {safe_title} (Encoding error)")
                
    win32gui.EnumWindows(enum_handler, None)

if __name__ == "__main__":
    list_windows()
