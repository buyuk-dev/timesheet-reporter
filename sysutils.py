def hide_terminal_window():
    try:
        import win32gui, win32con
        The_program_to_hide = win32gui.GetForegroundWindow()
        win32gui.ShowWindow(The_program_to_hide , win32con.SW_HIDE)
    except:
        pass