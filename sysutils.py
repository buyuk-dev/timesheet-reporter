def hide_terminal_window():
    try:
        import win32gui, win32con
        The_program_to_hide = win32gui.GetForegroundWindow()
        win32gui.ShowWindow(The_program_to_hide , win32con.SW_HIDE)
    except:
        pass


def load_config(path, module_name="config"):
    config_path = os.path.abspath(path)
    spec = importlib.util.spec_from_file_location(module_name, config_path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module
