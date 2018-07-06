import os
import sys
import traceback


workdir = os.path.dirname(os.path.abspath(__file__))
sys.path.append(workdir)
os.chdir(workdir)


if __name__ == '__main__':
    try:
        import sysutils
        import gui

        if "DEBUG" not in sys.argv:
            sysutils.hide_terminal_window()

        app = gui.GuiApp()
        app.run()

    except Exception as e:
        traceback.print_exc()
        input("press enter to exit...")
