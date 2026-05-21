import sys

if hasattr(sys, '_MEIPASS'):
    sys.path.append(sys._MEIPASS)

from ui.app import PayrollApp

if __name__ == "__main__":
    PayrollApp().run()