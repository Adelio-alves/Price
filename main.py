# -*- coding: utf-8 -*-
"""
main.py
"""

from app import PriceEditorApp


def main():
    app = PriceEditorApp()
    app.protocol("WM_DELETE_WINDOW", app.on_closing)
    app.mainloop()


if __name__ == "__main__":
    main()