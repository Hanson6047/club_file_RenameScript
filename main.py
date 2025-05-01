# TODO: implement this file
#!/usr/bin/env python
"""
社團檔案標準化命名與分類系統
=========================
入口點程式
"""

import sys
import tkinter as tk
from src.gui import ClubFileManagerGUI

def main():
    """主程序入口點"""
    root = tk.Tk()
    app = ClubFileManagerGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()