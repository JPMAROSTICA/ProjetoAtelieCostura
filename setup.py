import sys

from cx_Freeze import setup, Executable

build_exe_options = {"packages": ["os"], "includes": ["tkinter","customtkinter","CTkListbox","CTkMessagebox","datetime"]}

base = None
if sys.platform == "win32":
    base = "Win32GUI"

setup(
    name="AppAjuste",
    version="0.1",
    description="AppAjuste",
    options={"build_exe": build_exe_options},
    executables=[Executable("main.py", base=base)]
)
