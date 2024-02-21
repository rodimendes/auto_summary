import sys
from cx_Freeze import setup, Executable

# Dependencies are automatically detected, but it might need fine tuning.
build_exe_options = {"packages": ["os"], "includes": ["tkinter", "openpyxl", "selenium", "pyautogui"], "include_files": ["registro-sumario.xlsx"]}

# GUI applications require a different base on Windows (the default is for
# a console application).
base = None
if sys.platform == "win32":
    base = "Win32GUI"

setup(
    name="Descanso da esposa",
    version="0.1",
    description="Preenche sumários automaticamente",
    options={"build_exe": build_exe_options},
    executables=[Executable("UI.py", base=base)]
)