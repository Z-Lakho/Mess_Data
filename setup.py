from cx_Freeze import setup, Executable

executables = [
    Executable(
        script="main.py",
        base="Win32GUI",
    )
]

setup(
    name="MessManagementApp",
    version="1.0",
    description="Mess Management System",
    options={
        "build_exe": {
            "packages": ["tkinter", "pandas", "openpyxl", "numpy"],
            "include_files": ["data/mess_data.xlsx"],
        }
    },
    executables=executables
)