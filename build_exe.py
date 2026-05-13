import PyInstaller.__main__

PyInstaller.__main__.run(
    [
        "pyinstaller_entry.py",
        "--onefile",
        "--name=visiowings",
        "--console",
        "--specpath=.",
        "--clean",
        "--hidden-import=win32com.client",
        "--hidden-import=pythoncom",
        "--hidden-import=pywintypes",
        "--hidden-import=watchdog.observers",
        "--hidden-import=watchdog.events",
        "--add-data=README.md;.",
        "--add-data=LICENSE;.",
    ]
)
