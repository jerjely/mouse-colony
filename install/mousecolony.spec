# PyInstaller recipe for a standalone Mouse Colony folder (no Python needed on the target PC).
# Build with install/build_standalone_windows.bat (Windows) or:
#   pyinstaller --noconfirm install/mousecolony.spec
HERE = SPECPATH + "/.."
try:                                        # optional drag & drop support
    from PyInstaller.utils.hooks import collect_data_files
    DND = collect_data_files("tkinterdnd2")
    DND_IMPORTS = ["tkinterdnd2"]
except Exception:
    DND, DND_IMPORTS = [], []
try:                                        # Google Sheets sharing
    from PyInstaller.utils.hooks import collect_submodules
    GS_IMPORTS = collect_submodules("gspread") + collect_submodules("google.auth") + \
        collect_submodules("google.oauth2") + \
        collect_submodules("google_auth_oauthlib") + collect_submodules("requests_oauthlib") + \
        collect_submodules("oauthlib") + ["requests"]
except Exception:
    GS_IMPORTS = []

a = Analysis(
    [HERE + "/mouse_colony_tool.py"],
    pathex=[HERE],
    datas=[(HERE + "/mousecolony.ico", "."), (HERE + "/mousecolony.png", "."), (HERE + "/README.md", "."), (HERE + "/GOOGLE_SETUP.md", ".")] + DND,
    hiddenimports=["openpyxl", "tkinter", "tkinter.ttk", "tkinter.filedialog", "tkinter.messagebox", "tkinter.simpledialog"] + DND_IMPORTS + GS_IMPORTS,
    excludes=["PyQt5", "PyQt6", "PySide2", "PySide6", "IPython", "jupyter", "notebook", "matplotlib",
              "numpy", "pandas", "scipy", "PIL"],
    noarchive=False,
)
pyz = PYZ(a.pure)
exe = EXE(pyz, a.scripts, [], exclude_binaries=True, name="MouseColony",
          console=False, icon=HERE + "/mousecolony.ico", upx=False)
coll = COLLECT(exe, a.binaries, a.datas, name="MouseColony", upx=False)
