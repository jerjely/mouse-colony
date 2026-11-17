# PyInstaller recipe for Mouse Colony (no Python needed on the target computer).
#   Windows: dist/MouseColony/MouseColony.exe        (packed into a Setup .exe by install/windows_installer.iss)
#   macOS:   dist/Mouse Colony.app                   (packed into a .pkg by install/make_pkg.sh)
# Build: pyinstaller --noconfirm --distpath dist --workpath build install/mousecolony.spec
# The GitHub workflow .github/workflows/build-installers.yml runs this on GitHub's Windows and Mac machines.
import re
import sys

HERE = SPECPATH + "/.."
VERSION = re.search(r"^VERSION = '([^']+)'", open(HERE + "/mouse_colony_tool.py", encoding="utf-8").read(), re.M).group(1)
MAC = sys.platform == "darwin"

try:                                        # optional drag & drop support
    from PyInstaller.utils.hooks import collect_data_files
    DND = collect_data_files("tkinterdnd2")
    DND_IMPORTS = ["tkinterdnd2"]
except Exception:
    DND, DND_IMPORTS = [], []
try:                                        # Google Sheets sharing and Google sign-in
    from PyInstaller.utils.hooks import collect_submodules
    GS_IMPORTS = collect_submodules("gspread") + collect_submodules("google.auth") + \
        collect_submodules("google.oauth2") + collect_submodules("google_auth_oauthlib") + \
        collect_submodules("requests_oauthlib") + collect_submodules("oauthlib") + ["requests"]
except Exception:
    GS_IMPORTS = []

DOCS = ["mousecolony.ico", "mousecolony.png", "README.md", "GOOGLE_SETUP.md", "CAGE_CARDS_FORMAT.md",
        "GEMINI_CARD_INSTRUCTIONS.txt"]

a = Analysis(
    [HERE + "/mouse_colony_tool.py"],
    pathex=[HERE],
    datas=[(HERE + "/" + f, ".") for f in DOCS] + DND,
    hiddenimports=["openpyxl", "tkinter", "tkinter.ttk", "tkinter.filedialog", "tkinter.messagebox",
                   "tkinter.simpledialog", "PIL.Image", "PIL.ImageOps"] + DND_IMPORTS + GS_IMPORTS,
    excludes=["PyQt5", "PyQt6", "PySide2", "PySide6", "IPython", "jupyter", "notebook", "matplotlib",
              "numpy", "pandas", "scipy"],
    noarchive=False,
)
pyz = PYZ(a.pure)
exe = EXE(pyz, a.scripts, [], exclude_binaries=True, name="MouseColony", console=False,
          icon=HERE + ("/mousecolony.icns" if MAC else "/mousecolony.ico"), upx=False,
          argv_emulation=False)
coll = COLLECT(exe, a.binaries, a.datas, name="MouseColony", upx=False)
if MAC:
    app = BUNDLE(
        coll,
        name="Mouse Colony.app",
        icon=HERE + "/mousecolony.icns",
        bundle_identifier="hu.pte.volgyilab.mousecolony",
        version=VERSION,
        info_plist={
            "CFBundleName": "Mouse Colony",
            "CFBundleDisplayName": "Mouse Colony",
            "CFBundleShortVersionString": VERSION,
            "CFBundleVersion": VERSION,
            "NSHighResolutionCapable": True,
            "LSApplicationCategoryType": "public.app-category.productivity",
            "NSHumanReadableCopyright": "Volgyi lab, University of Pecs",
        },
    )
