# Making a new release (installers are built automatically)

GitHub builds the installers on its own Windows and Mac machines whenever a **version tag** is pushed, and attaches
them to a release on the repository's **Releases** page:

| File | For |
|---|---|
| `MouseColony-<version>-Windows-Setup.exe` | Windows 10/11 (64-bit) |
| `MouseColony-<version>-Mac-AppleSilicon.pkg` | Macs with M1, M2, M3, M4... |
| `MouseColony-<version>-Mac-Intel.pkg` | older Intel Macs (also runs on Apple silicon) |

## Steps (GitHub Desktop)

1. Put the new `mouse_colony_tool.py` (and any changed files) into your local copy of the repository.
2. Check the version line near the top of `mouse_colony_tool.py`, e.g. `VERSION = '2.3.3'`.
3. GitHub Desktop: write a summary such as `Version 2.3.3` → **Commit to main** → **Push origin**.
4. **History** tab → right-click that commit → **Create Tag...** → `v2.3.3` (a `v` + exactly the VERSION) → **Create Tag**.
5. **Push origin** again (it now says *Push tags*/*Push origin* - this sends the tag).
6. On github.com → **Actions** tab: "Build installers" runs (about 10-15 minutes). When it is green, the release
   *Mouse Colony v2.3.3* with the three installers is on the **Releases** page (right-hand side of the repository).

If the tag and VERSION don't match, the build stops with a message saying so: fix one of them, delete the tag
(History → right-click the tag → Delete Tag; also delete it on github.com → Releases/Tags if it was pushed), and tag again.

**Test build without a release:** github.com → Actions → *Build installers* → **Run workflow**. The installers are
then downloadable at the bottom of that run's page (*Artifacts*).

## Good to know

- The installers are not signed with a paid Microsoft/Apple certificate, so the first launch shows a warning:
  - Windows: "Windows protected your PC" → *More info* → *Run anyway*.
  - Mac: right-click the `.pkg` → **Open** → **Open** (or System Settings → Privacy & Security → *Open Anyway*).
  Signing would need an Apple Developer account (99 USD/year) and a Windows code-signing certificate; the workflow
  can be extended for that later.
- Windows installs for the current user without admin rights (`%LOCALAPPDATA%\Programs\Mouse Colony`), with a
  Start-menu entry and optionally a desktop icon; uninstall from Windows Settings → Apps. Mac installs
  `Mouse Colony.app` into Applications.
- Installing a new version over the old one keeps everything: settings and the shared-sheet connection
  (`.mouse_colony.json` in the user folder), backups and workbooks (Documents/Mouse colony).
- Each build runs a self-test of the packaged program (all parts present, Tcl/Tk works) before it is released.
- The Intel Mac build uses GitHub's `macos-15-intel` machine. GitHub plans to retire its Intel Macs eventually; when
  that happens, remove the Intel entry in `.github/workflows/build-installers.yml` - Intel Macs can still use the
  `install_mac.command` installer.
- The scripts that do the work: `.github/workflows/build-installers.yml` (the steps), `install/mousecolony.spec`
  (PyInstaller), `install/windows_installer.iss` (Windows Setup, Inno Setup), `install/make_pkg.sh` (Mac package),
  `install/ci_version.py` (version check).
