# Mouse Colony

Breeder ages, relatedness (inbreeding) check and pair suggestions for the Völgyi lab mouse colony.
Works on the workbook `Volgyi_mice_colony.xlsx` (Breeders, Pairs, Cages, Lines, Litters sheets).

## Installing

**Windows:** unzip, then double-click **`install_windows.bat`**. You don't need admin rights.
It finds Python 3.8+ (python.org or Anaconda), creates a private environment in
`%LOCALAPPDATA%\MouseColony`, installs openpyxl and adds a **Mouse Colony** shortcut to the Desktop and
the Start menu. If you don't have Python yet, install it from python.org first and tick
"Add python.exe to PATH".
If Windows shows "Windows protected your PC", click *More info → Run anyway*. To avoid that,
right-click the zip → Properties → *Unblock* before unzipping.

**macOS:** double-click **`install_mac.command`** (if macOS refuses, right-click → Open).
It creates **Mouse Colony** in `~/Applications`, which Spotlight also finds. Use Python from
python.org; Homebrew's Python needs `brew install python-tk` first.

**Linux / Raspberry Pi:** in a terminal in the unzipped folder run `bash install_linux.sh`. It installs Tkinter
with apt if needed (asks for your password once), creates a private environment in `~/.local/share/MouseColony`
and adds **Mouse Colony** to the menu (Office) and the desktop. Raspberry Pi 4 or 5 with Raspberry Pi OS is
recommended; use a screen of at least 1280x720.

**Standalone version for PCs without Python** (e.g. the animal-house computer): after installing on
your own PC, run `install\build_standalone_windows.bat`. It produces `dist\MouseColony\` with
`MouseColony.exe` (about 30 MB). Zip that folder and copy it over; nothing to install there.

**Manual:** `pip install openpyxl`, then `python mouse_colony_tool.py [workbook.xlsx]`.

## The workbook

The installer copies the workbook to `Documents\Mouse colony\` the first time only; reinstalling or
updating never overwrites it. To use a shared copy (OneDrive, network drive), click
**Open workbook...** once: the program remembers the last workbook opened.

Every save from the program first copies the workbook to a `backups` folder next to it.
Close the workbook in Excel before saving from the program.

## Reading cage-card photos (version 2.0)

**Read card photos...** (or drop photos onto the window) reads the cards with Google Gemini, using a free API
key from aistudio.google.com that each computer enters once. The result opens on the usual review screen, where
values Gemini was unsure about start un-ticked. Without a key: attach `GEMINI_CARD_INSTRUCTIONS.txt` and the
photos in the Gemini chat, copy the answer and click **Paste cards** (version 2.1). See **CAGE_CARDS_FORMAT.md**.

## Shared sheet: local first (version 2.3)

In shared mode every change is saved on this computer at once and uploaded in the background a few seconds
later (top bar: *All saved to Google* / *Uploading...* / *N changes not uploaded yet*). Quitting uploads what is
left. Offline work is kept and uploaded later. Rows two people added at the same time are both kept; only
edits of the same cell ask whose version wins.

## Animal-house log photos (version 2.2)

Photos of the EGÉR log page are read like cage cards. **Kivitel** rows terminate the mice, **áthelyezés** rows
move them to the new cage; leválasztás is left to the cage cards. Applied rows are remembered and never applied
twice. See **CAGE_CARDS_FORMAT.md**.

## Shared with colleagues: Google Sheets + local backups (version 1.8, sign-in 1.9)

The colony can live in a shared **Google Sheet** that everyone's app reads and writes - see **GOOGLE_SETUP.md**
(one person sets it up in ~15 minutes; colleagues drop a *connection file* onto the app and **sign in with their
own Google account**). Who has access is decided with the sheet's normal **Share** button.
Before every change the app fetches the latest version and then writes only the changed cells; if a colleague
changed the same cells at the same moment, the change is refused (nothing is overwritten) and the latest version
is shown. The sheet's **ChangeLog** tab records who changed what. The app refreshes every 5 minutes.
Every computer keeps local copies (after each change and at start; all copies of the last 7 days, one per day up
to 90 days, then one per month). Without internet the app keeps working on the local copy (red banner) and uploads later.
**Shared database... -> Restore from a backup** puts an older copy back (the current one is backed up first).

## Weaning and "needs genotyping" (version 1.7)

Genotypes are shown as **+/+** (homo +), **+/-** (het) and **-/-** (homo -, wild type); **?** = not genotyped.
When you wean a litter, the pups are flagged **needs genotyping** unless their genotype is certain from the parents
(e.g. -/- x +/+ gives only +/-). The wean screen shows the expected pups for the chosen pair. Untick *Unknown* to
open the genotyping grid straight after saving. Flagged mice are amber in **All mice**; *Show: Needs genotyping*
lists them, and **Genotyping results...** opens on the flagged mice of the cage. The flag clears by itself when
every allele of the mouse has a result (or untick it in the mouse dialog). Mice added to an empty cage or from
cage cards are flagged the same way when their genotype is not certain.

## Lines: mixes and new lines (version 1.5)

Every line field (cage card, cage, breeder, new pair) has **＋ Combine lines / new line…** at the bottom of its list.
Tick one line to use it, or several to make a cross: the app reuses an existing mix with the same parts (in any
order, e.g. *CarCre* + *GCaMP3-Thy1* = *GCaMP3-Thy1 + CarCre*) or adds a new line whose alleles are the union of
the parts. **Define a new line...** adds a line with its alleles, background, allowed partners and "must keep".
On the Overview tab: *New line...*, *Combine lines (cross)...* and *Edit alleles of selected line...*.
When a new pair is made of two different lines, the pair dialog proposes the mixed line.
The Lines sheet column *Made of* lists the parts of a mixed line (e.g. *Cx36KO + PV-Cre* = `FL-Cx36KO; PV-Cre`),
so combining those parts finds the existing line. *Merge selected line into...* (Overview) renames a duplicate
line everywhere and removes it; an automatic duplicate made by *Combine* is merged on start.

## Individual mice, ear marks, genotyping, card import (version 1.3)

Every stock mouse has its own ID (M0001...) on the *Mice* sheet, with its cage, **ear mark**, birth, parents,
genotype and status. Ear marks are given automatically per cage in this order: none, L, R, LR, L2, R2 (then
L2R, LR2, L2R2). A mark never changes when cage-mates leave; a new mouse gets the next free mark. The marks of the
mice that existed before version 1.3 were assigned automatically - check them against the real ears and correct
them with a double-click on the mouse.

- **All mice** opens each cage to show its mice. Double-click a mouse to change its ear mark, sex, genotype or
  notes, or to move it to another cage.
- **Genotyping results...**: one row per mouse, hom / het / wt / ? per allele, and a "Genotyping" procedure is logged.
  `het*` means "not tested, but certain from the parents".
- **Log procedure** and **Terminate / remove** let you tick the individual mice. Each terminated mouse gets its own
  history row with its procedures.
- The pairing check and the suggestions work on individual mice, so a genotyped het and hom in the same cage are
  told apart.
- **Import cage cards...** (or drop a `.json` file onto the window): see `CAGE_CARDS_FORMAT.md`.
- **Enter cage card...**: type a card by hand (also from the review screen: *Add a card by hand*). Picking a cage
  fills in what the workbook has, so you only change what the card says. Typed cards go through the same review,
  and *Save cards as .json* keeps them as a file.

## Procedures, terminations and genotypes (version 1.2)

**Log procedure** (All mice or Breeders tab): anything done to a cage, part of a cage, one breeder or a whole
breeding cage - injections, surgery, genotyping, imaging... Stored on the *Procedures* sheet.

**Terminate / remove**: choose how many mice of a stock cage (or which animals of a breeding cage) and the
outcome (experiment, tissue, culled over age / genotype / surplus, found dead, transferred). One row goes to the
*Terminated* sheet with line, sex, birth, age, parents, genotype and every procedure that cohort had. The cage
count goes down (or the cage becomes empty). Breeder rows stay with status *Terminated* so pedigrees still work.
A pair that loses its male (or all females) ends; the animal left alive stays in the cage as a stock row.
The **History** tab shows both lists with filters by line, date range and text.

**Genotypes**: each line lists the alleles it carries (Lines sheet, last column, e.g. `Cx36-flox; PV-Cre`).
Breeders and stock cages get hom / het / wt / mixed / ? per allele (edit dialogs). C57Bl6J always counts as wt.
When a cage is not genotyped but its parents make the result certain (wt x hom -> het), the app uses that.
The pairing check and pair suggestions show the expected pups, e.g. `Cx36-flox: 50% hom, 50% het`.

## Updating / removing

Updating: unzip the new version and run the installer again.
Removing: `uninstall_windows.bat` (Windows), or on the Mac delete the app,
`~/Library/Application Support/MouseColony` and `~/.mouse_colony.json`. The workbook is not touched.

If the program closes unexpectedly, the error is in `mouse_colony_error.log` in your user folder.
The Start menu also has a **Mouse Colony (console)** shortcut that shows messages while it runs.
