# Sharing the colony with colleagues (Google Sheets)

One person does steps 1-7 once (about 15 minutes). Colleagues then only do step 8.
Everything here is free.

## How it works

- The colony workbook lives in a Google Drive (for example the lab's) as a **Google Sheet**.
- You decide who can use it with the normal Drive **Share** button. Everyone signs in with **their own Google
  account** in the app (once per computer). To add a new colleague, share the sheet with them. To remove someone,
  remove them from the sharing list.
- **Changes are saved on your computer first**, instantly, and uploaded in the background a few seconds later.
  The top bar shows *All saved to Google*, *Uploading...* or *N changes not uploaded yet*. When you quit, anything
  still waiting is uploaded first. Without internet you can keep working: the changes wait on your computer
  (also over a restart) and go up when the connection is back.
- Only the cells you changed are uploaded. If a colleague added rows at the same time (e.g. both of you logged a
  procedure), your rows are placed after theirs, and new mice get the next free ID if theirs took it. Only if you
  both changed **the same cell** does the app ask whose version to keep (yours are kept in a file either way).
- Every 2 minutes (and with **Reload**) the app picks up your colleagues' changes.
- Every change is listed on the sheet's **ChangeLog** tab (when, who with which Google account, which computer).
- Every computer keeps a **local copy** after each change, in *Documents/Mouse colony/backups* by default.
  Without internet the app works on that copy (red banner) and uploads later.

Steps 1-4 give the app a "login ID". Google asks for this from every program that lets people sign in with Google.
It is not a password, and it gives no access to anything by itself.

## 1. Create a Google Cloud project

1. Go to <https://console.cloud.google.com> and sign in, for example with the lab Google account.
2. Top bar: project picker -> **New project** -> name it `Mouse Colony` -> **Create**, then select it.

## 2. Turn on the two APIs

**APIs & Services -> Library**. Search for and **Enable** both:

- **Google Sheets API**
- **Google Drive API** (the app uses it to download the whole sheet in one go)

## 3. Set up the sign-in screen

1. Menu -> **Google Auth Platform** (older consoles call it *OAuth consent screen*) -> **Get started**.
2. App name: `Mouse Colony`. User support e-mail: yours. **Next**.
3. Audience: **External**. **Next**. Contact e-mail: yours. Agree, then **Create**.
4. Go to **Audience** -> **Publish app** -> **Confirm**. The status should say **In production**.

> If you skip step 4 ("Testing"), every sign-in expires after 7 days, and only the people listed as test users
> can sign in.

You don't need to send the app for Google's verification. A small in-house app works for up to 100 people;
Google just shows a warning screen at the first sign-in (see step 7).

## 4. Create the login ID

1. **Google Auth Platform -> Clients -> Create client**.
2. Application type: **Desktop app**. Name: `Mouse Colony`. **Create**.
3. **Download JSON** right away, because Google may not show the secret again later. Keep this file
   (for example `client_secret_....json`).

## 5. Put the workbook into Google Drive as a Google Sheet

1. Close Mouse Colony and Excel.
2. Go to <https://drive.google.com> (the lab account, or your own) -> **New -> File upload** -> choose your
   current `Volgyi_mice_colony.xlsx` (in *Documents/Mouse colony* if you used the installer).
3. Right-click the uploaded file -> **Open with -> Google Sheets**, then **File -> Save as Google Sheets**.
4. Work with this new **Google Sheet** (green icon), not the uploaded .xlsx. You can delete the .xlsx from Drive.

## 6. Share the sheet

In the Google Sheet: **Share** -> add each colleague's Google address as **Editor**, including the account
used on the animal-house computer. Colleagues who only look things up can be **Viewers**: the app shows them
everything, but their changes are refused.

If the sheet is in a personal Drive and that person leaves the lab, they should first make someone else the
owner (Share -> the person -> *Transfer ownership*), or move the sheet to a shared drive.

## 7. Connect your app

In Mouse Colony: **Shared database...**

1. **Your name**: it becomes the default "Person" and appears in the change log.
2. **Local backup folder**: keep the default or choose another folder.
3. **Google Sheet link**: copy it from the browser address bar.
4. **Login ID file**: **Browse...** -> the JSON from step 4.
5. Click **Sign in with Google and connect**. A browser window opens:
   - Choose the Google account that the sheet is shared with.
   - If it says **"Google hasn't verified this app"**: click **Advanced -> Go to Mouse Colony (unsafe)**.
     The app is your lab's own, which is why Google hasn't checked it.
   - Tick all the boxes if Google lists them, then click **Continue**.
   - "Signed in to Mouse Colony" appears. Close the tab.

The top bar then shows *Shared Google Sheet: ...*. The app remembers the sign-in on this computer.

To connect your colleagues, click **Make a connection file for colleagues...** and send them the file
(e-mail is fine: it holds the sheet link and the login ID, no password).

## 8. Colleagues

1. Install Mouse Colony as usual.
2. Drop the connection file onto the window, or use **Shared database -> Use a connection file from a
   colleague**.
3. Sign in with Google in the browser window (same screens as in step 7.5), then enter your name. Done.

If it says the sheet is not shared with you, ask the owner to add your Google address (step 6).

## Good to know

- **What the app may do with your account:** Google's screen asks to "see, edit, create and delete your Google
  Sheets" and "see and download your Google Drive files". Google has no narrower permission for "this one sheet".
  The app only ever opens the sheet in its link. The sign-in is saved on that computer only (in
  `.mouse_colony.json` in your user folder). **Shared database -> Sign out** removes it. You can also revoke it
  at <https://myaccount.google.com/permissions>.
- **Shared computers** (for example the animal house): sign in with the lab account, not a personal one.
- **Editing the Google Sheet in the browser:**
  - Changing a value is fine; the app picks it up within 5 minutes, or straight away with **Reload**.
  - Do **not** insert, delete or sort rows, or rename tabs, in the browser. Use the app for that. For looking
    things up, use *Data -> Create a filter view*, which does not reorder rows for others.
- **Restore:** *Shared database -> Restore from a backup* lists the local copies by date. Restoring makes the
  chosen copy the shared version for everyone. The current version is backed up first, so a restore can be
  undone.
- **Backups kept:** every copy from the last 7 days, then one per day up to 90 days, then one per month.
- **Back to a local file:** *Shared database -> Use a local workbook instead*. A local copy of the shared
  sheet is in the backup folder (`shared_colony_latest.xlsx`).
- **"Sign-in expired":** sign in again in *Shared database*. If it happens every week, step 3.4 (Publish app)
  was skipped.
- **Limits:** Google allows 60 requests per minute per user. One change uses about 3, so this is far more than
  a lab needs, and it is free.

## Alternative: a robot account instead of personal sign-ins

Instead of signing in, the app can also use a *service-account key*: Google Cloud -> IAM & Admin -> Service
accounts -> create one -> Keys -> Add key -> JSON. Then share the sheet with the service account's e-mail address
as Editor, and choose that key file as the "Login ID file" in step 7. Anyone with the key can edit the sheet, so
its connection file must stay inside the lab. Personal sign-ins are the better choice in most labs.
