# VDI Download Recorder (demo)

Two tools. Run both inside the VDI, then send back the output so the download macro can be built.

## 1. env-check.ps1 - what the VDI allows

```
powershell -ExecutionPolicy Bypass -File env-check.ps1 -Url "https://your-site/reports"
```

If scripts are blocked, open PowerShell and paste the file contents in.
It writes `vdi-env-report.txt` to your Desktop. It covers tools present (Power Automate,
AutoHotkey, curl), the execution policy, whether DevTools is blocked, and whether the site
opens with your Windows login without a separate sign-in.

## 2. recorder.js - record the clicks

Pick one way to start it:

- **Bookmarklet:** create a new favorite in Edge, name it `Record`, and paste the whole
  contents of `bookmarklet.txt` as the URL. Open the site and click the favorite.
- **Console:** press F12, open Console, paste `recorder.js`, press Enter.

A red "REC" box appears at the bottom right. Do the download steps the way you normally
would. Use **Note** to explain anything the recorder cannot see, such as "waited 30s".

- Alert pages that open in a new tab start recording by themselves, and keep recording
  when they reload. Only the first tab needs the recorder pasted in.
- If the first tab itself fully reloads, start the recorder again there. The steps so far are kept.
- When finished, click **Save** (saves `recording.json`) or **Copy**.

What gets recorded: page URLs, link and button text, field names, dropdown choices, and
download links on each page. Typed text and passwords are never recorded.

Before sharing the recording, check it for internal URLs or names you do not want to
send outside the company.

If the site sends you to a separate sign-in page on another domain, the recorder treats it
as a separate site. Add a Note such as "SSO login here" instead of recording it.

## 3. ecm-export.js - export transactions for many alerts

Built from the recording. For each alert (ECM) ID it:

1. Opens the alert's **Requirements** page and reads Customer Name, Customer Internal ID and
   Velocity Alert ID.
2. Opens the **Statistics** page, runs **Export Transactions**, and saves the file to Downloads
   as `<ECM ID>_<Velocity Alert ID>_<file name>`.

At the end it saves `ECM_Alert_Info_<date>_<time>.csv` to Downloads with one row per alert:
ECM ID, Velocity Alert ID, Customer Name, Customer Internal ID, Transaction File, Result.
**Copy for Excel** copies the same table, ready to paste into a sheet.

1. Log in to ECM in Edge and open the Dashboard.
2. Start it: Ctrl+Shift+J, paste `ecm-export.js`, Enter. Or add `ecm-export-bookmarklet.txt`
   as a favorite named `Export` and click it.
3. The box is pre-filled with the alert IDs from My Case Queue. Edit the list, or paste your
   own IDs (one per line, or separated by commas).
4. Press **Start**. Each alert shows OK or the reason it failed.

Each alert's Statistics page is loaded in a hidden frame (not a pop-up, so nothing is
blocked). The export is requested from that page so the file can be named with the alert ID.
If that gets no file, the tool clicks the real Export button instead, and that file keeps
ECM's original name.

The first time, Edge may ask to allow this site to download multiple files. Choose **Allow**.
If the session expires, it stops and asks you to log in again.
