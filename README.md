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
