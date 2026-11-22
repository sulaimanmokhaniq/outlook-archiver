# Outlook Archiver

A Windows program that copies **exactly the Outlook items you choose** into an archive on
the laptop, so they can be read with no internet connection. The screen is available in
**English and Arabic**.

## What it does

1. **Folders**: any mailbox folder, including sub-folders, the Online Archive and shared mailboxes.
2. **Item types**: emails and drafts, meeting invitations, calendar, contacts, tasks, notes.
3. **Time period**: this year, last year, the last 12 months, "older than 1/2/3/5 years",
   a specific year, or any From/To dates.
4. **Filters** (optional): sender, To/CC, subject, words in the text, category,
   with or without attachments.
5. **Save as**, in any combination:
   * **PST file**: opens inside Outlook with the same folders.
   * **Separate files** for each item: Outlook `.msg`, web page `.mht`, PDF (needs Word), or text.
     Calendar items are saved as `.ics` and contacts as `.vcf`.
   * **Attachments** as normal files.
   * **Excel list** of every item, ready to sort and filter.
6. **Repeats and clean-up**:
   * **Skip items already archived.** The program remembers what it archived before. If an old
     archive folder is deleted, its items count as not archived again.
   * **Remove from the mailbox after archiving.** This is off by default and must be confirmed
     every time. Details are below.

Click **Count items** before archiving to see how many items match, by type and by year.
**Show the list** displays the matching items themselves. While archiving, the program shows
progress and the time left. At the end, it shows a summary and an **Open archive folder**
button.

Your choices are remembered for next time. The "remove from mailbox" option is the
exception: it always starts off.

## Space, and the check that nothing is missing

**Before starting**, the confirm screen shows:
* how many items are chosen and how big they are;
* the free space on the laptop now, and about how much is left after the archive is saved
  (Start is blocked if it won't fit);
* the size of the mailbox now, and after archiving (it changes only if "Remove from mailbox"
  is on);
* the size of Outlook's own file on the laptop (`.ost`).

**After copying**, the program checks the archive before anything else happens:
* every chosen item is looked up in the finished PST, one by one, by type, subject, sender
  and date;
* the total size in Outlook is compared with the total size in the PST;
* every saved file must exist and not be empty.

The result appears on the final screen, in the archive's README, and in the **Check** column
of the Excel list ("OK", "Not copied", "Not found in the PST", "File missing"). Anything
missing is listed by name. **"Remove from mailbox" runs only if every item was found in the
PST**; if even one is missing, nothing is removed.

## Removing items from the mailbox (optional)

This option exists to free up mailbox space. It is built to be safe:

* Nothing is removed until **every item has been copied into the PST** and the PST has been
  **checked** to hold them all. If the check fails, nothing is removed.
* Removed items go to **Deleted Items**, so they can be restored until that folder is emptied.
* These are **never removed**: contacts, tasks, recurring meetings, meetings that haven't
  happened yet, and anything already in Deleted Items.
* If you click **Stop** during copying, nothing is removed.
* The Excel list marks every item that was removed.

## Works with the new Outlook - no need to switch

If your boss uses the **new Outlook**, he can keep using it. The program starts **classic
Outlook hidden in the background**, uses it to read the mailbox and to make **PST files**, and
closes it again when the program is closed. The screen shows "Classic Outlook in the background".

* Classic Outlook must be **installed** (it comes with the Microsoft 365 desktop apps and is
  usually already on the laptop next to the new Outlook). This is not a limit of this program:
  PST files can only be made with classic Outlook - even the new Outlook's own PST features need it.
* The first time only, if classic Outlook was never set up on the laptop, it may show a window
  asking to sign in with the work account. Finish it once; after that it stays hidden.
* If the "New Outlook" switch stops classic Outlook from starting in the background, the program
  turns the switch off for the few seconds classic Outlook needs to start, then puts it back
  exactly as it was. Your boss keeps the new Outlook.

## Two ways to connect: classic Outlook or Microsoft 365 online

Click the connection button at the top of the screen to choose. The program remembers the choice.

| | Classic Outlook | Microsoft 365 online |
|---|---|---|
| Works with | classic **and new** Outlook (classic runs hidden in the background) | any Outlook, even with classic Outlook not installed |
| How it reads the mailbox | through Outlook on this laptop | straight from Microsoft 365, after you sign in with the work account |
| PST file | ✅ | ❌ (only classic Outlook can make PST files) |
| Original emails | .msg | **.eml** (opens in the new and classic Outlook), calendar **.ics**, contacts **.vcf** |
| Web page / PDF / text / attachments / Excel list | ✅ | ✅ |
| Sizes, check, skip already archived | ✅ | ✅ |
| Remove from mailbox | after the PST is checked | after the original **.eml** files are checked (emails only; calendar, contacts and tasks stay) |
| Internet | only for items not yet downloaded | needed while archiving |

If classic Outlook is not found, the program says so and offers **Use Microsoft 365 online**.

### One-time setup for Microsoft 365 online

Microsoft only lets a program sign in to a work mailbox if the organization has registered
it. This is done once, usually by IT, and takes about 5 minutes:

1. Sign in to https://entra.microsoft.com, then go to **Applications → App registrations → New registration**.
2. Name: `Outlook Archiver`. Supported account types: **this organization only**. Click **Register**.
3. Copy the **Application (client) ID** from the overview page.
4. Go to **Authentication → Add a platform → Mobile and desktop applications**, tick
   `http://localhost`, then click **Configure**.
5. Go to **API permissions → Add a permission → Microsoft Graph → Delegated** and add `User.Read`,
   `Mail.ReadWrite`, `Calendars.Read`, `Contacts.Read` and `Tasks.Read`. If your organization
   requires it, click **Grant admin consent**.
6. In the program, open the connection button, paste the Application (client) ID, and click
   **Use Microsoft 365 online**. A Microsoft sign-in page opens in the browser. Sign in once;
   the program then stays signed in, and the session is stored encrypted for this Windows user.

`Mail.ReadWrite` is only needed for "Remove from mailbox". The program never sends email.

## What it needs

* Windows 10 or 11, with **Microsoft Edge**, which is built into Windows. The screen opens
  in its own Edge window.
* **Either** classic Outlook (the version with a **File** menu at the top left, signed in to
  the mailbox), **or** a Microsoft 365 work account for the online mode (see above).
  * Using the **new Outlook** is fine: classic mode runs classic Outlook hidden in the background
    (see above). Only a laptop with no classic Outlook at all needs the online mode.
* Internet: classic mode needs it only for items not yet downloaded to the laptop; online mode
  needs it while archiving. The finished archive works fully offline either way.

## Getting the program

### Option A: download the ready-made `.exe`

1. Open https://github.com/sulaimanmokhaniq/outlook-archiver/actions and choose
   **Build Outlook Archiver (Windows)** on the left.
2. Open the latest run with a green tick and download **OutlookArchiver-windows** at the
   bottom of the page.
3. Unzip it and double-click `OutlookArchiver.exe`.

Windows may warn that the program is from an unknown publisher, because it isn't signed.
Click **More info → Run anyway**. On a company laptop, IT may need to allow it.

### Option B: run from source

Install Python 3.12 from python.org, then in this folder run:

```bash
pip install -r requirements.txt
```

```bash
python OutlookArchiver.py
```

To build the `.exe` yourself, double-click `build_exe.bat`.

### Try it without Outlook

```bash
python OutlookArchiver.py --demo
```

This opens the program with a made-up mailbox. It works on any computer.

## Using the archive

* **PST**: if "Show it in Outlook when done" was ticked, it already appears in Outlook's folder
  list. Otherwise, in Outlook choose **File → Open & Export → Open Outlook Data File** and pick
  the `.pst`.
* Each archive folder has a README (`README.txt`, or `اقرأني.txt` in Arabic) that lists what
  was saved and any items that had problems.
* Keep a second copy of the archive folder on an external drive. Don't keep `.pst` files inside
  a OneDrive folder; OneDrive does not sync them reliably.

## Privacy and safety

* Everything stays on the laptop. The screen is served only to this laptop (`127.0.0.1`).
  Every request must carry a secret key that only the program's own window knows.
* Settings, the archive history and a technical `log.txt` are kept in
  `%APPDATA%\OutlookArchiver`. Deleting that folder resets the program; it never touches any
  archive.

## Tips

* Large mailboxes take time. Archive one year at a time, and leave Outlook open while it runs.
* "Words in the text" and the attachments filter open every item, so they are slower.
* Windows has a 260-character path limit, so long subjects are shortened in file names.
  The full subject is always in the Excel list.
* If Outlook asks "A program is trying to access e-mail address information", the laptop's
  antivirus status is not reported to Outlook. Choose **Allow**, or ask IT.

## For developers

* `archiver/core.py`: filters, item types, file naming and the removal rules. This file doesn't touch Outlook.
* `archiver/base.py`: the steps of a run (copy, check, remove), the Excel list and the README.
* `archiver/outlook_backend.py`: talks to classic Outlook through COM (`pywin32`).
* `archiver/demo_backend.py`: the made-up mailbox for `--demo` and the tests.
* `archiver/server.py`: the local server behind the screen, and the background job runner.
* `archiver/web/`: the screen (HTML, CSS, JavaScript; English and Arabic).
* `archiver/i18n.py`: the program's own messages in English and Arabic.
* `archiver/storage.py`: remembered settings and archive history.

Run the tests:

```bash
python -m pytest -q tests
```

Options: `--demo` uses the made-up mailbox, `--no-window` only prints the address, and
`--port N` picks the port.
