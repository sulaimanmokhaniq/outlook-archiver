# Outlook Archiver

A small Windows program that copies **exactly the Outlook items you choose** into a local
archive on the laptop, so they can be read with no internet connection.

You choose:

1. **Folders**: any mailbox folder, including sub-folders, the Online Archive and shared mailboxes.
2. **Item types**: emails (and drafts), meeting invitations, calendar items, contacts, tasks and notes.
3. **Time period**: a whole year, "older than N years", the last 12 months, or any From/To dates.
4. **Extra filters**: sender, To/CC, subject, words in the text, category, with or without attachments.
5. **Output** (any combination):
   * **PST file**: an Outlook data file that opens inside Outlook. The original folder structure is kept.
   * **.msg files**: one Outlook file per item; double-click to open.
   * **.mht web pages**: open in any browser, no Outlook needed. Calendar items are saved as `.ics` and contacts as `.vcf`.
   * **PDF**: needs Microsoft Word on the same laptop.
   * **.txt**: plain text.
   * **Attachments** saved as normal files.
   * **Excel list** (`Archive list.csv`): one row per item, ready to sort and filter.

Files are sorted into folders by Outlook folder and year, by item type and year, or by year and month.

> **Nothing is deleted from Outlook.** The tool only makes copies. Every run makes a new
> archive folder, so an older archive is never overwritten.

The "Count matching items" button shows how many items match, by type and by year, before
anything is copied.

## What it needs

* Windows 10 or 11.
* **Classic Outlook**, the version with a **File** menu at the top left, signed in to the
  mailbox. It comes with the Microsoft 365 desktop apps.
  * If Outlook shows a **"New Outlook"** switch at the top right and it is ON, the laptop is
    using the new Outlook. Turn the switch OFF to go back to classic Outlook. The new Outlook
    can't be controlled by other programs, so this tool can't read from it.
* Internet is needed only while archiving items that are not already downloaded to the laptop.
  The finished archive works fully offline.

## Getting the program

### Option A: download the ready-made `.exe`

Every change to this folder is built automatically on GitHub:

1. Open https://github.com/sulaimanmokhaniq/outlook-archiver/actions and choose **Build Outlook Archiver (Windows)** on the left.
2. Open the latest green run and download **OutlookArchiver-windows** at the bottom of the page.
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

This shows the program with a made-up mailbox. It works on any computer and needs no Outlook.

## Using the archive

* **PST**: if "Show it in Outlook when done" was ticked, it already appears in Outlook's folder list.
  Otherwise, in Outlook choose **File → Open & Export → Open Outlook Data File** and pick the `.pst`.
* Each archive folder has a `README.txt` that lists what was saved and any items that had problems.
* Keep a second copy of the archive folder on an external drive.
  Don't keep `.pst` files inside a OneDrive folder; OneDrive does not sync them reliably.

## Tips

* Large mailboxes take time. Archive one year at a time, and leave Outlook open while it runs.
* "Words in the text" and the attachments filter open every item, so they are slower.
* Windows has a 260-character path limit, so long subjects are shortened in file names.
  The full subject is always in the Excel list.
* If Outlook shows "A program is trying to access e-mail address information", the laptop's
  antivirus status is not reported to Outlook. Choose **Allow**, or ask IT.

## For developers

* `archiver/core.py`: filters, item types and file naming. This file doesn't touch Outlook.
* `archiver/base.py`: the steps of an archive run, the Excel list and the README.
* `archiver/outlook_backend.py`: reads and copies items through classic Outlook's COM interface (`pywin32`).
* `archiver/demo_backend.py`: the made-up mailbox for `--demo` and the tests.
* `archiver/gui.py`: the window (Tkinter, part of Python).

Run the tests:

```bash
python -m pytest -q tests
```
