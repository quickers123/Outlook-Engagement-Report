#Email Engagement Reporting

## Purpose

`ENGAGEMENT_REPORT_SCRIPT.py` creates an Excel report from a local Microsoft Outlook Offline Data File (`.ost`). It is designed to identify email contacts and provide a detailed record of email interactions for one mailbox or a combination of mailboxes contained in the OST.

The script is standalone. It does not require any earlier version of the script.

## What the script produces

Each run creates a timestamped folder named:

```text
External_Contacts_Reports_YYYYMMDD_HHMMSS
```

The folder contains:

- An Excel workbook with three worksheets:
  - **External Contacts**: one row per contact, sorted by `TotalUniqueMessages` from highest to lowest.
  - **Pivot**: an alphabetical PivotTable summary by contact domain and email address.
  - **Email Interactions**: one row per deduplicated email interaction, sorted from newest to oldest.
- `processing_log.json`: processing statistics, selected options, errors, duplicate counts, OST-copy details and workbook-validation results.

Before starting a new run, the script moves earlier `External_Contacts_Reports_*` folders into an `Archive` folder beside the script.

## Requirements

### Supported environment

- Windows
- Desktop Microsoft Excel
- Python 3.10 or later recommended
- Access to the relevant local Outlook profile and OST file
- Enough free disk space to create a working copy of the OST and the report

### Python packages

Install the required packages into the same Python environment used to run the script:

```powershell
python -m pip install libpff-python-windows openpyxl pywin32
```

If `python` is not recognised, use the Python executable or launcher configured on the computer, such as:

```powershell
py -m pip install libpff-python-windows openpyxl pywin32
```

## Important: synchronise Outlook before running the report

This script reads a local Outlook Offline Data File (`.ost`). It does not connect directly to Exchange Online or Outlook on the web.

The report can only include email that has already been downloaded into the local OST by the Outlook desktop application. The email source will therefore only be as current as the last successful desktop Outlook synchronisation.

Before running the script:

1. Open the Outlook desktop application while connected to the organisation's network or the internet.
2. Allow Outlook to finish receiving and synchronising the latest email.
3. If required, use **Send/Receive** > **Send/Receive All Folders** or update the relevant mailbox folders.
4. Confirm that the latest expected messages are visible in desktop Outlook.
5. Close Outlook fully before asking the script to copy and refresh the source OST.

Messages visible only in Outlook on the web, but not yet synchronised into desktop Outlook's local OST, will not appear in the report.

## Recommended folder setup

Place the script in a dedicated working folder, for example:

```text
C:\Users\<username>\secure\
    ENGAGEMENT_REPORT_SCRIPT.py
```

The script creates and manages these subfolders automatically:

```text
Archive\
Source_OST\
External_Contacts_Reports_YYYYMMDD_HHMMSS\
```

Do not place the script inside the Outlook data folder.

## Running the script

### Standard interactive run

1. Open the Outlook desktop application and allow it to synchronise the latest email.
2. Confirm the latest expected messages are visible in desktop Outlook.
3. Close Microsoft Outlook fully before refreshing the source OST. This includes classic Outlook and new Outlook.
4. Open PowerShell or Command Prompt.
5. Change to the folder containing the script:

```powershell
cd "C:\Users\<username>\secure"
```

6. Run:

```powershell
python ENGAGEMENT_REPORT_SCRIPT.py
```

The script will guide the user through the remaining choices.

## Interactive choices

### 1. Update the source email file

The script begins with:

```text
Do you want to update the source email file? [Y/N]:
```

Choose:

- **Y** to search Windows user profiles for Outlook OST files and refresh the working copy.
- **N** to use the largest existing OST in `Source_OST`, or an OST beside the script if `Source_OST` contains none.

### 2. Select the OST

If **Y** is selected, the script searches:

```text
C:\Users\<profile>\AppData\Local\Microsoft\Outlook
```

It lists accessible OST files from largest to smallest and marks the largest as recommended. The largest file is often the main mailbox cache, but the user can select any listed file.

Example:

```text
Outlook email files found (largest first):
  1. Profile: Example.User
     File: example.user@organisation.govt.nz.ost
     Size: 5.2 GB
     Modified: 24 Sep 2026 09:09  [RECOMMENDED: largest file]

Select an OST file [1 recommended], or C to cancel:
```

Press **Enter** to accept the recommended file, enter another listed number, or enter **C** to cancel the refresh.

### 3. OST copy and safety controls

The source OST is copied into `Source_OST`. The script does not rename, move, delete or intentionally modify the source OST.

The script:

- Checks for classic Outlook and new Outlook processes.
- Asks the user to close Outlook before copying.
- Checks available disk space.
- Copies to a temporary `.partial` file.
- Validates that source and copied file sizes match.
- Promotes the completed copy only after validation.
- Preserves the earlier working copy if the refresh fails.
- Uses a read-only streamed-copy fallback if the normal Windows copy encounters a file lock.

### 4. Select mailbox scope

The script inspects the OST and lists detected mailboxes or inboxes.

Choose:

- **1** for one mailbox.
- **2** for a combination of two or more mailboxes.

For a combined report, enter mailbox numbers separated by commas, for example:

```text
1, 3
```

### 5. Select address scope

Choose one of the following:

```text
1. External addresses only
2. Internal addresses only
3. All addresses
```

By default, `tewhatuora.govt.nz` is treated as the internal domain.

### 6. Wait for completion

When processing is complete, the script displays the report folder and row counts:

```text
Done: <report folder>
External contacts: <count>
Interaction rows: <count>
```

## Opening the Excel report

Open the generated `.xlsx` workbook in desktop Excel.

The script writes extracted email content as literal text, disables automatic PivotTable refresh on opening, and validates the workbook package before reporting success. The workbook should therefore open without Excel asking to repair unreadable content.

If Excel still displays a repair warning:

1. Do not treat the warning as an expected step.
2. Retain the workbook and `processing_log.json`.
3. If Excel offers a repair log, save or copy the log details.
4. Record which worksheet or XML part Excel identifies.
5. Report the issue with those files and details.

The workbook does not require macros. An **Enable Content** prompt is not expected as part of the standard workflow. If it appears, do not enable content automatically unless the user has confirmed the workbook and its source are trusted.

## Worksheet guide

### External Contacts

Contains contact-level aggregates, including:

- Contact domain
- Contact email
- Contact type
- First and last engagement
- Inbound message count
- Outbound message count
- Unresolved message count
- Total unique messages
- Mailboxes engaged through, when multiple mailboxes were selected

Rows are ordered by `TotalUniqueMessages` from highest to lowest.

### Pivot

Provides an Excel PivotTable summarising inbound, outbound and total unique messages by contact domain and email address.

The PivotTable remains alphabetically ordered. It is saved with its data and does not refresh automatically when the workbook opens. Use Excel's manual **Refresh** command if required.

### Email Interactions

Contains message-level details, including:

- Date and time
- Contact domain and type
- Direction
- From, To, Cc and Bcc
- Subject
- Plain-text message content
- Content format
- Attachment names and count
- Internet message ID
- Source mailbox
- OST folder path

Rows are ordered from newest to oldest.

## Command-line options

### Use a specific OST

```powershell
python ENGAGEMENT_REPORT_SCRIPT.py --ost "C:\Path\To\mailbox.ost"
```

The supplied path is used when the source-refresh workflow is not selected.

### Set an output parent folder

```powershell
python ENGAGEMENT_REPORT_SCRIPT.py --output "D:\OST Reports"
```

The timestamped report folder is created inside the specified folder.

### Add or replace internal domains

```powershell
python ENGAGEMENT_REPORT_SCRIPT.py --internal-domain organisation.govt.nz
```

Repeat the option to specify multiple internal domains:

```powershell
python ENGAGEMENT_REPORT_SCRIPT.py --internal-domain organisation.govt.nz --internal-domain subsidiary.govt.nz
```

### Include automated addresses

Automated local parts such as `no-reply`, `noreply`, `postmaster` and `mailer-daemon` are excluded by default. To include them:

```powershell
python ENGAGEMENT_REPORT_SCRIPT.py --include-automated
```

### Skip the source-update question

```powershell
python ENGAGEMENT_REPORT_SCRIPT.py --skip-source-prompt
```

This uses the largest existing OST in `Source_OST`, or the largest OST beside the script if no working copy exists in `Source_OST`.

### Combined example

```powershell
python ENGAGEMENT_REPORT_SCRIPT.py --skip-source-prompt --output "D:\OST Reports" --include-automated
```

## Troubleshooting

### Recent email is missing from the report

The script reports only email present in the copied local OST. Reopen the Outlook desktop application, synchronise the relevant mailbox folders, confirm the expected messages are visible in desktop Outlook, close Outlook fully, and rerun the report with the source-update option.

Checking Outlook on the web alone is not sufficient because the script does not read from the web mailbox.

### A source OST is locked

Close classic Outlook and new Outlook, then wait for the applications to finish shutting down before trying again.

If the OST remains locked:

1. Confirm Outlook is not visible in Task Manager.
2. Run the script again.
3. If the lock persists, restart Windows and run the script before reopening Outlook.

The script does not forcibly terminate Outlook.

### No OST files are found

Confirm that:

- The user has an Outlook profile on the computer.
- The OST exists under the profile's local Outlook folder.
- The person running the script can access that profile folder.
- An OST already exists in `Source_OST` if the source-update prompt is skipped.

A specific OST can also be supplied with `--ost`.

### The wrong OST is recommended

The recommendation is based on file size only. Select another listed number if the required mailbox uses a different OST. The script does not delete any candidate OST files.

### Insufficient disk space

The destination drive must have enough free space for the OST working copy plus a safety allowance. Remove unneeded files, select another working location, or free space before rerunning.

### A required Python package is missing

Run:

```powershell
python -m pip install libpff-python-windows openpyxl pywin32
```

Ensure the package installation and script execution use the same Python interpreter.

### The Pivot worksheet is missing

The data workbook is retained if Excel cannot create the native PivotTable. Review the console warning and `processing_log.json`. Confirm that desktop Excel and `pywin32` are available.

### Some message bodies or attachment names are unavailable

The content available depends on what the OST and `libpff` expose. The script records unavailable content or filename fallbacks rather than changing the source file.

### Message content is truncated

Excel limits a cell to 32,767 characters. Longer content is truncated and marked accordingly in the cell.

## Data handling and security

The report may contain email addresses, message metadata, message bodies, attachment names and other potentially sensitive information.

Users should:

- Run the script only on OST files they are authorised to access.
- Store the script, working OST and reports in an approved secure location.
- Apply the organisation's information-management, privacy, retention and access-control requirements.
- Avoid distributing the generated workbook more widely than necessary.
- Securely remove working copies and reports when they are no longer required and policy permits deletion.

## Important limitations

- The script reads a copied local OST and does not connect directly to Exchange Online or Outlook on the web.
- Report currency depends on the most recent successful synchronisation performed by the Outlook desktop application. Email not downloaded into the local OST is outside the report source.
- Results reflect the messages available in the selected OST at the time it was copied.
- Mailbox identification is inferred from OST folders and sampled message headers.
- Duplicate messages are consolidated using the internet message ID where available, with a fallback key otherwise.
- Invitees, recipients or addresses appearing in messages do not by themselves establish a person's role or organisational relationship.
- The report should be reviewed before being used for formal decisions or external distribution.

## Version

This README applies to:

```text
ENGAGEMENT_REPORT_SCRIPT.py
```
