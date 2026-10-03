# Email Folder OneDrive Mirror Backup

## Background

The live working folder for the Word/Outlook automation is:

```text
C:\Users\Loliver\Documents\emails
```

Testing showed that when documents were saved directly into a OneDrive-synchronized folder, Word, Outlook, VBA automation, and OneDrive occasionally interfered with one another. The workflow became reliable after moving the working files out of OneDrive.

The goal is therefore:

- Keep the live working files in a local folder.
- Maintain a mirrored copy in OneDrive.
- Allow remote viewing through OneDrive.
- Avoid OneDrive interacting with active Word documents.
- Keep the solution simple enough to troubleshoot remotely.

## Folder Layout

### Working Folder

```text
C:\Users\Loliver\Documents\emails
```

### OneDrive Mirror

```text
C:\Users\Loliver\OneDrive\EmailBackup
```

### Log File

```text
C:\Users\Loliver\OneDrive\EmailBackupSync.log
```

The log file is intentionally stored outside the mirrored folder.

---

# Batch File

Save as:

```text
C:\Scripts\SyncEmailBackup.bat
```

Contents:

```bat
@echo off

set SOURCE=C:\Users\Loliver\Documents\emails
set DEST=C:\Users\Loliver\OneDrive\EmailBackup
set LOG=C:\Users\Loliver\OneDrive\EmailBackupSync.log

echo ================================================== >> "%LOG%"
echo Backup started %DATE% %TIME% >> "%LOG%"

robocopy "%SOURCE%" "%DEST%" /MIR /R:1 /W:1 /XJ /TEE /LOG+:"%LOG%"

echo Backup finished %DATE% %TIME% >> "%LOG%"
echo. >> "%LOG%"
```

---

# Why Robocopy Instead of XCopy

Both would work for this application.

Robocopy was chosen because:

- Better handling of locked files.
- Built-in folder mirroring with `/MIR`.
- Better logging.
- Automatic skipping of unchanged files.
- Better long-term reliability.

---

# Explanation of Robocopy Switches

## /MIR

Mirrors the source folder into the destination.

This means:

- New files are copied.
- Changed files are updated.
- Renamed files are synchronized.
- Deleted files are removed from the destination.

The OneDrive folder should therefore closely match the local working folder.

## /R:1

Retry failed files once.

## /W:1

Wait one second before retrying.

## /XJ

Excludes junction points and prevents accidental recursive copies.

## /TEE

Displays progress on screen and writes to the log at the same time.

## /LOG+

Appends to the existing log file.

---

# Handling Open Files

If a document is open in Word when the backup runs, Robocopy will normally:

- Skip the locked file.
- Continue backing up other files.
- Record the problem in the log.

At the next scheduled run, the file should be copied once Word releases the lock.

This behavior is preferable to interfering with the active editing session.

---

# Initial Testing

1. Create:

```text
C:\Users\Loliver\OneDrive\EmailBackup
```

2. Run:

```text
C:\Scripts\SyncEmailBackup.bat
```

3. Confirm files appear in:

```text
C:\Users\Loliver\OneDrive\EmailBackup
```

4. Confirm the log file is updated.

5. Repeat the test with a Word document open.

---

# Task Scheduler Setup

Open:

```text
Task Scheduler
```

Select:

```text
Create Task
```

Do not use Create Basic Task.

## General Tab

Name:

```text
Sync Email Backup
```

Enable:

```text
Run whether user is logged on or not
Run with highest privileges
```

## Trigger 1

Create a daily trigger.

Advanced settings:

```text
Repeat task every: 1 hour
For a duration of: Indefinitely
Enabled
```

## Trigger 2

Create a second trigger:

```text
At log on
User: Loliver
```

This causes a backup shortly after startup.

## Actions Tab

Program:

```text
C:\Scripts\SyncEmailBackup.bat
```

## Conditions Tab

Disable:

```text
Start the task only if the computer is on AC power
```

## Settings Tab

Enable:

```text
Run task as soon as possible after a scheduled start is missed
```

Select:

```text
If the task is already running:
Do not start a new instance
```

---

# Startup and Shutdown Behavior

The schedule survives reboots and shutdowns.

Typical operation:

```text
08:00 Computer starts
08:01 Logon trigger runs
09:00 Hourly run
10:00 Hourly run
...
```

If the computer is shut down overnight, the task remains configured and resumes automatically when Windows starts again.

---

# Design Benefits

- Word and Outlook use only local files.
- OneDrive never interacts with active documents.
- Remote access is available through OneDrive.
- Backup logic is independent of the VBA macro.
- Locked files do not stop the backup job.
- Troubleshooting information is retained in a log.
- Easy to maintain remotely through Remote Desktop when required.
