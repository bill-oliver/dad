***note that this document describes some initial attempts to get the macro working in a onedrive folder***
***after a number of attempts I decided to use a folder outside of onedrive***

# Word/Outlook Macro Troubleshooting Session

## Goal
Create a Word template for a 97-year-old user that:

1. Prompts for a document name.
2. Saves to:
   `C:\Users\Loliver\OneDrive\Documentation\Les2026\emails\`
3. Creates an Outlook email to Meryl.
4. CCs Bill.
5. Attaches the saved document.
6. Displays the email.

## Current Macro

```vb
Sub SendToMeryl()
    Dim FileName As String
    Dim FolderPath As String
    Dim FullName As String

    FolderPath = "C:\Users\Loliver\OneDrive\Documentation\Les2026\emails\"

    FileName = InputBox("What would you like to call this letter?", "Send to Meryl")

    If Trim(FileName) = "" Then Exit Sub

    FullName = FolderPath & FileName & ".docx"

    ActiveDocument.SaveAs2 _
        FileName:=FullName, _
        FileFormat:=wdFormatXMLDocument

    ActiveDocument.Save
    DoEvents

    Dim olApp As Object
    Dim olMail As Object

    Set olApp = CreateObject("Outlook.Application")
    Set olMail = olApp.CreateItem(0)

    With olMail
        .To = "meryl.oliver@gmail.com"
        .CC = "bill@oliverassociates.ca"
        .Subject = "Draft for Review: " & FileName
        .Attachments.Add ActiveDocument.FullName
        .Display
    End With

    ActiveDocument.Close SaveChanges:=False
End Sub
```

## Findings

### Confirmed Working
- SaveAs2 succeeds.
- File is saved to OneDrive.
- Outlook.Application object is created.
- Outlook MailItem object is created.

### Failure Point
Occurs at:

```vb
.Attachments.Add ActiveDocument.FullName
```

Error resembles:

> We can't open [file]. It is possible the file is already open.

### Additional Evidence
Attempting:

```vb
FileCopy FullName, TempFile
```

produces:

> Run-time error 70: Permission Denied

This strongly suggests the newly-created document remains locked after SaveAs2.

### Important Discovery
Closing the document appears to release the lock, but closing the document also aborts the running macro.

## Open Questions

1. Is the file lock caused by Word, OneDrive, or both?
2. Can Word's built-in "Share > Email > Send as Attachment" functionality bypass the lock?
3. Is there a way to create a mail message from Word without manually calling Attachments.Add?
4. Can a non-locked temporary copy be generated and attached?

## Environment

Word Version:

Microsoft 365 MSO
Version 2608
Build 16.0.20326.20072
64-bit

OneDrive Path:

`C:\Users\Loliver\OneDrive\Documentation\Les2026\emails\`
