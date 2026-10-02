Private const bDebug As Boolean = True   ' Allow template to be opened directly for debugging;

Private Const DRAFT_FOLDER As String = "C:\Users\boliv\Documents\emails\"  ' Test Location
'Private Const DRAFT_FOLDER As String = "C:\Users\Loliver\Documents\emails\"

' When the template is opened directly, prompt for a new or existing draft.
Sub AutoOpen()

    ' Drafts are stored in DRAFT_FOLDER; if we are in that folder, assume we are opening an existing file
    ' Otherwise, start the draft workflow to create a new draft.
    If StrComp(ActiveDocument.Path & "\", DRAFT_FOLDER, vbTextCompare) <> 0 Then
        StartDraftWorkflow
    End If
End Sub

Sub StartDraftWorkflow()

    Dim choice As VbMsgBoxResult
    Dim strFileName As String
    Dim strFullName As String
    Dim existingChoice As VbMsgBoxResult
    Dim fileSystem As Object
    Dim starterDocument As Document

    ' Keep a reference so the unused starter document can be closed if another draft is opened.
    Set starterDocument = ActiveDocument
    choice = MsgBox( _
        "Create a new draft?" & vbCrLf & vbCrLf & _
        "Yes: create a new draft" & vbCrLf & _
        "No: open an existing draft", _
        vbYesNoCancel + vbQuestion, _
        "Drafts for Meryl")

    If choice = vbCancel Then 
        If Not bDebug Then
            starterDocument.Close SaveChanges:=wdDoNotSaveChanges  ' Don't allow access to the template
        End If
        Exit Sub
    End If

    If choice = vbNo Then
        OpenDraft
        If Not ActiveDocument Is starterDocument Then
            starterDocument.Close SaveChanges:=wdDoNotSaveChanges
        End If
        Exit Sub
    End If

    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    If Not fileSystem.FolderExists(DRAFT_FOLDER) Then
        MsgBox "Draft folder not found:" & vbCrLf & DRAFT_FOLDER, vbExclamation
        Exit Sub
    End If

    Do
        strFileName = InputBox("What would you like to call this draft?", "New Draft")
        If Trim$(strFileName) = "" Then Exit Sub

        ' Replace characters that Windows does not allow in file names.
        strFileName = Replace(strFileName, "\", "-")
        strFileName = Replace(strFileName, "/", "-")
        strFileName = Replace(strFileName, ":", "-")
        strFileName = Replace(strFileName, "*", "-")
        strFileName = Replace(strFileName, "?", "")
        strFileName = Replace(strFileName, """", "")
        strFileName = Replace(strFileName, "<", "")
        strFileName = Replace(strFileName, ">", "")
        strFileName = Replace(strFileName, "|", "-")
        strFullName = DRAFT_FOLDER & strFileName & ".docm"

        If Dir$(strFullName) = "" Then Exit Do       ' **** BREAK FROM LOOP IF FILE DOES NOT EXIST ****

        ' 
        '  File already exists; ask the user if they want to reuse it or choose a different name.
        existingChoice = MsgBox( _
            "A draft named '" & strFileName & "' already exists." & vbCrLf & vbCrLf & _
            "Open the existing draft?", _
            vbYesNoCancel + vbQuestion, _
            "Draft Already Exists")

        If existingChoice = vbYes Then
            Documents.Open FileName:=strFullName
            starterDocument.Close SaveChanges:=wdDoNotSaveChanges
            Exit Sub
        ElseIf existingChoice = vbCancel Then
            Exit Sub
        End If
    Loop

    ' Save the draft in Word's macro-enabled document format.
    ActiveDocument.SaveAs2 _
        FileName:=strFullName, _
        FileFormat:=wdFormatXMLDocumentMacroEnabled

End Sub

Sub OpenDraft()

    Dim fd As FileDialog
    Dim strSelectedFile As String

    ' Let the user browse saved drafts in the configured folder.
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "Select a saved draft"
    fd.InitialFileName = DRAFT_FOLDER
    fd.Filters.Clear
    fd.Filters.Add "Draft emails", "*.docm"

    If fd.Show = -1 Then
        strSelectedFile = fd.SelectedItems(1)
        Documents.Open FileName:=strSelectedFile
    End If

End Sub
