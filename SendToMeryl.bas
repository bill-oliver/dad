Sub SendToMeryl()

    Dim strFileName As String
    Dim strFolderPath As String
    Dim strFullName As String
    Dim strTempFile As String

'    strFolderPath = "C:\Users\Loliver\OneDrive\Documentation\Les2026\emails\"
    strFolderPath = "C:\Users\boliv\OneDrive\Documents\00Temp\test"

    strFileName = InputBox( _
        "What would you like to call this letter?", _
        "Send to Meryl")

    If Trim(strFileName) = "" Then Exit Sub

    strFileName = Replace(strFileName, "\", "-")
    strFileName = Replace(strFileName, "/", "-")
    strFileName = Replace(strFileName, ":", "-")
    strFileName = Replace(strFileName, "*", "-")
    strFileName = Replace(strFileName, "?", "")
    strFileName = Replace(strFileName, """", "")
    strFileName = Replace(strFileName, "<", "")
    strFileName = Replace(strFileName, ">", "")
    strFileName = Replace(strFileName, "|", "-")

    strTempFile = Environ("TEMP") & "\" & strFileName & ".docx"
    strFullName = strFolderPath & strFileName & ".docx"

    If Dir(strFullName) <> "" Then
        MsgBox _
            "A draft named '" & strFileName & _
            "' already exists." & vbCrLf & vbCrLf & _
            "Please choose a different name.", _
            vbExclamation, _
            "Name Already Used"
        Exit Sub
    End If

    ' Store original setting so we can restore it later
'    originalBackgroundSetting = Options.BackgroundSave
    
    ' FORCE Word to wait for file writing to completely finish
'    Options.BackgroundSave = False

    ' Save the document to a temporary file first
    '
    MsgBox "1"
    ActiveDocument.SaveAs2 _
        FileName:=strTempFile, _
        FileFormat:=wdFormatXMLDocument

    ' Yield execution to the OS to make sure Word handles the file handle change
'    DoEvents
    
    MsgBox "2"
    Dim strSavedFile As String
    strSavedFile = strTempFile
    
    If Dir(strSavedFile) = "" Then
        MsgBox "Could not find saved file:" & vbCrLf & strSavedFile
        Exit Sub
    End If
    
'    ActiveDocument.SendMail


    ' Now save the document to the final location
    '
    MsgBox "3"
    ActiveDocument.SaveAs2 _
        FileName:=strFullName, _
        FileFormat:=wdFormatXMLDocument
    
    ' Yield execution to the OS to make sure Word handles the file handle change
    DoEvents

    ' create the email and attach the temporary file
    '
    MsgBox "4"
    Dim olApp As Object
    Dim olMail As Object
    
    Set olApp = CreateObject("Outlook.Application")
    Set olMail = olApp.CreateItem(0)
    
    With olMail
    
        .To = "meryl.oliver@gmail.com"
        .CC = "bill@oliverassociates.ca"
    
        .Subject = "Draft for Review: " & FileName
    
        .Body = "Hi Meryl," & vbCrLf & vbCrLf & _
                "Please review the attached draft and return your edits." & _
                vbCrLf & vbCrLf & _
                "Thanks."
    
        .Attachments.Add strTempFile
    
        .Display
    
    End With

    MsgBox "5"
    ' Restore user's original settings
    Options.BackgroundSave = originalBackgroundSetting

    ' Clean up temporary file
'    Kill strTempFile

    '  Shutdown word without saving changes
    '
'    ActiveDocument.Close SaveChanges:=False
    Application.Quit SaveChanges:=wdDoNotSaveChanges
    
End Sub
