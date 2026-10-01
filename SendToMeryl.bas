Sub SendToMeryl()

    Dim strFileName As String
    Dim strFolderPath As String
    Dim strFullName As String

    strFolderPath = "C:\Users\boliv\Documents\emails\"

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

    ActiveDocument.SaveAs2 _
        FileName:=strFullName, _
        FileFormat:=wdFormatXMLDocument

    ' Yield execution to the OS to make sure Word handles the file handle change
    DoEvents
    
    Dim strSavedFile As String
    strSavedFile = strFullName
    
    If Dir(strSavedFile) = "" Then
        MsgBox "Could not find saved file:" & vbCrLf & strSavedFile
        Exit Sub
    End If
    
    ' create the email and attach the temporary file
    '
    Dim olApp As Object
    Dim olMail As Object
    
    Set olApp = CreateObject("Outlook.Application")
    Set olMail = olApp.CreateItem(0)
    
    With olMail
    
        .To = "meryl.oliver@gmail.com"
        .CC = "bill@oliverassociates.ca"
    
        .Subject = "Draft for Review: " & strFileName
    
        .Body = "Hi Meryl," & vbCrLf & vbCrLf & _
                "Please review the attached draft and return your edits." & _
                vbCrLf & vbCrLf & _
                "Thanks," & _
                vbCrLf & vbCrLf & _
                "Dad."
        .Attachments.Add strFullName
    
        .Display
    
    End With


    ' Restore user's original settings
    Options.BackgroundSave = originalBackgroundSetting

'    Application.Quit SaveChanges:=wdDoNotSaveChanges
    
End Sub
