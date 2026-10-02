Sub SendToMeryl()

    '  SendToMeryl.bas:  macro to send the current documnent to meryl and copy me on the email.  
    '  ---------------
    '
    '  The macro should be associated with a button on the ribbon.  
    '  The file is saved in a dedicated folder as a macro-enabled document to presurve the macro for later editing
    '


    Dim strFileName As String
    Dim strFolderPath As String
    Dim strFullName As String

    strFolderPath = "C:\Users\boliv\Documents\emails\"  ' Test location

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

    strFullName = strFolderPath & strFileName & ".docm"

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
        FileFormat:=wdFormatXMLDocumentMacroEnabled

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
