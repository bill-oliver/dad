#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Sub SendToMeryl()

    Dim strFileName As String
    Dim strFolderPath As String
    Dim strFullName As String
'    Dim strTempFile As String

'    strFolderPath = "C:\Users\Loliver\OneDrive\Documentation\Les2026\emails\"
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

'    strTempFile = Environ("TEMP") & "\" & strFileName & ".docx"
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

    ' Save the document to the final location
    '
    MsgBox "1"
    ActiveDocument.SaveAs2 _
        FileName:=strFullName, _
        FileFormat:=wdFormatXMLDocument

    ' Yield execution to the OS to make sure Word handles the file handle change
    DoEvents
    
    MsgBox "2"
    Dim strSavedFile As String
    strSavedFile = strFullName
    
    If Dir(strSavedFile) = "" Then
        MsgBox "Could not find saved file:" & vbCrLf & strSavedFile
        Exit Sub
    End If
    
'    ActiveDocument.SendMail

    ' create the email and attach the temporary file
    '
    MsgBox "3"
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

    ' Now save the document to the final location
    '
'    MsgBox "4"
'    ActiveDocument.SaveAs2 _
'        FileName:=strFullName, _
'        FileFormat:=wdFormatXMLDocument
    
    ' Yield execution to the OS to make sure Word handles the file handle change
'    DoEvents
    
'    Sleep 30000 ' Wait for 30 seconds

    MsgBox "5"
    ' Restore user's original settings
    Options.BackgroundSave = originalBackgroundSetting

    ' Clean up temporary file
'    Kill strTempFile

    '  Shutdown word without saving changes
    '
'    ActiveDocument.Close SaveChanges:=False
'    Application.Quit SaveChanges:=wdDoNotSaveChanges
    
End Sub
