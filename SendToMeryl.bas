Sub SendToMeryl()
    Dim draftDocument As Document
    Dim strFullName As String
    Dim strFileName As String
    Dim strTo As String
    Dim strCC As String
    Dim olApp As Object
    Dim olMail As Object

    On Error GoTo HandleError

    ' Set the email addresses for email recipients.
    strTo = "meryl.oliver@gmail.com"
    strCC = "bill@oliverassociates.ca; johnnieo@shaw.ca"

    ' Keep a reference to this draft; Outlook taking focus should not change which document is closed.
    Set draftDocument = ActiveDocument
    If draftDocument.Path = "" Then
        MsgBox "Save this draft with the template's new-draft button before preparing the email.", _
            vbExclamation, "Draft Not Saved"
        Exit Sub
    End If

    ' Include the latest edits in the attachment before creating the email.
    draftDocument.Save
    strFullName = draftDocument.FullName
    strFileName = draftDocument.Name
    If InStrRev(strFileName, ".") > 0 Then
        strFileName = Left$(strFileName, InStrRev(strFileName, ".") - 1)
    End If

    Set olApp = CreateObject("Outlook.Application")
    Set olMail = olApp.CreateItem(0)

    ' Display the message for review; Outlook remains open for the user to send it.
    With olMail
        .To = strTo
        .CC = strCC
        .Subject = "Draft for Review: " & strFileName
        .Body = "Hi Meryl," & vbCrLf & vbCrLf & _
                "Please review the attached draft and return your edits." & _
                vbCrLf & vbCrLf & _
                "Thanks," & vbCrLf & vbCrLf & _
                "Dad."
        .Attachments.Add strFullName
        .Display
    End With

    ' Close only this draft when other Word documents are open; quit Word if it is the last one.
    If Application.Documents.Count = 1 Then
        Application.Quit SaveChanges:=wdPromptToSaveChanges
    Else
        draftDocument.Close SaveChanges:=wdDoNotSaveChanges
    End If
    Exit Sub

HandleError:
    MsgBox "Could not prepare the review email: " & Err.Description, _
        vbExclamation, "Send to Meryl"
End Sub
