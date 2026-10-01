Sub SendToMeryl()
Dim olApp As Object, olMail As Object
ActiveDocument.Save
Set olApp = CreateObject("Outlook.Application")
Set olMail = olApp.CreateItem(0)
With olMail
 .To = "meryl.oliver@gmail.com"
 .CC = "bill@oliverassociates.ca"
 .Subject = "Draft for Review: " & ActiveDocument.Name
 .Body = "Hi Meryl," & vbCrLf & vbCrLf & "Please review the attached draft and return your edits."
 .Attachments.Add ActiveDocument.FullName
 .Display
End With
End Sub