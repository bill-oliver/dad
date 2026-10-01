Sub OpenDraft()

    Dim fd As FileDialog
    Dim strSelectedFile As String

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "Select a saved draft"
    fd.InitialFileName = "C:\Users\boliv\Documents\emails\"
    fd.Filters.Clear
    fd.Filters.Add "Draft emails", "*.docm"

    If fd.Show = -1 Then
        strSelectedFile = fd.SelectedItems(1)
        Documents.Open FileName:=strSelectedFile
    Else
        MsgBox "No draft selected.", vbInformation
    End If

End Sub