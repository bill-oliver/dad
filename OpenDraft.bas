Sub OpenDraft()

    '  OpenDraft.bas:  macro to open a previously saved draft.  
    '  -------------
    '
    Dim fd As FileDialog
    Dim strSelectedFile As String
    Dim strFolderPath As String

    strFolderPath = "C:\Users\boliv\Documents\emails\"  ' Test location

    ' Display the file dialog to select a draft
    '
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "Select a saved draft email"
    fd.InitialFileName = strFolderPath
    fd.Filters.Clear
    fd.Filters.Add "Draft emails", "*.docm"

    If fd.Show = -1 Then
        strSelectedFile = fd.SelectedItems(1)
        Documents.Open FileName:=strSelectedFile
    Else
        MsgBox "No file selected.", vbInformation
    End If

End Sub