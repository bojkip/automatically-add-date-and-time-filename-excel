
Sub CustomFileSave()
    On Error Resume Next
    
    Dim strName As String
    Dim dlgSave As FileDialog
    
  
    strName = ActiveWorkbook.BuiltinDocumentProperties("Title")
    
  
    strName = strName & ActiveWorkbook.Name & "_" & Format(Now, "dd-MM-yyyy_hh-mm") & "h"
    

    Set dlgSave = Application.FileDialog(msoFileDialogSaveAs)
    
    With dlgSave
        .InitialFileName = strName
        If .Show = -1 Then
           
            ActiveWorkbook.SaveAs .SelectedItems(1)
        Else
            
            MsgBox "Document saving has been cancelled"
        End If
    End With
End Sub


