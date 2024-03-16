
# Automatically add date and time to filename when saving a Excel document


1.
![Snimka zaslona 2024-03-16 180851](https://github.com/bojkip/automatically-add-date-and-time-filename-excel/assets/91488932/6d96a26e-c812-4be9-a354-2aff25f378a4)

Alt + F8 or Alt + F11

2.
![Snimka zaslona 2024-03-16 180931](https://github.com/bojkip/automatically-add-date-and-time-filename-excel/assets/91488932/2b5e149d-24b5-425e-a0c1-2b4c08f3f9dc)
Copy⬇️ and paste⬆️ code (or import code file):

```

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



```

3.
![Snimka zaslona 2024-03-16 181147](https://github.com/bojkip/automatically-add-date-and-time-filename-excel/assets/91488932/175930bb-8bb8-48f6-9734-4cf6e4367ffa)

Alt + F8

4.
![Snimka zaslona 2024-03-16 181336](https://github.com/bojkip/automatically-add-date-and-time-filename-excel/assets/91488932/81ba785c-6f14-40a9-8eac-1972685933bd)








