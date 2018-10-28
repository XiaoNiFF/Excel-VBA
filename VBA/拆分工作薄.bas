Private Sub CommandButton1_Click()
'选择待拆分的工作薄，自动拆分到同目录下

    Dim myfile$
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim i&
    Application.ScreenUpdating = False
    With Application.FileDialog(msoFileDialogFilePicker)    '选择待拆分的工作薄
        .AllowMultiSelect = False   '不能多选
        .Show
        myfile = .SelectedItems(1)
    End With
    'MsgBox myfile
    Set wb = Workbooks.Open(myfile)
    'MsgBox wb.Path
    'MsgBox ActiveWorkbook.Name
    For Each ws In wb.Worksheets
        ws.Copy
        With ActiveWorkbook
            .SaveAs Filename:=wb.Path & "\" & ws.Name, FileFormat:=xlWorkbookDefault
            .Close True
        End With
        i = i + 1
    Next
    wb.Close False
    MsgBox "已经拆分了" & i & "个表格"
    Set wb = Nothing
    Application.ScreenUpdating = True
End Sub
