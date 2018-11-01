Attribute VB_Name = "剪贴板函数"
Sub test()

    c2c (Range("A1").Value)
    Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).PasteSpecial Paste:=xlPasteAll

End Sub

Public Function c2c(strtext As String) '自定义剪贴板函数
    With CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}") 'DataObject类ID
        .settext strtext
        .putinclipboard
    End With
End Function
