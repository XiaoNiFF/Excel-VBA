Attribute VB_Name = "多条件查询"
Sub pipei()
'多条件查询

    Dim arr, brr
    Dim d As Object
    Dim i%, j%, k%
    Dim m%, n%
    Dim s$
    arr = Range("A1").CurrentRegion
    '被查询区域装入数组arr，数组代替单元格区域，效率更高
    Set d = CreateObject("scripting.dictionary")
    For i = 2 To UBound(arr, 1)
        For j = 2 To UBound(arr, 2)
            '将查询区域即数组arr装入字典的键值对，多条件，用行和列定位
            s = arr(i, 1) & "+" & arr(1, j)
            d(s) = arr(i, j)
        Next j
    Next i
    brr = Cells(1, Columns.Count).End(xlToLeft).CurrentRegion
    '查询区域装入数组brr，后面配合字典查询
    For k = 2 To UBound(brr, 1)
        s = brr(k, 1) & "+" & brr(k, 2)
        If d.exists(s) Then
            brr(k, 3) = d(s)
            m = m + 1
        Else
            brr(k, 3) = "N/A"
            n = n + 1
        End If
    Next
    With Cells(1, Columns.Count).End(xlToLeft).CurrentRegion '设置区域基本格式
        .NumberFormat = "@"
        .Font.Name = "宋体"
        .Font.Size = 16
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Value = brr
    End With
    Set d = Nothing
    MsgBox "查找到" & m & "个成绩" & vbNewLine & "未找到" & n & "个成绩"
End Sub
