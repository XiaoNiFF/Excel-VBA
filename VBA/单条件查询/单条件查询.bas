Attribute VB_Name = "单条件查询"
Sub pipei()
'字典法替代vlookup但条件查询

    Dim arr, brr
    Dim d As Object
    Dim i&, j&
    arr = Range("A1:B" & Cells(1048576, 1).End(xlUp).Row)
    '被查询区域装入数组arr，数组代替单元格区域，效率更高
    Set d = CreateObject("scripting.dictionary")
    For i = 2 To UBound(arr, 1)
    '将查询区域即数组arr装入字典的键值对
        d(arr(i, 1)) = arr(i, 2)
    Next
    brr = Range("E1:F" & Cells(1048576, 5).End(xlUp).Row)
    '查询区域装入数组brr，后面配合字典查询
    For j = 2 To UBound(brr, 1)
        If d.exists(brr(j, 1)) Then
            brr(j, 2) = d(brr(j, 1))
        Else
            brr(j, 2) = "N/A"
        End If
    Next
    With Range("E1:F" & Cells(1048576, 5).End(xlUp).Row)
        .NumberFormat = "@"
        .Font.Name = "宋体"
        .Font.Size = 16
        .Value = brr
    End With
    Set d = Nothing
End Sub
