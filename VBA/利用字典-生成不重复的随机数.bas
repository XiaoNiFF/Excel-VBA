Sub 随机数()

    Dim m As Integer
    Dim d As Object
    Dim s As String
    Dim arr

    Set d = CreateObject("scripting.dictionary")

    Max = 1 '随机数下限
    Min = 50    '随机数上限
    x = 10  '随机数个数

    ReDim arr(1 To x, 1 To 1)

        For m = 1 To x

            Do
                s = Int((Max - Min + 1) * Rnd) + Min
                's = Int((Max - Min + 1) * Rnd + Min) 帮助文件中的描述
            Loop While d.exists(s)

            d(s) = ""
            arr(m, 1) = s

        Next

    Range("A1", "A" & x) = arr
    Set d = Nothing
End Sub
