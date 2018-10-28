Private Sub cb1_Click()
'Excel窗体组件，自定义上下限和随机数个数，生成随机数

Dim m As Integer
    Dim d As Object
    Dim s As Integer
    Dim str As String
    Dim arr

    Set d = CreateObject("scripting.dictionary")

    Max = Val(tb1.Text)  '随机数上限
    Min = Val(tb2.Text)    '随机数下限
    x = Val(tb3.Text)  '随机数个数
    
    'ReDim arr(1 To x, 1 To 1)

        For m = 1 To x

            Do
                s = Int((Max - Min + 1) * Rnd) + Min
                's = Int((Max - Min + 1) * Rnd + Min) 帮助文件中的描述
            Loop While d.exists(s)

            d(s) = ""
            str = s & " " & str

        Next

    tb4.Text = str
    Set d = Nothing

End Sub
