Sub 将数据转为数组()
'
' 将数据转为数组
'

'
    
    Dim companyName As String
    
    Sheets("test").Select
    ' x的值为A列中最后一个非空单元格行号
    x = Range("A65536").End(3).Row
    
    
    ' 重新定义数组a，使之上限为x
    Dim rowArray
    ReDim totalArray(x)
    
    ' i从A1依次往下循环，直到最后一行
    For i = 1 To x
        companyName = Range("A" & i)
        
         If companyName Like "*海风*" Then
        ' 将A列单元格值依次存放于数组中
        rowArray = Range("$A$" & i & ":$J$" & i & "")
        totalArray(i - 1) = rowArray
        End If
    Next
    MsgBox UBound(totalArray)
    MsgBox totalArray(3)(1, 3)

End Sub
