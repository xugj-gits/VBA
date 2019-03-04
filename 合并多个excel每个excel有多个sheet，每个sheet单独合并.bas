Sub 合并当前目录下所有工作簿的全部工作表()
    Dim MyPath, MyName, AWbName
    Dim Wb As Workbook, WbN As String
    Dim G As Long
    Dim Num As Long
    Dim BOX As String
    flag = 0
     
    Application.ScreenUpdating = False
    MyPath = ActiveWorkbook.Path
    MyName = Dir(MyPath & "\" & "*.xls")
    AWbName = ActiveWorkbook.Name
    Num = 0
  
     
    Do While MyName <> ""
        If MyName <> AWbName Then
            Set Wb = Workbooks.Open(MyPath & "\" & MyName)
            Num = Num + 1
             
             
                For G = 1 To Wb.Sheets.Count
                     
                    If flag = 0 Then
                    Sheets.Add after:=Sheets(Sheets.Count)
                     
                        With ActiveSheet
                               .Name = Wb.Sheets(G).Name
                           Wb.Sheets(G).UsedRange.Copy .Cells(.Range("A65536").End(xlUp).Row, 1)
                           .UsedRange.Rows.AutoFit
                           .UsedRange.Columns.AutoFit
                        End With
                    Else
                          With Workbooks(1).Worksheets(G + 3)
                          ' MsgBox .Name & "--" & Wb.Sheets(G).Name
                           If G = 2 Then
                            Wb.Sheets(G).Range("a2", Wb.Sheets(G).Cells.SpecialCells(xlCellTypeLastCell)).Copy .Cells(.Range("A65536").End(xlUp).Row + 2, 1)
                           Else
                            Wb.Sheets(G).Range("a2", Wb.Sheets(G).Cells.SpecialCells(xlCellTypeLastCell)).Copy .Cells(.Range("A65536").End(xlUp).Row + 1, 1)
                            End If
                             
                           .UsedRange.Rows.AutoFit
                           .UsedRange.Columns.AutoFit
                          End With
                           
                    End If
                Next
                'flag 为0时候为第一个打开的excel，此时产生列，sheet名
                 flag = 1
                WbN = WbN & Chr(13) & Wb.Name
                Wb.Close False
           ' End With
        End If
        MyName = Dir
    Loop
        Range("A1").Select
         
         
    Application.ScreenUpdating = True
    MsgBox "共合并了" & Num & "个工作薄下的全部工作表。如下：" & Chr(13) & WbN, vbInformation, "提示"
End Sub
