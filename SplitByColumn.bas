Attribute VB_Name = "模块1"
Sub SplitByColumn()
    ' 声明变量
    Dim lastRow As Long
    Dim lastCol As Long
    Dim i As Long
    Dim j As Long
    Dim dict As Object
    Dim sht As Worksheet
    Dim newSht As Worksheet
    Dim newWb As Workbook
    Dim filePath As String
    Dim fileName As String
    
    ' 创建一个字典对象，用于存储要拆分的列的唯一值
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' 获取要拆分列的最后一行的行号，修改为根据第七列进行拆分
    lastRow = Cells(Rows.Count, 7).End(xlUp).Row
    ' 获取第一行的最后一列的列号
    lastCol = Cells(1, Columns.Count).End(xlToLeft).Column
    
    ' 将第七列的唯一值添加到字典中，从第二行开始（假设第一行是表头）
    For i = 2 To lastRow
        ' 检查字典中是否已经存在该值，如果不存在则添加到字典中
        If Not dict.Exists(Cells(i, 7).Value) Then
            dict.Add Cells(i, 7).Value, ""
        End If
    Next i
    
    ' 为每个唯一值创建一个新的工作簿并复制数据
    For Each Key In dict.Keys
        ' 创建一个新的工作簿
        Set newWb = Workbooks.Add
        ' 获取新工作簿中的第一个工作表
        Set newSht = newWb.Sheets(1)
        ' 将新工作表的名称设置为该唯一值
        newSht.Name = Key
        ' 将新工作簿的名称设置为该唯一值
        newWb.SaveAs Filename:=ThisWorkbook.Path & "\" & Key & ".xlsx"
        
        ' 复制表头
        For j = 1 To lastCol
            newSht.Cells(1, j).Value = ThisWorkbook.Sheets(1).Cells(1, j).Value
        Next j
        
        ' 复制对应的数据行
        i = 2
        For Each sht In ThisWorkbook.Sheets
            ' 获取当前工作表要拆分列的最后一行的行号
            lastRow = sht.Cells(Rows.Count, 7).End(xlUp).Row
            For i = 2 To lastRow
                ' 如果当前行的值等于唯一值
                If sht.Cells(i, 7).Value = Key Then
                    ' 将该行数据复制到新工作表中
                    newSht.Cells(newSht.Cells(Rows.Count, 1).End(xlUp).Row + 1, 1).Resize(1, lastCol).Value = sht.Cells(i, 1).Resize(1, lastCol).Value
                End If
            Next i
        Next sht
    Next Key
End Sub
