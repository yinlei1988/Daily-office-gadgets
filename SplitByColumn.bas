Attribute VB_Name = "ģ��1"
Sub SplitByColumn()
    ' ��������
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
    
    ' ����һ���ֵ�������ڴ洢Ҫ��ֵ��е�Ψһֵ
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' ��ȡҪ����е����һ�е��кţ��޸�Ϊ���ݵ����н��в��
    lastRow = Cells(Rows.Count, 7).End(xlUp).Row
    ' ��ȡ��һ�е����һ�е��к�
    lastCol = Cells(1, Columns.Count).End(xlToLeft).Column
    
    ' �������е�Ψһֵ��ӵ��ֵ��У��ӵڶ��п�ʼ�������һ���Ǳ�ͷ��
    For i = 2 To lastRow
        ' ����ֵ����Ƿ��Ѿ����ڸ�ֵ���������������ӵ��ֵ���
        If Not dict.Exists(Cells(i, 7).Value) Then
            dict.Add Cells(i, 7).Value, ""
        End If
    Next i
    
    ' Ϊÿ��Ψһֵ����һ���µĹ���������������
    For Each Key In dict.Keys
        ' ����һ���µĹ�����
        Set newWb = Workbooks.Add
        ' ��ȡ�¹������еĵ�һ��������
        Set newSht = newWb.Sheets(1)
        ' ���¹��������������Ϊ��Ψһֵ
        newSht.Name = Key
        ' ���¹���������������Ϊ��Ψһֵ
        newWb.SaveAs Filename:=ThisWorkbook.Path & "\" & Key & ".xlsx"
        
        ' ���Ʊ�ͷ
        For j = 1 To lastCol
            newSht.Cells(1, j).Value = ThisWorkbook.Sheets(1).Cells(1, j).Value
        Next j
        
        ' ���ƶ�Ӧ��������
        i = 2
        For Each sht In ThisWorkbook.Sheets
            ' ��ȡ��ǰ������Ҫ����е����һ�е��к�
            lastRow = sht.Cells(Rows.Count, 7).End(xlUp).Row
            For i = 2 To lastRow
                ' �����ǰ�е�ֵ����Ψһֵ
                If sht.Cells(i, 7).Value = Key Then
                    ' ���������ݸ��Ƶ��¹�������
                    newSht.Cells(newSht.Cells(Rows.Count, 1).End(xlUp).Row + 1, 1).Resize(1, lastCol).Value = sht.Cells(i, 1).Resize(1, lastCol).Value
                End If
            Next i
        Next sht
    Next Key
End Sub
