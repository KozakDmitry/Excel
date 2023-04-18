Attribute VB_Name = "Module1"
Function FindAndCopy(searchTerm As Long, columnToSearch As Variant)
    Dim ws As Worksheet
    Dim found As Range
    Dim copyTo As Range
    Dim i As Long
    
    Set copyTo = Selection
    
    For Each ws In ThisWorkbook.Worksheets
            For i = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).column
                For j = 1 To 14
                    If ws.Cells(j, i).value = columnToSearch Then
                        Set found = ws.Cells.Columns(i).Find(What:=searchTerm, LookIn:=xlValues, LookAt:=xlWhole)
                        If Not found Is Nothing Then
                            Dim lastRow As Range
                            Set lastRow = ActiveSheet.Range("A" & ActiveSheet.Rows.Count).End(xlUp).Offset(1)
                            
                            ws.Cells(j, i).EntireRow.Copy
                            lastRow.PasteSpecial xlPasteValues
                            lastRow.Resize(1, ws.Cells(j, i).EntireRow.Columns.Count).Font.Bold = True
                            
                            ws.Cells(found.Row, i).Offset(0, 1).EntireRow.Copy
                            lastRow.Offset(1).PasteSpecial xlPasteValues
                            lastRow.value = lastRow.value & " " & ws.Name
                        End If
                        Exit For
                    End If
                Next j
            Next i
            
    Next ws
    
    If copyTo.Address = Selection.Address Then
        MsgBox "Ничего не найдено."
    End If
End Function

Sub SearchAndCopy()
    Dim searchValue As Long
    Dim columnToSearch As Variant
    
    ' Получаем значения от пользователя
    searchValue = InputBox("Введите значение для поиска:")
    columnToSearch = InputBox("Введите название столбца для поиска:")
    
    ' Вызываем функцию FindAndCopy с заданными параметрами
    FindAndCopy searchValue, columnToSearch
End Sub

