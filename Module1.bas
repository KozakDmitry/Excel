Attribute VB_Name = "Module1"
    Function FindAndCopy(searchTerm As Variant, columnToSearch As Variant)
    Dim ws As Worksheet
    Dim found As Range
    Dim copyTo As Range
    Dim i As Long
    Dim choose As String
    Dim lastRow As Range
    

    Set copyTo = Selection
    
    For Each ws In ActiveWorkbook.Worksheets
            For i = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
                For j = 1 To 14
                    If ws.Cells(j, i).Value = columnToSearch Then
                        Set found = ws.Cells.Columns(i).Find(What:=searchTerm, LookIn:=xlValues, LookAt:=xlWhole)
                        If Not found Is Nothing Then
                           
                            Set lastRow = ActiveSheet.Range("A" & ActiveSheet.Rows.Count).End(xlUp).Offset(1)
                           
                            ws.Cells(j, i).EntireRow.Copy
                            
                            
                            ' choose = InputBox("Ввести это в документ или нет?")
                            ' If (choose = "нет") Then
                            lastRow.PasteSpecial xlPasteValues
                            lastRow.Resize(1, ws.Cells(j, i).EntireRow.Columns.Count).Font.Bold = True
                            ws.Cells(found.Row, i).Offset(0, 1).EntireRow.Copy
                            lastRow.Offset(1).PasteSpecial xlPasteValues
                            lastRow.Value = ws.Name
                            FindAndInsertAfter lastRow

                        End If
                        Exit For
                    End If
                Next j
            Next i
            
    Next ws
End Function

Sub SearchAndCopy()
    Dim searchValue As Variant
    Dim columnToSearch As Variant
    

 
    columnToSearch = InputBox("Введите название столбца для поиска:")
    searchValue = InputBox("Введите значение для поиска:")
 
    If searchValue = "" Or columnToSearch = "" Then
    MsgBox "Неправильный ввод"
    Else
    FindAndCopy searchValue, columnToSearch
    End If
    
    
End Sub

Function FindAndInsertAfter(lastRow)
    
    Dim obj As OLEObject
    Dim targetObject As OLEObject
    Dim wrdRange As Object
    Dim wdApp As Object
    Dim searchText As String
    Dim foundRange As Range
    Dim cell As Range
    Dim lastColumn As Long
    Dim startingCell As Range
    Dim rowRange As Range
    
    For Each ws In ActiveWorkbook.Worksheets
        For Each obj In ws.OLEObjects
            If Not obj.Verb = xlVerbChart Then
                Set targetObject = obj
                Exit For
                Exit For
            End If
        Next obj
    Next ws
    targetObject.Name = "WordDoc"
    
    Set startingCell = lastRow.Cells(1, 1)
    Debug.Print (startingCell)
    
    Set wdApp = targetObject.Object.Application
    Set wdDoc = targetObject.Object
   
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    searchText = "Значения для подстановки"
    
    
    Set rowRange = Range(startingCell.Offset(0, 1), startingCell.End(xlToRight))
    
    For Each ws In ActiveWorkbook.Worksheets
        Set foundRange = ws.UsedRange.Find(What:=searchText, LookIn:=xlValues, LookAt:=xlWhole)
        If Not foundRange Is Nothing Then
            Exit For
        End If
    Next ws
    Set secondRange = foundRange.Offset(1)
    
    For Each cell1 In rowRange

        For Each cell2 In Range(secondRange.Offset(0, 1), secondRange.End(xlToRight))
        
            If cell1.Value = cell2.Value Then
                wdApp.Visible = True
                wdDoc.Activate

                wdDoc.Content.Find.Execute findText:=cell2.Offset(1, 0).Value, ReplaceWith:=cell1.Offset(-1, 0).Value
                
                wdDoc.Inactivate
                

            End If
        Next cell2
    Next cell1
        
End Function



