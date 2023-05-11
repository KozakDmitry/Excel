Attribute VB_Name = "Module1"
    Function FindAndCopy(searchTerm As Variant, columnToSearch As Variant)
    Dim ws As Worksheet
    Dim found As Range
    Dim copyTo As Range
    Dim i As Long
    Dim choose As String
    Set copyTo = Selection
    
    For Each ws In ActiveWorkbook.Worksheets
            For i = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
                For j = 1 To 14
                    If ws.Cells(j, i).Value = columnToSearch Then
                        Set found = ws.Cells.Columns(i).Find(What:=searchTerm, LookIn:=xlValues, LookAt:=xlWhole)
                        If Not found Is Nothing Then
                            Dim lastRow As Range
                            Set lastRow = ActiveSheet.Range("A" & ActiveSheet.Rows.Count).End(xlUp).Offset(1)
                            ws.Cells(j, i).EntireRow.Copy
                            
                            
                            ' choose = InputBox("Ввести это в документ или нет?")
                            ' If (choose = "нет") Then
                            lastRow.PasteSpecial xlPasteValues
                            lastRow.Resize(1, ws.Cells(j, i).EntireRow.Columns.Count).Font.Bold = True
                            ws.Cells(found.Row, i).Offset(0, 1).EntireRow.Copy
                            lastRow.Offset(1).PasteSpecial xlPasteValues
                            lastRow.Value = lastRow.Value & " " & ws.Name
                            FindAndInsertAfter (lastRow)
                            
                            'ElseIf (choose = "да") Then
                            'Dim wdDoc As Object
                            'Set wdDoc = GetObject(, "Word.Application").ActiveDocument
                            'Dim rng As Object
                            'Set rng = wdDoc.Range(Start:=0, End:=0)
                                                        
                            
                            'End If
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

Function FindAndInsertAfter(lastRow As Variant)
    
    Dim obj As OLEObject
    Dim targetObject As OLEObject
    Dim wrdRange As Object
    Dim wdApp As Object
    Dim searchText As String
    Dim foundRange As Range
    Dim cell As Range
    Dim lastColumn As Long
    
    
    For Each ws In ActiveWorkbook.Worksheets
        For Each obj In ws.OLEObjects
            If Not obj.Verb = xlVerbChart Then
                Set targetObject = obj
                Exit For
            End If
        Next obj
    Next ws
    targetObject.Name = "WordDoc"
    Debug.Print (targetObject.Name)
    
    Set wdApp = targetObject.Object.Application
    Set wdDoc = targetObject.Object
   
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    searchText = "Значения для подстановки"
    
    For Each ws In ActiveWorkbook.Worksheets
        Set foundRange = ws.UsedRange.Find(What:=searchText, LookIn:=xlValues, LookAt:=xlWhole)
        If Not foundRange Is Nothing Then
            Exit For
        End If
    Next ws
    Set secondRange = foundRange.Offset(1)

    

    For Each cell1 In lastRow

        For Each cell2 In foundRange
     
            If cell1.Value = cell2.Value Then
                wdApp.Visible = True
                wdDoc.Activate

                wdDoc.Content.Find.Execute findText:=cell2.Offset(1, 0).Value, ReplaceWith:=cell1.Offset(-1, 0).Value
                
                wdDoc.Inactivate
                

            End If
        Next cell2
    Next cell1
        
    End If
End Function

Sub test(lastRow As Variant)
    
    
    Sub FindMatchingCells()
    Dim targetRow1 As Range
    Dim targetRow2 As Range
    Dim cell1 As Range
    Dim cell2 As Range
    
    

End Sub


