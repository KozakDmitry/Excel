Attribute VB_Name = "Module1"
    Public originalDoc As Object
    
    
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
                            
                            

                            
                            lastRow.PasteSpecial xlPasteValues
                            lastRow.Resize(1, ws.Cells(j, i).EntireRow.Columns.Count).Font.Bold = True
                            ws.Cells(found.Row, i).Offset(0, 1).EntireRow.Copy
                            lastRow.Offset(1).PasteSpecial xlPasteValues
                            lastRow.Value = ws.Name
                            
                            

                        End If
                        Exit For
                    End If
                Next j
            Next i
            
    Next ws
    choose = MsgBox("Вопрос: Ввести это в документ?", vbYesNo)
    If choose = vbYes Then
        FindAndInsertAfter lastRow
    End If
    
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
    Dim wdDuplicate As OLEObject
    Dim searchText As String
    Dim foundRange As Range
    Dim cell As Range
    Dim lastColumn As Long
    Dim startingCell As Range
    Dim rowRange As Range
    Dim findTex As String
    Dim Index As Integer
    Dim textNum As String
    
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
    
    Set wdDuplicate = targetObject.Duplicate
    Set wdApp = targetObject.Object.Application
    Set wdDoc = wdDuplicate.Object
    
    
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
    Set forceRange = Range(foundRange.Offset(0, 1), foundRange.End(xlToRight))
    Index = 0
    For Each cell1 In rowRange
        For Each cell2 In forceRange
            If cell2.Value = "Имя листа" Then
                wdApp.Visible = True
                wdDoc.Activate
                wdDoc.Content.Find.Execute findText:=cell2.Offset(1, 0).Value, ReplaceWith:=rowRange.Cells(0).Value, Replace:=2, Wrap:=1
            End If
            If cell1.Value = cell2.Value Then
                wdApp.Visible = True
                wdDoc.Activate
                If IsNumeric(cell1.Offset(-1, 0).Value) Then
                    'Index = Index + 1
                    'If Index > 2 Then
                     '   textNum = NumberToText(cell1.Offset(-1, 0).Value)
                      '  wdDoc.Content.Find.Execute findText:=cell2.Offset(1, 0).Value, ReplaceWith:=cell1.Offset(-1, 0).Value & " (" & textNum & ")", Replace:=2, Wrap:=1
                    'End If
                Else
                    wdDoc.Content.Find.Execute findText:=cell2.Offset(1, 0).Value, ReplaceWith:=cell1.Offset(-1, 0).Value, Replace:=2, Wrap:=1
                End If
                'Excel.Application.Activate
                

            End If
        Next cell2
    Next cell1
        'wdApp.Quit
        
        
End Function


Function NumberToText(ByVal number As Double) As String
    Dim integerPart As Long
    Dim decimalPart As Long
    Dim text As String
    Dim units() As Variant
    Dim tens() As Variant

    units = Array("один", "два", "три", "четыре", "пять", "шесть", "семь", "восемь", "девять")
    tens = Array("десять", "двадцать", "тридцать", "сорок", "пятьдесят", "шестьдесят", "семьдесят", "восемьдесят", "девяносто")

    integerPart = Int(number)
    decimalPart = Round((number - integerPart) * 100)

    If integerPart < 10 Then
        text = units(integerPart - 1)
    ElseIf integerPart < 20 Then
        text = units(integerPart - 1) & "надцать"
    ElseIf integerPart < 100 Then
        text = tens(Int(integerPart / 10) - 1)
        If integerPart Mod 10 > 0 Then
            text = text & " " & units(integerPart Mod 10 - 1)
        End If
    ElseIf integerPart < 1000 Then
        text = units(Int(integerPart / 100) - 1) & "сто"
        If integerPart Mod 100 >= 20 Then
            text = text & " " & tens(Int((integerPart Mod 100) / 10) - 1)
            If integerPart Mod 10 > 0 Then
                text = text & " " & units(integerPart Mod 10 - 1)
            End If
        ElseIf integerPart Mod 100 > 0 Then
            text = text & " " & units(integerPart Mod 100 - 1)
        End If
    ElseIf integerPart < 10000 Then
        text = units(Int(integerPart / 1000) - 1) & " тысяч"
        If integerPart Mod 1000 > 0 Then
            text = text & " " & NumberToText(integerPart Mod 1000)
        End If
    ElseIf integerPart < 1000000 Then
        text = NumberToText(Int(integerPart / 1000)) & " тысяч"
        If integerPart Mod 1000 > 0 Then
            text = text & " " & NumberToText(integerPart Mod 1000)
        End If
    Else
        text = "Неверное значение"
    End If

    If decimalPart > 0 Then
        text = text & " целых " & NumberToText(decimalPart) & " сотых"
    End If

    NumberToText = text
End Function
