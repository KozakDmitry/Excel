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
Sub test()
    
    Dim columnToSearch As Variant
    Dim x As Variant
    

 
    columnToSearch = InputBox("Число:")
    Debug.Print (columnToSearch)
    columnToSearch = NumberToCurrency(columnToSearch)
    Debug.Print (columnToSearch)

End Sub

Function NumberToCurrency(ByVal number As Double) As String
    Dim integerPart As Long
    Dim decimalPart As String
    Dim curr As String
    Dim integerWords As String
    
    integerPart = Int(number)
    decimalPart = Right(Format(number, "0.00"), 2)
    
    If integerPart = 0 Then
        NumberToCurrency = "ноль рублей " & decimalPart & " копеек"
        Exit Function
    End If
    
    curr = "рубль"
    If integerPart Mod 10 > 1 And integerPart Mod 10 < 5 And (integerPart Mod 100 < 10 Or integerPart Mod 100 >= 20) Then
        curr = "рубля"
    ElseIf integerPart Mod 10 <> 1 Or integerPart Mod 100 = 11 Then
        curr = "рублей"
    End If
    
    integerWords = NumberToText(integerPart)
    
    NumberToCurrency = integerWords & " " & curr
    
    If decimalPart <> "00" Then
        NumberToCurrency = NumberToCurrency & " " & decimalPart & " копеек"
    End If
End Function

Function NumberToText(ByVal number As Double) As String
    Dim units() As String
    Dim tens() As String
    Dim text As String
    Dim wholePart As Long
    Dim decimalPart As Long
    Dim remainder As Long
    
    units = Split("ноль один два три четыре пять шесть семь восемь девять", " ")
    tens = Split("десять двадцать тридцать сорок пятьдесят шестьдесят семьдесят восемьдесят девяносто", " ")
    
    wholePart = Fix(number)
    decimalPart = Round((number - wholePart) * 100)
    
    text = ""
    remainder = wholePart Mod 100
    
    If remainder < 10 Then
        text = units(remainder)
    ElseIf remainder < 20 Then
        text = units(remainder Mod 10) & "надцать"
    Else
        text = tens(remainder \ 10 - 1)
        If remainder Mod 10 <> 0 Then
            text = text & " " & units(remainder Mod 10)
        End If
    End If
    
    If wholePart >= 100 And wholePart < 1000 Then
        text = units(wholePart \ 100) & " сто " & text
    ElseIf wholePart >= 1000 And wholePart < 1000000 Then
        text = NumberToText(wholePart \ 1000) & " тысяч " & text
    ElseIf wholePart >= 1000000 Then
        text = NumberToText(wholePart \ 1000000) & " миллионов " & text
    End If
    
    If decimalPart > 0 Then
        text = text & " " & decimalPart & " копеек"
    End If
    
    NumberToText = text
End Function


Function NumberToTextInGroup(ByVal number As Long) As String
    Dim units() As String
    Dim tens() As String
    Dim text As String
    Dim remainder As Long
    
    units = Split("одна две три четыре пять шесть семь восемь девять", " ")
    tens = Split("десять двадцать тридцать сорок пятьдесят шестьдесят семьдесят восемьдесят девяносто", " ")
    
    text = ""
    remainder = number Mod 100
    
    If remainder < 10 Then
        text = units(remainder)
    ElseIf remainder < 20 Then
        text = units(remainder Mod 10) & "надцать"
    Else
        text = tens(remainder \ 10 - 1)
        If remainder Mod 10 <> 0 Then
            text = text & " " & units(remainder Mod 10)
        End If
    End If
    
    If number >= 100 Then
        text = units(number \ 100 - 1) & "сот " & text
    End If
    
    NumberToTextInGroup = text
End Function
