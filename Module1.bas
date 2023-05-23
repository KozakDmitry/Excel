Attribute VB_Name = "Module1"
    Public originalDoc As Object
    Public ЦелаяЧасть, Текст, коп, цнт As String
    Public единМ, единЖ, десят1, десятки, сотни, тысячи, _
    миллионы, миллиарды, рубли, копейки, доллары, центы, _
    евры, №склона, Шаг As Variant
    Public группаЕСТЬ As Boolean
    Public localSelect As Range
    Public Рр(1 To 14), точкаРазд As Integer
    
    
    Function FindAndCopy(searchTerm As Variant, columnToSearch As Variant)
    Dim ws As Worksheet
    Dim found As Range
    Dim copyTo As Range
    Dim I As Long
    Dim choose As String
    Dim lastRow As Range
    

    Set copyTo = selection
    
    For Each ws In ActiveWorkbook.Worksheets
            For I = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
                For j = 1 To 14
                    If ws.Cells(j, I).Value = columnToSearch Then
                        Set found = ws.Cells.Columns(I).Find(What:=searchTerm, LookIn:=xlValues, LookAt:=xlWhole)
                        If Not found Is Nothing Then
                           
                            Set lastRow = ActiveSheet.Range("A" & ActiveSheet.Rows.Count).End(xlUp).Offset(1)
                            
                            ws.Cells(j, I).EntireRow.Copy
                            
                            

                            
                            lastRow.PasteSpecial xlPasteValues
                            lastRow.Resize(1, ws.Cells(j, I).EntireRow.Columns.Count).Font.Bold = True
                            ws.Cells(found.Row, I).Offset(0, 1).EntireRow.Copy
                            lastRow.Offset(1).PasteSpecial xlPasteValues
                            lastRow.Value = ws.Name
                            
                            

                        End If
                        Exit For
                    End If
                Next j
            Next I
            
    Next ws
    Set localSelect = lastRow
    Debug.Print (localSelect.Address)
    
    choose = MsgBox("Вопрос: Ввести это в документ?", vbYesNo)
    If choose = vbYes Then
        ВставитьПоследнюю lastRow
    End If
    
End Function

Sub НайтиВставить()
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

Function ВставитьПоследнюю(functionSelect)
    
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
    

    
    If functionSelect Is Nothing Then
        MsgBox ("Нечего вставлять")
        Exit Function
    End If
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
    
    Set startingCell = localSelect.Cells(1, 1)
    
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
                    If LCase(cell2.Offset(2, 0).Value) = "да" Then
                    textNum = БелРуб(cell1.Offset(-1, 0).Value, True)
                    wdDoc.Content.Find.Execute findText:=cell2.Offset(1, 0).Value, ReplaceWith:=cell1.Offset(-1, 0).Value & " BYN (" & textNum & ")", Replace:=2, Wrap:=1
                    Else
                    wdDoc.Content.Find.Execute findText:=cell2.Offset(1, 0).Value, ReplaceWith:=cell1.Offset(-1, 0).Value, Replace:=2, Wrap:=1
                    End If
                Else
                    wdDoc.Content.Find.Execute findText:=cell2.Offset(1, 0).Value, ReplaceWith:=cell1.Offset(-1, 0).Value, Replace:=2, Wrap:=1
                End If
                'Excel.Application.Activate
 
            End If
        Next cell2
    Next cell1
        'wdApp.Quit
        
        
End Function




Function БелРуб(Сумма, Optional сКопейками As Boolean)
' Сумма белорусских рублей прописью _
  в диапазоне от 0 до 999 млрд. с копейками
' создана 21.02.02 (Николай Домарёнок)
' если параметр сКопейками = ЛОЖЬ, _
  то текст "00 копеек" не добавляется к результату.

СтройМат    ' Объявление массива исходных текстовых значений
рубли = Array("белорусский рубль", "белорусских рубля", "белорусских рублей")
копейки = Array("копейка", "копейки", "копеек")

Текст = ""  ' Очистка строки преобразуемого текста

' Определение положения точки разделения целой и дробной частей:

точкаРазд = InStr(1, Сумма, ".", 1) + InStr(1, Сумма, ",", 1) + InStr(1, Сумма, "=", 1)
If точкаРазд = 0 Then
    коп = "00"
    точкаРазд = Len(Сумма) + 1
Else
    коп = Left(Mid(Сумма, точкаРазд + 1, 2) & "00", 2)
End If

' Формирование целой части 12-разрядной суммы с лидирующими нулями
' и проверка переполнения диапазона преобразования числа в текст:

ЦелаяЧасть = Right("0000000000000" & Mid(Сумма, 1, точкаРазд - 1), 13)
If Val(ЦелаяЧасть) > 999999999999# Then
    БелРуб = "Cумма выходит за границы допустимого диапазона (0-999999999999.99)."
    Exit Function
End If
ЦелаяЧасть = Right("000000000000" & Mid(Сумма, 1, точкаРазд - 1), 12)

' Присвоение переменным Рр(1)...Рр(14) значений соответствующих разрядов
' преобразуемой суммы:
For I = 1 To 12                     ' для рублей
  Рр(I) = Val(Mid(ЦелаяЧасть, I, 1))
Next I
For I = 13 To 14                    ' для копеек
  Рр(I) = Val(Mid(коп, I - 12, 1))
Next I

' Формирование преобразуемой текстовой строки по триадам разрядов от старших к младшим

СтаршиеРазряды
Группа Шаг:=9
If группаЕСТЬ = True Then
    Склонять Шаг:=9
    Текст = Текст & рубли(№склона) & " "
Else
    №склона = 2
    Текст = Текст & рубли(№склона) & " "
End If

' Выделение прописной буквой начала преобразуемой строки

Текст = UCase(Mid(Trim(Текст), 1, 1)) & Mid(Trim(Текст), 2)

' Проверка условия об указании копеек в текстовой строке

If Not сКопейками Or коп = "00" Then
    Текст = Текст & " "
Else
    Склонять Шаг:=11
    Текст = Текст & " " & коп & " " & копейки(№склона)
End If

' Окончательная запись текстовой строки преобразуемой суммы

    БелРуб = Текст
End Function


Function СтройМат()
' Процедура группировки исходных элементов формируемой текстовой строки
'
единМ = Array("", "один ", "два ", "тpи ", "четыpе ", "пять ", "шесть ", "семь ", "восемь ", "девять ")
единЖ = Array("", "одна ", "две ", "тpи ", "четыpе ", "пять ", "шесть ", "семь ", "восемь ", "девять ")
десят1 = Array("десять ", "одиннадцать ", "двенадцать ", "тринадцать ", "четырнадцать ", "пятнадцать ", "шестнадцать ", "семнадцать ", "восемнадцать ", "девятнадцать ")
десятки = Array("", "десять ", "двадцать ", "тpидцать ", "сорок ", "пятьдесят ", "шестьдесят ", "семьдесят ", "восемьдесят ", "девяносто ")
сотни = Array("", "сто ", "двести ", "тpиста ", "четыреста ", "пятьсот ", "шестьсот ", "семьсот ", "восемьсот ", "девятьсот ")
тысячи = Array("тысяча", "тысячи", "тысяч")
миллионы = Array("миллион", "миллиона", "миллионов")
миллиарды = Array("миллиард", "миллиарда", "миллиардов")
End Function

Function СтаршиеРазряды()
' Формирование текста старших 9 разрядов суммы

Группа Шаг:=0
If группаЕСТЬ = True Then
    Склонять Шаг:=0
    Текст = Текст & миллиарды(№склона) + " "
End If
Группа Шаг:=3
If группаЕСТЬ = True Then
    Склонять Шаг:=3
    Текст = Текст & миллионы(№склона) + " "
End If
Группа Шаг:=6
If группаЕСТЬ = True Then
    Склонять Шаг:=6
    Текст = Текст & тысячи(№склона) + " "
End If
End Function

Function Группа(Шаг)
' Процедура преобразования в текст группы (триады) чисел

If Val(Mid(ЦелаяЧасть, 1 + Шаг, 3)) <> 0 Then
    Текст = Текст & сотни(Рр(1 + Шаг))      ' текстовая запись сотен
    If Рр(2 + Шаг) = 1 Then
        Текст = Текст & десят1(Рр(3 + Шаг)) ' текстовая запись чисел от 11 до 19
    Else
        Текст = Текст & десятки(Рр(2 + Шаг)) & IIf(Шаг = 6, _
        единЖ(Рр(3 + Шаг)), единМ(Рр(3 + Шаг)))
        ' тестовая запись десятков от 10 до 90 и единиц мужского рода _
        с определением разрядов тысяч в женском роде
    End If
    группаЕСТЬ = True
Else
    группаЕСТЬ = False
End If
End Function
    
Function Склонять(Шаг)
'Процедура склонения по падежам в единственном и множественном числе
'единиц измерения по группам (триадам)

If Рр(2 + Шаг) = 1 Then         ' проверка на числа от 10 до 19
    №склона = 2                 ' миллиардов, миллионов, тысяч, рублей
Else
    Select Case Рр(3 + Шаг)     ' проверка на числа от 1 до 9
        Case 1
            №склона = 0         ' (1) миллиард, миллион, тысяча, рубль
        Case 2 To 4
            №склона = 1         ' (2,3,4) миллиарда, миллиона, тысячи, рубля
        Case Else
            №склона = 2         ' (0,5..9) миллиардов, миллионов, тысяч, рублей
    End Select
End If
End Function

