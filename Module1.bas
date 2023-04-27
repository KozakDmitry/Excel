Attribute VB_Name = "Module1"
Function FindAndCopy(searchTerm As Variant, columnToSearch As Variant)
    Dim ws As Worksheet
    Dim found As Range
    Dim copyTo As Range
    Dim i As Long
    Dim choose As String
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
                            
                            
                            ' choose = InputBox("������ ��� � �������� ��� ���?")
                            ' If (choose = "���") Then
                            lastRow.PasteSpecial xlPasteValues
                            lastRow.Resize(1, ws.Cells(j, i).EntireRow.Columns.Count).Font.Bold = True
                            ws.Cells(found.Row, i).Offset(0, 1).EntireRow.Copy
                            lastRow.Offset(1).PasteSpecial xlPasteValues
                            lastRow.value = lastRow.value & " " & ws.Name
                            
                            'ElseIf (choose = "��") Then
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
    If copyTo.Address = Selection.Address Then
        
    End If
End Function

Sub SearchAndCopy()
    Dim searchValue As Variant
    Dim columnToSearch As Variant
    

    searchValue = InputBox("������� �������� ��� ������:")
    columnToSearch = InputBox("������� �������� ������� ��� ������:")
    
    If searchValue = "" Or columnToSearch = "" Then
    MsgBox "������������ ����"
    Else
    FindAndCopy searchValue, columnToSearch
    End If
    
    
End Sub

Function FindAndInsertAfter()
    
    Dim InsertAfter As String
    Dim wrdApp As Object
    Dim wrdDoc As Object
    
    Dim wrdRange As Object
    
    Dim findText As String
    Dim insertText As String
    Dim found As Boolean
    
    ' ���������� ����� ��� ������ � �������
    findText = ""
    insertText = "����������� �����"
    
    ' ������� ������ ���������� Word
    Set wrdApp = CreateObject("Word.Application")
    
    ' ��������� ��������
    Set wrdDoc = wrdApp.Documents.Open("����_�_���������")
    
    ' ������������� �������� ������ ���� ��������
    Set wrdRange = wrdDoc.Content
    
    ' ���� �����
    With wrdRange.Find
        .Text = findText
        .Execute
        If .found Then
            ' ����� �����, ��������� ����� ����� ����
            wrdRange.InsertAfter insertText
            found = True
        End If
    End With
    
    ' ���� ����� �� ������, ������� ��������� �� ������
    If Not found Then
        MsgBox "����� �� ������"
    End If
    
    ' ��������� �������� � ���������� Word
    wrdDoc.Close
    wrdApp.Quit
    
    ' ����������� ������, ������� ��������� Word
    Set wrdRange = Nothing
    Set wrdDoc = Nothing
    Set wrdApp = Nothing

End Function

