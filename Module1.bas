Attribute VB_Name = "Module1"
    Public originalDoc As Object
    Public ����������, �����, ���, ��� As String
    Public �����, �����, �����1, �������, �����, ������, _
    ��������, ���������, �����, �������, �������, �����, _
    ����, �������, ��� As Variant
    Public ���������� As Boolean
    Public localSelect As Range
    Public ��(1 To 14), ��������� As Integer
    
    
    Function FindAndCopy(searchTerm As Variant, columnToSearch As Variant)
    Dim ws As Worksheet
    Dim found As Range
    Dim copyTo As Range
    Dim i As Long
    Dim choose As String
    Dim lastRow As Range
    Dim NumberDocToSearch As String
    
    
    '������������ ��� ����������� ������ ����� ��������
    NumberDocToSearch = "�"
    Set copyTo = selection
    
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
                            Set searchNumberRange = ws.UsedRange.Rows("2:2")
                            Set foundCell = searchNumberRange.Find(What:=NumberDocToSearch, LookIn:=xlValues, LookAt:=xlPart)
                            lastRow.Value = foundCell.Value
                        End If
                        Exit For
                    End If
                Next j
            Next i
            
    Next ws
    Set localSelect = lastRow
    'Debug.Print (localSelect.Address)
    
    choose = MsgBox("������: ������ ��� � ��������?", vbYesNo)
    If choose = vbYes Then
        ����������������� lastRow
    End If
    
End Function

Sub �������������()
    Dim searchValue As Variant
    Dim columnToSearch As Variant
    

 
    columnToSearch = "���.�"
    searchValue = InputBox("������� �������� ��� ������:")
 
    If searchValue = "" Then
    MsgBox "������������ ����"
    Else
    FindAndCopy searchValue, columnToSearch
    End If
    
    
End Sub
Function ItemExists(col As Collection, item As Variant) As Boolean
    Dim i As Integer
    ItemExists = False
    For i = 1 To col.Count
        If col(i) = item Then
            ItemExists = True
            Exit Function
        End If
    Next i
End Function
Function �����������������(functionSelect)
    
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
    
    
    Dim dict As Object
    Dim dictForNumbers As New Collection
    dictForNumbers.Add "�������"
    dictForNumbers.Add "��������� �����"
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    dict.Add "��� �����", "�������"
    dict.Add "���.�", "�����������"
    dict.Add "��� ���-��", "�������"
    dict.Add "������������", "���������"
    dict.Add "���� ����� (��� �������)", "�������"
    dict.Add "��������� �����", "���������������"
    dict.Add "�������", "����������"
    
    
    
    If functionSelect Is Nothing Then
        MsgBox ("������ ���������")
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

    
    
    Set rowRange = Range(startingCell.Offset(0, 1), startingCell.End(xlToRight))
    
     
    ' For Each ws In ActiveWorkbook.Worksheets
     '   Set foundRange = ws.UsedRange.Find(What:=searchText, LookIn:=xlValues, LookAt:=xlWhole)
      '  If Not foundRange Is Nothing Then
       '     Exit For
    '    End If
   '  Next ws
   ' Set secondRange = foundRange.Offset(1)
   ' Set forceRange = Range(foundRange.Offset(0, 1), foundRange.End(xlToRight))
    Index = 0
    
  
    For Each cell1 In rowRange
            If cell1.Address = rowRange.Cells(1).Address Then
                wdApp.Visible = True
                wdDoc.Activate
                wdDoc.Content.Find.Execute findText:=dict("��� �����"), ReplaceWith:=rowRange.Cells(0).Value, Replace:=2, Wrap:=1
          
            End If
            If Not IsEmpty(cell1.Value) And dict.Exists(cell1.Value) Then
                wdApp.Visible = True
                wdDoc.Activate
                If IsNumeric(cell1.Offset(-1, 0).Value) Then
                    If ItemExists(dictForNumbers, cell1.Value) Then
                        textNum = ������(cell1.Offset(-1, 0).Value, True)
                        wdDoc.Content.Find.Execute findText:=dict(cell1.Value), ReplaceWith:=cell1.Offset(-1, 0).Value & " BYN (" & textNum & ")", Replace:=2, Wrap:=1
                        Else
                        wdDoc.Content.Find.Execute findText:=dict(cell1.Value), ReplaceWith:=cell1.Offset(-1, 0).Value, Replace:=2, Wrap:=1
                        End If
                Else
                    wdDoc.Content.Find.Execute findText:=dict(cell1.Value), ReplaceWith:=cell1.Offset(-1, 0).Value, Replace:=2, Wrap:=1
                End If
                'Excel.Application.Activate
 
            End If
    Next cell1
        'wdApp.Quit
        wdDuplicate.Delete
        
End Function




Function ������(�����, Optional ���������� As Boolean)
' ����� ����������� ������ �������� _
  � ��������� �� 0 �� 999 ����. � ���������
' ������� 21.02.02 (������� ��������)
' ���� �������� ���������� = ����, _
  �� ����� "00 ������" �� ����������� � ����������.

��������    ' ���������� ������� �������� ��������� ��������
����� = Array("����������� �����", "����������� �����", "����������� ������")
������� = Array("�������", "�������", "������")

����� = ""  ' ������� ������ �������������� ������

' ����������� ��������� ����� ���������� ����� � ������� ������:

��������� = InStr(1, �����, ".", 1) + InStr(1, �����, ",", 1) + InStr(1, �����, "=", 1)
If ��������� = 0 Then
    ��� = "00"
    ��������� = Len(�����) + 1
Else
    ��� = Left(Mid(�����, ��������� + 1, 2) & "00", 2)
End If

' ������������ ����� ����� 12-��������� ����� � ����������� ������
' � �������� ������������ ��������� �������������� ����� � �����:

���������� = Right("0000000000000" & Mid(�����, 1, ��������� - 1), 13)
If Val(����������) > 999999999999# Then
    ������ = "C���� ������� �� ������� ����������� ��������� (0-999999999999.99)."
    Exit Function
End If
���������� = Right("000000000000" & Mid(�����, 1, ��������� - 1), 12)

' ���������� ���������� ��(1)...��(14) �������� ��������������� ��������
' ������������� �����:
For i = 1 To 12                     ' ��� ������
  ��(i) = Val(Mid(����������, i, 1))
Next i
For i = 13 To 14                    ' ��� ������
  ��(i) = Val(Mid(���, i - 12, 1))
Next i

' ������������ ������������� ��������� ������ �� ������� �������� �� ������� � �������

��������������
������ ���:=9
If ���������� = True Then
    �������� ���:=9
    ����� = ����� & �����(�������) & " "
Else
    ������� = 2
    ����� = ����� & �����(�������) & " "
End If

' ��������� ��������� ������ ������ ������������� ������

����� = UCase(Mid(Trim(�����), 1, 1)) & Mid(Trim(�����), 2)

' �������� ������� �� �������� ������ � ��������� ������

If Not ���������� Or ��� = "00" Then
    ����� = ����� & " "
Else
    �������� ���:=11
    ����� = ����� & " " & ��� & " " & �������(�������)
End If

' ������������� ������ ��������� ������ ������������� �����

    ������ = �����
End Function


Function ��������()
' ��������� ����������� �������� ��������� ����������� ��������� ������
'
����� = Array("", "���� ", "��� ", "�p� ", "����p� ", "���� ", "����� ", "���� ", "������ ", "������ ")
����� = Array("", "���� ", "��� ", "�p� ", "����p� ", "���� ", "����� ", "���� ", "������ ", "������ ")
�����1 = Array("������ ", "����������� ", "���������� ", "���������� ", "������������ ", "���������� ", "����������� ", "���������� ", "������������ ", "������������ ")
������� = Array("", "������ ", "�������� ", "�p������ ", "����� ", "��������� ", "���������� ", "��������� ", "����������� ", "��������� ")
����� = Array("", "��� ", "������ ", "�p���� ", "��������� ", "������� ", "�������� ", "������� ", "��������� ", "��������� ")
������ = Array("������", "������", "�����")
�������� = Array("�������", "��������", "���������")
��������� = Array("��������", "���������", "����������")
End Function

Function ��������������()
' ������������ ������ ������� 9 �������� �����

������ ���:=0
If ���������� = True Then
    �������� ���:=0
    ����� = ����� & ���������(�������) + " "
End If
������ ���:=3
If ���������� = True Then
    �������� ���:=3
    ����� = ����� & ��������(�������) + " "
End If
������ ���:=6
If ���������� = True Then
    �������� ���:=6
    ����� = ����� & ������(�������) + " "
End If
End Function

Function ������(���)
' ��������� �������������� � ����� ������ (������) �����

If Val(Mid(����������, 1 + ���, 3)) <> 0 Then
    ����� = ����� & �����(��(1 + ���))      ' ��������� ������ �����
    If ��(2 + ���) = 1 Then
        ����� = ����� & �����1(��(3 + ���)) ' ��������� ������ ����� �� 11 �� 19
    Else
        ����� = ����� & �������(��(2 + ���)) & IIf(��� = 6, _
        �����(��(3 + ���)), �����(��(3 + ���)))
        ' �������� ������ �������� �� 10 �� 90 � ������ �������� ���� _
        � ������������ �������� ����� � ������� ����
    End If
    ���������� = True
Else
    ���������� = False
End If
End Function
    
Function ��������(���)
'��������� ��������� �� ������� � ������������ � ������������� �����
'������ ��������� �� ������� (�������)

If ��(2 + ���) = 1 Then         ' �������� �� ����� �� 10 �� 19
    ������� = 2                 ' ����������, ���������, �����, ������
Else
    Select Case ��(3 + ���)     ' �������� �� ����� �� 1 �� 9
        Case 1
            ������� = 0         ' (1) ��������, �������, ������, �����
        Case 2 To 4
            ������� = 1         ' (2,3,4) ���������, ��������, ������, �����
        Case Else
            ������� = 2         ' (0,5..9) ����������, ���������, �����, ������
    End Select
End If
End Function

