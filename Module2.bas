Attribute VB_Name = "Module1"
 Function ChangeNumber()
   Dim selectedCell As Range
    Dim selectedColumn As Range
    Dim cell As Range
    Dim cellContent As String
    Dim n As Integer
    Dim ins As String
    n = 9
    ins = "80"
    
    Set selectedCell = Selection
    
    Set selectedColumn = selectedCell.EntireColumn
    
    For Each cell In selectedColumn.Cells
         cellContent = cell.Value
         cellContent = Replace(cellContent, " ", "")
         cellContent = Replace(cellContent, "-", "")
         If Len(cellContent) >= n Then
            Dim lastNChars As String
            lastNChars = Right(cellContent, n)
            
            Dim modifiedContent As String
            modifiedContent = ins & lastNChars
            
            
           cell.Value = modifiedContent
   
         End If
    Next cell
    
End Function
