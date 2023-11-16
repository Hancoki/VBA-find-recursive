Public Function Find(TypeToFind As String, FindStr As Variant, Optional StartRow As Variant, Optional StartColumn As Variant, Optional Target As Variant) As Integer
     
    Dim Source As String: Source = ActiveSheet.Name
    
    Worksheets(ActiveSheet.Name).Activate ' activate current sheet if last function call was on other sheet
    
    If IsMissing(StartRow) Then StartRow = 1 ' default value if starting row is missing
    
    If IsMissing(StartColumn) Then StartColumn = 1 ' default value if starting column is missing
    
    If Not IsMissing(Target) Then Worksheets(Target).Activate ' use target sheet if it is needed
    
    With Range(Cells(StartRow, StartColumn), Cells(CountRows, CountColumns)) ' find string in your defined range
    
        Set CellResult = .Find(What:=FindStr, After:=Cells(StartRow, StartColumn), LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows)
    
    End With
    
    If TypeToFind = "r" Then ' get row if type 'r' was selected
        Find = CellResult.Row
        
    ElseIf TypeToFind = "c" Then ' get column if type 'c' was selected
        Find = CellResult.Column
        
    End If
    
    Worksheets(Source).Activate ' activate the source sheet after function call

End Function

Public Function CountRows() As Integer

    CountRows = ActiveSheet.UsedRange.Rows.Count ' Count used rows in active worksheet

End Function

Public Function CountColumns() As Integer

    CountColumns = ActiveSheet.UsedRange.Columns.Count ' Count used columns in active worksheet

End Function