Sub Delete_Sheets()
Dim SheetCount As Integer
Dim i As Integer
Dim wildcard As String
Dim length As Integer

IdxBegin = InputBox("Type in the FIRST index of sheets you want to delete:", "Deletion of Worksheets", 2)
IdxEnd = InputBox("Type in the LAST index of sheets you want to delete:", "Deletion of Worksheets", ActiveWorkbook.Sheets.Count)
'length = Len(wildcard)

Application.DisplayAlerts = False
SheetCount = ActiveWorkbook.Sheets.Count
ReDim SheetNames(1 To SheetCount)

For i = 1 To SheetCount
    SheetNames(i) = ActiveWorkbook.Sheets(i).Name
Next i

For i = IdxBegin To IdxEnd
    'If Left(SheetNames(i), length) = wildcard Then
    'Worksheets(SheetNames(i)).Delete
    Worksheets(SheetNames(i)).Delete
    
    'End If
Next i
Application.DisplayAlerts = True
End Sub
