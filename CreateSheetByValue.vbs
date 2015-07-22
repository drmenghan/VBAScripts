Sub CreateSheet()
    Dim MyCell As Range, MyRange As Range
    
    RangeBegin = InputBox("Type in the position of your FIRST new Sheet you want to create:", "Creation of Worksheets", "A2")
     
    Set MyRange = Sheets(1).Range(RangeBegin)
    Set MyRange = Range(MyRange, MyRange.End(xlDown))

    For Each MyCell In MyRange
        
        Sheets.Add After:=Sheets(Sheets.Count) 'creates a new worksheet
        NewSheetName = Left(Replace(MyCell.Value, " ", ""), 15)
        Sheets(Sheets.Count).Name = NewSheetName ' renames the new worksheet
    Next MyCell
    
