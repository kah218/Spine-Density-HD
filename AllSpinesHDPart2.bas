Attribute VB_Name = "AllSpineHDPart2"
Sub AllSpineHDPart2()
    Application.ScreenUpdating = False
    
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    ActiveSheet.Name = "All Spines Duplicate Columns"
    
    Dim lr As Long
    lr = Cells.Find(What:="*", _
                    After:=Range("A1"), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False).Row
    
    Dim lc As Long
    lc = wb.Sheets("All Spines Duplicate Columns").Cells(1, Columns.Count).End(xlToLeft).Column
    
    Range(Cells(1, "B"), Cells(1, lc)).Copy
    
    wb.Sheets.Add After:=wb.Sheets("All Spines Duplicate Columns")
    ActiveSheet.Name = "All Spines Compiled"
    wb.Sheets("All Spines Compiled").Range(Cells(2, "A"), Cells(2, lc)).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    
    'Pull the Unique Headers and paste onto the new sheet.
    Dim arr As New Collection, c
    Dim cFirstArray() As Variant
    Dim i As Long

    cFirstArray = Range(Cells(2, "B"), Cells(2, lc))

    On Error Resume Next
    For Each c In cFirstArray
    arr.Add c, c
    Next

    For i = 1 To arr.Count
    wb.Sheets("All Spines Compiled").Cells(1, i) = arr(i)
    Next

    Range(Cells(2, "A"), Cells(2, lc)).EntireRow.delete
    Rows("1:1").Select
    Range("AD1").Activate
    Selection.Columns.AutoFit
    
    
    'FIND COLUMNS MATCHING NEW HEADERS AND PASTE CONTEXTS ONTO NEW SHEET.
    
    
            Sheets("All Spines Duplicate Columns").Select
    
    Dim t As Long
    Dim s As Long
    For t = 1 To lc
        For s = 2 To lc
            Dim lr_source
            lr_source = wb.Sheets("All Spines Duplicate Columns").Cells(Rows.Count, s).End(xlUp).Row
            If wb.Sheets("All Spines Duplicate Columns").Cells(1, s).Value = wb.Sheets("All Spines Compiled").Cells(1, t).Value Then
                wb.Sheets("All Spines Duplicate Columns").Range(Cells(2, s), Cells(lr_source, s)).Copy
                wb.Sheets("All Spines Compiled").Cells(Rows.Count, t).End(xlUp).Offset(1, 0).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            End If
        Next s
    Next t

    'Turn Screen Updating back on and close raw data sheet without saving our changes/calculations
    Application.ScreenUpdating = True
    wb.Save
        
End Sub
