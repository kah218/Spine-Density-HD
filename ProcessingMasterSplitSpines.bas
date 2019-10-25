Attribute VB_Name = "ProcessingMasterSplitSpines"
Sub ProcessingMasterSplitSpines()
'Excel 2016
'ProcessingMasterSplitSpines Macro
'Mushroom vs Thin Spines have their own pages
    'Turn off screen flickering as Excel does stuff
    Application.ScreenUpdating = False
  
  'Set Current Workbook
    Dim ws As Workbook
    Set ws = ActiveWorkbook
  
    Dim lr As Long
    lr = Cells.Find(What:="*", _
                    After:=Range("A1"), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False).Row
  
    ws.Sheets("Sheet1").Name = "Original"
    
    ws.Sheets("Original").Copy After:=ws.Sheets("Original")
    ActiveSheet.Name = "Processing"
    
    'Process and copy raw data
    ws.Sheets("Processing").Range("J1").Formula = "Animal ID"
    ws.Sheets("Processing").Range("J2").Formula = "=A2"
    ws.Sheets("Processing").Range("J2").AutoFill Destination:=Range("J2:J" & lr & ""), Type:=xlFillDefault
    ws.Sheets("Processing").Range("K1").Formula = "Region"
    ws.Sheets("Processing").Range("K2").Formula = "=MID(B2, 1, 3)"
    ws.Sheets("Processing").Range("K2").AutoFill Destination:=Range("K2:K" & lr & ""), Type:=xlFillDefault
    ws.Sheets("Processing").Range("L1").Formula = "Cell"
    ws.Sheets("Processing").Range("L2").Formula = "=MID(B2, 7, 2)"
    ws.Sheets("Processing").Range("L2").AutoFill Destination:=Range("L2:L" & lr & ""), Type:=xlFillDefault
    ws.Sheets("Processing").Range("M1").Formula = "Proximal or Distal"
    
    'Does this Range need to be here? (Next 2 Lines)
    'Dim rng As Range
    'Set rng = Range("M2:M" & lr & "")
        
    Dim i As Long
    For i = 1 To lr
        If InStr(1, Cells(i, "C").Text, "P", vbTextCompare) > 0 Then
                ws.Sheets("Processing").Range("M" & i & "").Formula = "Proximal"
            ElseIf InStr(1, Cells(i, "C").Text, "D", vbTextCompare) > 0 Then
                ws.Sheets("Processing").Range("M" & i & "").Formula = "Distal"
        End If
    Next i
  
    ws.Sheets("Processing").Range("R1").Formula = "Slice"
    ws.Sheets("Processing").Range("R2").Formula = "=MID(B2, 4, 1)"
    ws.Sheets("Processing").Range("R2").AutoFill Destination:=Range("R2:R" & lr & ""), Type:=xlFillDefault
    ws.Sheets("Processing").Range("I1").Formula = "Cell Numbers Only"
    ws.Sheets("Processing").Range("I2").Formula = "=SUM(MID(0&L2,LARGE(INDEX(ISNUMBER(--MID(L2,ROW($1:$99),1))*ROW($1:$99),),ROW($1:$99))+1,1)*10^ROW($1:$99)/10)"
    ws.Sheets("Processing").Range("I2").AutoFill Destination:=Range("I2:I" & lr & ""), Type:=xlFillDefault
    ws.Sheets("Processing").Range("N1").Formula = "Concatenate_Mushroom"
    ws.Sheets("Processing").Range("N2").Formula = "=Concatenate(J2,"" "", K2,"" "",M2,"" "",R2,"" "",I2,"" "",""mushroom"")"
    ws.Sheets("Processing").Range("N2").AutoFill Destination:=Range("N2:N" & lr & ""), Type:=xlFillDefault
    ws.Sheets("Processing").Range("O1").Formula = "Concatenate_Thin"
    ws.Sheets("Processing").Range("O2").Formula = "=Concatenate(J2,"" "", K2,"" "",M2,"" "",R2,"" "",I2,"" "",""thin"")"
    ws.Sheets("Processing").Range("O2").AutoFill Destination:=Range("O2:O" & lr & ""), Type:=xlFillDefault
    'ws.Sheets("Processing").Range("P1").Formula = "For Compiling Cells into Slices Mushrooms"
    'ws.Sheets("Processing").Range("P2").Formula = "=Concatenate(J2,"" "", K2,"" "",M2,"" "",R2,"" "",""mushroom"")"
    'ws.Sheets("Processing").Range("P2").AutoFill Destination:=Range("P2:P" & lr & ""), Type:=xlFillDefault
    'ws.Sheets("Processing").Range("Q1").Formula = "For Compiling Cells into Slices Thin"
    'ws.Sheets("Processing").Range("Q2").Formula = "=Concatenate(J2,"" "", K2,"" "",M2,"" "",R2,"" "",""thin"")"
    'ws.Sheets("Processing").Range("Q2").AutoFill Destination:=Range("Q2:Q" & lr & ""), Type:=xlFillDefault
    'ws.Sheets("Processing").Range("S1").Formula = "For Compiling Slices into Regions Mushroom"
    'ws.Sheets("Processing").Range("S2").Formula = "=Concatenate(J2,"" "", K2,"" "",M2,"" "",""mushroom"")"
    'ws.Sheets("Processing").Range("S2").AutoFill Destination:=Range("S2:S" & lr & ""), Type:=xlFillDefault
    'ws.Sheets("Processing").Range("T1").Formula = "For Compiling Slices into Regions Thin"
    'ws.Sheets("Processing").Range("T2").Formula = "=Concatenate(J2,"" "", K2,"" "",M2,"" "",""thin"")"
    'ws.Sheets("Processing").Range("T2").AutoFill Destination:=Range("T2:T" & lr & ""), Type:=xlFillDefault
    ws.Sheets("Processing").Range("N2:N" & lr & "").Copy
    
    'Time to take all of this data and pull the unique MUSHROOM cell information from it
    ws.Sheets.Add After:=ws.Sheets("Processing")
    ActiveSheet.Name = "Cells_Mushroom"

    ws.Sheets("Cells_Mushroom").Range("A2:A" & lr & "").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    Dim arr As New Collection, a
    Dim aFirstArray() As Variant
    Dim k As Long

    aFirstArray = Range("A2:A" & lr & "")

    On Error Resume Next
    For Each a In aFirstArray
    arr.Add a, a
    Next

    For k = 1 To arr.Count
    ws.Sheets("Cells_Mushroom").Cells(1, k) = arr(k)
    Next
    
    Range("A2:A" & lr & "").delete
    Rows("1:1").Select
    Range("AD1").Activate
    Selection.Columns.AutoFit
    
    Columns("A:A").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlToRight
    ws.Sheets("Cells_Mushroom").Range("A3").Formula = "Density"
    ws.Sheets("Cells_Mushroom").Range("A4").Formula = "HD"
    
    'Time to take all of this data and pull the unique THIN cell information from it
    ws.Sheets.Add After:=ws.Sheets("Cells_Mushroom")
    ActiveSheet.Name = "Cells_Thin"
    
    ws.Sheets("Processing").Range("O2:O" & lr & "").Copy
    ws.Sheets("Cells_Thin").Range("A2:A" & lr & "").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    Dim arr2 As New Collection, a2
    Dim aFirstArray2() As Variant
    Dim k2 As Long

    aFirstArray2 = Range("A2:A" & lr & "")

    On Error Resume Next
    For Each a2 In aFirstArray2
    arr2.Add a2, a2
    Next

    For k2 = 1 To arr2.Count
    ws.Sheets("Cells_Thin").Cells(1, k2) = arr2(k2)
    Next
    
    Range("A2:A" & lr & "").delete
    Rows("1:1").Select
    Range("AD1").Activate
    Selection.Columns.AutoFit
    
    Columns("A:A").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlToRight
    ws.Sheets("Cells_Thin").Range("A3").Formula = "Density"
    ws.Sheets("Cells_Thin").Range("A4").Formula = "HD"
        
    'Time to take all of this data and compile it by cell (Mushroom and then thin)
    ws.Sheets("Cells_Mushroom").Select
    Dim lc_cells As Long
    lc_cells = ws.Sheets("Cells_Mushroom").Cells(1, Columns.Count).End(xlToLeft).Column

        Sheets("Cells_Mushroom").Select
               
    'ws.Sheets("Cells_Mushroom").Range(Cells(1, 2), Cells(1, lc_cells)).Copy
    'ws.Sheets("Processing").Cells(1, 27).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
    '    Sheets("Processing").Select
    
    'With ws.Sheets("Processing")
    'Dim avg_cells_mush As Long
    '    For avg_cells_mush = 27 To (27 + lc_cells)
    '        Cells(2, avg_cells_mush) = WorksheetFunction.AverageIf(Range(Cells(2, 14), Cells(lr, 14)), Cells(1, avg_cells_mush), Range(Cells(2, 5), Cells(lr, 5)))
    '    Next avg_cells_mush
    'End With
    
    Dim ws2 As Worksheet
    Set ws2 = ws.Sheets("Processing")
    Dim ws3 As Worksheet
    Set ws3 = ws.Sheets("Cells_Mushroom")
    Dim ws4 As Worksheet
    Set ws4 = ws.Sheets("Cells_Thin")
    Dim avg_cells_mush As Long
    For avg_cells_mush = 2 To lc_cells
        ws3.Select
            ws3.Cells(3, avg_cells_mush).Formula = Application.WorksheetFunction.AverageIf(ws2.Range(ws2.Cells(2, 14), ws2.Cells(lr, 14)), ws3.Cells(1, avg_cells_mush), ws2.Range(ws2.Cells(2, 5), ws2.Cells(lr, 5)))
            ws3.Cells(4, avg_cells_mush).Formula = Application.WorksheetFunction.AverageIf(ws2.Range(ws2.Cells(2, 14), ws2.Cells(lr, 14)), ws3.Cells(1, avg_cells_mush), ws2.Range(ws2.Cells(2, 7), ws2.Cells(lr, 7)))
        ws4.Select
            ws4.Cells(3, avg_cells_mush).Formula = Application.WorksheetFunction.AverageIf(ws2.Range(ws2.Cells(2, 15), ws2.Cells(lr, 15)), ws4.Cells(1, avg_cells_mush), ws2.Range(ws2.Cells(2, 6), ws2.Cells(lr, 6)))
            ws4.Cells(4, avg_cells_mush).Formula = Application.WorksheetFunction.AverageIf(ws2.Range(ws2.Cells(2, 15), ws2.Cells(lr, 15)), ws4.Cells(1, avg_cells_mush), ws2.Range(ws2.Cells(2, 8), ws2.Cells(lr, 8)))
    Next avg_cells_mush
    
        ws3.Select
    
    'Create labels to use to sort into regions for both Mushroom and Thin
    Dim region_mush As Long
    For region_mush = 2 To lc_cells
        If InStr(1, ws3.Cells(1, region_mush), "Proximal", vbTextCompare) > 0 Then
                ws3.Cells(2, region_mush).FormulaR1C1 = "=CONCATENATE(MID(R1C" & region_mush & ", 1, 15),"" "",""mushroom"")"
        ElseIf InStr(1, ws3.Cells(1, region_mush), "Distal", vbTextCompare) > 0 Then
                 ws3.Cells(2, region_mush).FormulaR1C1 = "=CONCATENATE(MID(R1C" & region_mush & ", 1, 13),"" "",""mushroom"")"
        Else: ws3.Cells(2, region_mush).FormulaR1C1 = "=CONCATENATE(MID(R1C" & region_mush & ", 1, 6),"" "",""mushroom"")"
        End If
        If InStr(1, ws4.Cells(1, region_mush), "Proximal", vbTextCompare) > 0 Then
                ws4.Cells(2, region_mush).FormulaR1C1 = "=CONCATENATE(MID(R1C" & region_mush & ", 1, 15),"" "",""thin"")"
        ElseIf InStr(1, ws4.Cells(1, region_mush), "Distal", vbTextCompare) > 0 Then
                 ws4.Cells(2, region_mush).FormulaR1C1 = "=CONCATENATE(MID(R1C" & region_mush & ", 1, 13),"" "",""thin"")"
        Else: ws4.Cells(2, region_mush).FormulaR1C1 = "=CONCATENATE(MID(R1C" & region_mush & ", 1, 6),"" "",""thin"")"
        End If
    Next region_mush
    
    'Pull unique headers for regions-Mushroom
    ws3.Range(Cells(2, 2), Cells(2, lc_cells)).Copy
    ws.Sheets.Add After:=ws4
    ActiveSheet.Name = "Regions_Mushroom"
    Dim ws5 As Worksheet
    Set ws5 = ws.Sheets("Regions_Mushroom")
    ws5.Range(Cells(2, 2), Cells(2, lc_cells)).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    Dim arr3 As New Collection, a3
    Dim aFirstArray3() As Variant
    Dim k3 As Long

    aFirstArray3 = Range(Cells(2, 2), Cells(2, lc_cells))

    On Error Resume Next
    For Each a3 In aFirstArray3
    arr3.Add a3, a3
    Next

    For k3 = 1 To arr3.Count
    ws.Sheets("Regions_Mushroom").Cells(1, k3) = arr3(k3)
    Next
    
    Range(Cells(2, 2), Cells(2, lc_cells)).delete
    Rows("1:1").Select
    Range("AD1").Activate
    Selection.Columns.AutoFit
    
    Columns("A:A").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlToRight
    ws.Sheets("Regions_Mushroom").Range("A2").Formula = "Density"
    ws.Sheets("Regions_Mushroom").Range("A3").Formula = "HD"

    'Pull unique headers for regions-Thin
        ws4.Select
    ws4.Range(Cells(2, 2), Cells(2, lc_cells)).Copy
    ws.Sheets.Add After:=ws5
    ActiveSheet.Name = "Regions_Thin"
    Dim ws6 As Worksheet
    Set ws6 = ws.Sheets("Regions_Thin")
    ws6.Range(Cells(2, 2), Cells(2, lc_cells)).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
        Sheets("Regions_Thin").Select
    
    Dim arr4 As New Collection, a4
    Dim aFirstArray4() As Variant
    Dim k4 As Long

    aFirstArray4 = Range(Cells(2, 2), Cells(2, lc_cells))

    On Error Resume Next
    For Each a4 In aFirstArray4
    arr4.Add a4, a4
    Next

    For k4 = 1 To arr4.Count
    ws.Sheets("Regions_Thin").Cells(1, k4) = arr4(k4)
    Next
    
    Range(Cells(2, 2), Cells(2, lc_cells)).delete
    Rows("1:1").Select
    Range("AD1").Activate
    Selection.Columns.AutoFit
    
    Columns("A:A").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlToRight
    ws.Sheets("Regions_Thin").Range("A2").Formula = "Density"
    ws.Sheets("Regions_Thin").Range("A3").Formula = "HD"
    
    'Averages for the regions
        ws5.Select
    Dim lc_region As Long
    lc_region = ws.Sheets("Regions_Mushroom").Cells(1, Columns.Count).End(xlToLeft).Column
    Dim avg_region As Long
    For avg_region = 2 To lc_region
        ws5.Select
            ws5.Cells(2, avg_region).Formula = Application.WorksheetFunction.AverageIf(ws3.Range(ws3.Cells(2, 2), ws3.Cells(2, lc_cells)), ws5.Cells(1, avg_region), ws3.Range(ws3.Cells(3, 2), ws3.Cells(3, lc_region)))
            ws5.Cells(3, avg_region).Formula = Application.WorksheetFunction.AverageIf(ws3.Range(ws3.Cells(2, 2), ws3.Cells(2, lc_cells)), ws5.Cells(1, avg_region), ws3.Range(ws3.Cells(4, 2), ws3.Cells(4, lc_region)))
        ws6.Select
            ws6.Cells(2, avg_region).Formula = Application.WorksheetFunction.AverageIf(ws4.Range(ws4.Cells(2, 2), ws4.Cells(2, lc_cells)), ws6.Cells(1, avg_region), ws4.Range(ws4.Cells(3, 2), ws4.Cells(3, lc_region)))
            ws6.Cells(3, avg_region).Formula = Application.WorksheetFunction.AverageIf(ws4.Range(ws4.Cells(2, 2), ws4.Cells(2, lc_cells)), ws6.Cells(1, avg_region), ws4.Range(ws4.Cells(4, 2), ws4.Cells(4, lc_region)))
    Next avg_region
    
    'Turn back on screen flickering as Excel does stuff
    Application.ScreenUpdating = True
    ws.Save

End Sub
