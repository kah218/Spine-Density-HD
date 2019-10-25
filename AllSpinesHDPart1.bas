Attribute VB_Name = "AllSpineHDPart1"
Sub ForBeccaCompiled()
'Excel 2016
' ForBeccaCompiled Macro
'CompilesSpineHeadDensities
'
  'Set Current Workbook
    Dim X As Workbook
    Set X = ActiveWorkbook
    
    'Find and set the last used row (lr) in the data range
    Dim lr As Long
    lr = Cells(Rows.Count, "A").End(xlUp).Row
    
    'Turn off screen flickering as Excel does stuff
    Application.ScreenUpdating = False
    
    'Process raw data
    Range("U1").Formula = "=CELL(""filename"")"
    Range("V1").Formula = "=SEARCH(""????_C?"",U1,1)"
    Range("W1").Formula = "=MID(U1,V1,3)"
    Range("X1").Formula = "=MID(U1,(V1-3),2)"
    Range("Y1").Formula = "=X1&"" ""&W1&"" ""&""mushroom"""
    Range("Z1").Formula = "=X1&"" ""&W1&"" ""&""thin"""
    Range("Y2").Formula = "=IF(K2 = ""mushroom"",H2,"""")"
    Range("Y2").AutoFill Destination:=Range("Y2:Y" & lr & ""), Type:=xlFillDefault
    Range("Z2").Formula = "=IF(K2 = ""thin"",H2,"""")"
    Range("Z2").AutoFill Destination:=Range("Z2:Z" & lr & ""), Type:=xlFillDefault
    
    'Remove blanks and copy raw data
    'CollMush = Collapse Mushroom & CollThin = Collapse Thin
    Range("Y1:Z" & lr & "").Copy
    Range("AA1:AB" & lr & "").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    Dim CollMush As Long
    For CollMush = lr To 1 Step -1
    If Cells(CollMush, "AA") = "" Then
        Range("AA" & CollMush & ":AA" & CollMush & "").delete
    End If
    Next CollMush

    Dim CollThin As Long
    For CollThin = lr To 1 Step -1
    If Cells(CollThin, "AB") = "" Then
        Range("AB" & CollThin & ":AB" & CollThin & "").delete
    End If
    Next CollThin
    
    Range("AA1:AB" & lr & "").Copy
      
    'Pasting the Data
    Dim FileName As String
    FileName = "/Users/kylie.a.huckleberry/Desktop/MasterSpineData.xlsx"
    
    Dim y As Workbook
    
    Set y = Workbooks.Open(FileName)
    
    With y
    Cells(1, Columns.Count).End(xlToLeft).Offset(0, 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Selection.Columns.AutoFit
    End With
    
    'Compile Columns
    'Dim lr2 As Long
    'lr2 = Cells(Rows.Count, "A").End(xlUp).Row
    
    'Dim NLC As Long
    'NLC = Cells(1, Columns.Count).End(xlToLeft).Offset(0, -1)
    'Dim lc As Long
    'lc = Cells(1, Columns.Count).End(xlToLeft).Offset(0, 0)
    
    'With Range("A1:LC")
    
    'Dim arr As New Collection, a
    'Dim aFirstArray() As Variant
    'Dim i As Long

    'aFirstArray = Range("A1:A245")

    'On Error Resume Next
    'For Each a In aFirstArray
    'arr.Add a, a
    'Next

    'For i = 1 To arr.Count
    'Cells(1, i) = arr(i)
    'Next
    
    'Turn Screen Updating back on and close raw data sheet without saving our changes/calculations
    Application.ScreenUpdating = True
    Application.DisplayAlerts = False
    X.Close SaveChanges:=False
    Application.DisplayAlerts = True
    y.Save
    

End Sub
