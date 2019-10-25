Attribute VB_Name = "AverageDensityAndHD"
Sub Mash_Keyboard()
Attribute Mash_Keyboard.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Mash_Keyboard Macro
'

    'Set Current Workbook
    Dim x As Workbook
    Set x = ActiveWorkbook
    
    'Find and set the last used row in the data range
    Dim lr As Long
    lr = Cells(Rows.Count, "A").End(xlUp).Row
    
    'Turn off screen flickering as Excel does stuff
    Application.ScreenUpdating = False
    
    'Process and copy raw data
    Range("U1").Formula = "=CELL(""filename"")"
    Range("V1").Formula = "=SEARCH(""????_C?"",U1,1)"
    Range("W1").Formula = "=MID(U1,V1,3)"
    Range("Z1").Formula = "=MID(U1,(V1-3),2)"
    Range("AA1").Formula = "=MID(U1,V1,8)"
    Range("AB1").Formula = _
        "=IF(W1=""BLA"",MID(U1,(V1+8),4),MID(U1,(V1+8),5))"
    Range("AC1").Formula = _
        "=SUMPRODUCT(1/COUNTIF(C2:C" & lr & ",C2:C" & lr & "&""""),C2:C" & lr & ")"
    Range("AD1").Formula = _
        "=COUNTIF(K2:K" & lr & ",""mushroom"")/AC1"
    Range("AE1").Formula = "=COUNTIF(K2:K" & lr & ",""thin"")/AC1"
    Range("AF1").Formula = _
        "=AVERAGEIF(K2:K" & lr & ",""mushroom"",H2:H" & lr & ")"
    Range("AG1").Formula = _
        "=AVERAGEIF(K2:K" & lr & ",""thin"",H2:H" & lr & ")"
    Range("Z1:AG1").Copy
    
    'Pasting the Data
    Dim FileName As String
    FileName = "/Users/kylie.a.huckleberry/Desktop/MasterSpineData.xlsx"
    
    Dim y As Workbook
    
    Set y = Workbooks.Open(FileName)
    
    With y
    Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    End With
    
    'Turn Screen Updating back on and close raw data sheet without saving our changes/calculations
    Application.ScreenUpdating = True
    x.Close SaveChanges:=False
    
End Sub



