Option Explicit

Sub Blend_Time_Calc()
'******************************************************************************
'******************************************************************************
'******************************************************************************


    Dim rng1 As Range, rng2 As Range, rng3 As Range, _
        rng4 As Range, rng5 As Range, rnglist As Range, _
        rngcrit As Range, r As Range
    Dim sht1 As Worksheet, sht2 As Worksheet, _
        sht3 As Worksheet, sht4 As Worksheet
    Dim fstrow As Long, lstrow As Long, L As Long, _
        iRow As Long, nxtrow As Long
    Dim sumval As Double
    Dim i As Variant, j As Variant, k As Variant, _
        vArray As Variant, St_Blend1 As Variant, _
        St_Blend2 As Variant, St_Blend3 As Variant
    
    
    Set sht1 = Sheets("Import2")
    Set sht2 = Sheets("Wrkspc")
    Set sht3 = Sheets("Wrkspc2")
    Set sht4 = Sheets("Report")
    
    Set rng1 = sht1.Range("$A$1").CurrentRegion
    Set rng2 = sht2.Range("$A$1").CurrentRegion
    Set rng3 = sht3.Range("$A$1").CurrentRegion
    Set rng4 = sht1.Range("$E$1:$E$50000")
    Set rng5 = sht4.Range("$A$1").CurrentRegion
    Set r = Range(sht1.Range("$A$1"), sht1.Range("$A$1").End(xlDown))
    
    rng2.Value = ""
    rng3.Value = ""
    rng5.Value = ""
        
        
    Application.Calculation = xlCalculationManual
'    Application.ScreenUpdating = False
        
    sht1.AutoFilterMode = False
            
    vArray = Application.Transpose(Range("StatorList"))
    
    St_Blend1 = Array("PC1009", "PC1009-IN713", "PC1010", _
                    "PC1017", "PC1022", "PC1028", _
                    "PC1033", "PC1034", "PC1036", _
                    "PC3008", "PC3010", "PC4002", _
                    "PC4005", "PC4007", "PC4007-FLNG", "PC4008", _
                    "PC4008-713")
    
    St_Blend2 = Array("HS3002", "PC2002", "PC2002-IN713")
    
    St_Blend3 = "PC1018"
    
    i = 0
    j = 0
    k = 0
    fstrow = 0
    lstrow = 0
    iRow = 0
        
    For i = LBound(vArray) To UBound(vArray)
    
        Debug.Print vArray(i)
        
        rng1.AutoFilter field:=6, _
            Criteria1:=vArray(i), _
            Operator:=xlFilterValues
            
        L = WorksheetFunction.Count(r.Cells.SpecialCells(xlCellTypeVisible))
            If L = 0 Then
                GoTo Continue_on
            End If
            
        j = vArray(i)
            If IsNumeric(Application.Match(j, St_Blend1, 0)) Then
                rng1.AutoFilter field:=7, _
                    Criteria1:="=4000", _
                    Operator:=xlFilterValues
                    
                GoTo Finish_Loop
                    
            ElseIf IsNumeric(Application.Match(j, St_Blend2, 0)) Then
                rng1.AutoFilter field:=7, _
                    Criteria1:="=3800", _
                    Operator:=xlFilterValues
                    
                GoTo Finish_Loop
                
            ElseIf j = St_Blend3 Then
                rng1.AutoFilter field:=7, _
                    Criteria1:="=3590", _
                    Operator:=xlFilterValues
                
                GoTo Finish_Loop
                    
Finish_Loop:
                    
                rng4.SpecialCells(xlCellTypeVisible).Copy
       
                With sht3
                    .Activate
                    Range("A1").PasteSpecial xlPasteValues
                    Range("A1").CurrentRegion.RemoveDuplicates Columns:=1, _
                        Header:=xlYes
                End With
                
                fstrow = rng3.Offset(1).Row
                lstrow = rng3(Rows.Count, 1).End(xlUp).Row
                
                Debug.Print fstrow, lstrow
                
                    For iRow = fstrow To lstrow
                        
                        k = sht3.Cells(iRow, 1).Value
                        
                        rng1.AutoFilter field:=5, _
                            Criteria1:=k, _
                            Operator:=xlFilterValues
                          
                        sumval = Application.WorksheetFunction.Subtotal(109, sht1.Range("L2:L50000"))
                        
                        Debug.Print sumval
                                            
                            If rng5.Cells(1, 1) = "" Then
                                nxtrow = 1
                            Else: nxtrow = rng5.Rows(Rows.Count).End(xlUp).Row + 1
                            End If
                        
                        rng5.Cells(nxtrow, 1).Value = j
                        rng5.Cells(nxtrow, 2).Value = k
                        rng5.Cells(nxtrow, 3).Value = sumval
                        
                    Next iRow
                    
                rng3.CurrentRegion.Value = ""
                
            End If

Continue_on:

    sht1.AutoFilterMode = False
                          
    Next i
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
End Sub
