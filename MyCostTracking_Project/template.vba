
Sub CrisInterfaceDraft2a()

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

On Error Resume Next

Application.ActiveWorkbook.Sheets("FA_Súvaha a VZaS").Select

Dim period As Integer       'porovnavacie obdobie (2 az 4)
Dim periodFactor As Integer 'posun stlpcov pre porovnanie
Dim periodColumns As String 'oznacenie stlpcov ktore sa porovnavaju
Dim periodNumber As Integer 'pocet obdobi vo vystupe Crisu
Dim periodNumberPL As Integer 'pocet obdobi vo vystupe Crisu P&L
Dim months As Integer       'pocet mesiacov v obdobi

period = 2
periodFactor = 4
periodNumber = 0
periodNumberPL = 0
months = 0

'kontrola poctu obdobi v Cris vystupe
'Suvaha
If Cells(14, 9) <> 0 Then
periodNumber = 4
ElseIf Cells(14, 7) <> 0 Then
periodNumber = 3
ElseIf Cells(14, 5) <> 0 Then
periodNumber = 2
ElseIf Cells(14, 3) <> 0 Then
periodNumber = 1
Else
periodNumber = 0
End If
'P&L
If Cells(103, 9) <> 0 Then
periodNumberPL = 4
ElseIf Cells(103, 7) <> 0 Then
periodNumberPL = 3
ElseIf Cells(103, 5) <> 0 Then
periodNumberPL = 2
ElseIf Cells(103, 3) <> 0 Then
periodNumberPL = 1
Else
periodNumberPL = 0
End If

'generovanie hlavicky porovnania (riadky 1-11 + riadok 100, stlpce L-Y)
For period = 2 To periodNumber
    If periodNumber < 2 Then
    Exit For
    End If
        
    If period = 2 Then
    periodColumns = "e/c"
    ElseIf period = 3 Then
    periodColumns = "g/e"
    ElseIf period = 4 Then
    periodColumns = "i/g"
    Else
    periodColumns = "N/A"
    End If

Cells(1, period + 10 + periodFactor * (period - 2)) = "Period"  'starts with Cells(1, 12)
Cells(1, period + 11 + periodFactor * (period - 2)) = period    'Cells(1, 13)
Cells(1, period + 12 + periodFactor * (period - 2)) = periodColumns     'Cells(1, 14)
'Cells(1, 15)= Cells(3,5)   '=(E3) vykaz k datumu
Cells(1, period + 13 + periodFactor * (period - 2)) = Cells(3, period + 3 + (periodFactor - 3) * (period - 2))
Cells(1, period + 13 + periodFactor * (period - 2)).NumberFormat = "d/m/yyyy"   'Cells(1, 15).NumberFormat
'Cells(2, 13)= Cells(4,5)   '=(E4) pocet mesiacov
Cells(2, period + 10 + periodFactor * (period - 2)) = "Months"  'Cells(2, 12)
Cells(2, period + 11 + periodFactor * (period - 2)) = Cells(4, period + 3 + (periodFactor - 3) * (period - 2))    'Cells(2, 13)
 
Cells(3, period + 10 + periodFactor * (period - 2)) = "BS 5% rel"     'Cells(3, 12), BS relevance
'Cells(3, 13)=0.05 * Cells(12, 5)   '=E12 celkove aktiva
Cells(3, period + 11 + periodFactor * (period - 2)) = 0.05 * Cells(14, period + 3 + (periodFactor - 3) * (period - 2))   'Cells(3, 13)
Cells(3, period + 12 + periodFactor * (period - 2)) = "P&L 5% rel"    'Cells(3, 14), P&L relevance
'Cells(3, 15)=0.05 * Cells(101, 5)   '=E101 celkove trzby
Cells(3, period + 13 + periodFactor * (period - 2)) = 0.05 * Cells(103, period + 3 + (periodFactor - 3) * (period - 2))  'Cells(3, 15)

Cells(4, period + 10 + periodFactor * (period - 2)) = "BS 2% rel"     'Cells(4, 12), BS relevance
'Cells(4, 13)=0.02 * Cells(12, 5)   '=E12 celkove aktiva
Cells(4, period + 11 + periodFactor * (period - 2)) = 0.02 * Cells(14, period + 3 + (periodFactor - 3) * (period - 2))   'Cells(4, 13)
Cells(4, period + 12 + periodFactor * (period - 2)) = "Profit 50% rel"    'Cells(4, 14), Profit relevance
'Cells(4, 15)=0.5 * Cells(144, 5)   '=E144 cisty zisk/strata
Cells(4, period + 13 + periodFactor * (period - 2)) = 0.5 * Cells(146, period + 3 + (periodFactor - 3) * (period - 2))   'Cells(4, 15)

Cells(13, period + 10 + periodFactor * (period - 2)) = "y/y"    'Cells(13, 12)
Cells(13, period + 11 + periodFactor * (period - 2)) = "share"  'Cells(13, 13)
Cells(13, period + 12 + periodFactor * (period - 2)) = "delta"  'Cells(13, 14)
Cells(13, period + 13 + periodFactor * (period - 2)) = "delta pp"   'Cells(13, 15)

Cells(102, period + 10 + periodFactor * (period - 2)) = "y/y"    'Cells(102, 12)
Cells(102, period + 11 + periodFactor * (period - 2)) = "share"  'Cells(102, 13)
Cells(102, period + 12 + periodFactor * (period - 2)) = "delta"  'Cells(102, 14)
Cells(102, period + 13 + periodFactor * (period - 2)) = "delta pp"   'Cells(102, 15)

Next

'Suvaha - porovnanie poloziek (generovanie a vypocet vzorcov)

For i = 14 To 100
    If periodNumber < 2 Then
    Exit For
    End If
    
    For period = 2 To periodNumber
    
    'y/y comparison, in column L: Cells(i, 12) = Cells(i, 5) / Cells(i, 3) - 1
    If Cells(i, period + 1 + (periodFactor - 3) * (period - 2)).Value = 0 Then      'Cells (14,3)
       Cells(i, period + 10 + periodFactor * (period - 2)) = "N/A"  'Cells (14,12)
    'Cells(14, 12) = Cells(14, 5) / Cells(14, 3) - 1
    Else
    Cells(i, period + 10 + periodFactor * (period - 2)) = _
        Cells(i, period + 3 + (periodFactor - 3) * (period - 2)) / Cells(i, period + 1 + (periodFactor - 3) * (period - 2)) - 1
    End If
        
    Cells(i, period + 10 + periodFactor * (period - 2)).NumberFormat = "0.0%"   'Cells(14, 12)
    
    'share comparison in column M: Cells(i, 13) = Cells(i, 5) / Cells(14, 5)
    'set 2nd e14 to $ (alias fix) through fixed row index
    Cells(i, period + 11 + periodFactor * (period - 2)) = _
    Cells(i, period + 3 + (periodFactor - 3) * (period - 2)) / Cells(14, period + 3 + (periodFactor - 3) * (period - 2))
    Cells(i, period + 11 + periodFactor * (period - 2)).NumberFormat = "0.0%"
    
    'delta comparison in column N: Cells(i, 14) = Cells(i, 5) - Cells(i, 3)
    Cells(i, period + 12 + periodFactor * (period - 2)) = _
    Cells(i, period + 3 + (periodFactor - 3) * (period - 2)) - Cells(i, period + 1 + (periodFactor - 3) * (period - 2))
    Cells(i, period + 12 + periodFactor * (period - 2)).NumberFormat = "#,##0"
    
    'delta pp comparison in column O: Cells(i, 15) = Cells(i, 13) - Cells(i, 3) / Cells(14, 3)
    Cells(i, period + 13 + periodFactor * (period - 2)) = _
    Cells(i, period + 11 + periodFactor * (period - 2)) - Cells(i, period + 1 + (periodFactor - 3) * (period - 2)) _
    / Cells(14, period + 1 + (periodFactor - 3) * (period - 2)) 'set 2nd c14 to $ (alias fix) through fixed row index
    Cells(i, period + 13 + periodFactor * (period - 2)).NumberFormat = "0.0%"
    
    Next
Next

'P&L - porovnanie poloziek (generovanie a vypocet vzorcov)
For i = 103 To 146
    If periodNumberPL < 2 Then
    Exit For
    End If
    
    For period = 2 To periodNumberPL
    
    'y/y comparison, in column L: Cells(i, 12) = Cells(i, 5) / Cells(i, 3) - 1
    If Cells(i, period + 1 + (periodFactor - 3) * (period - 2)).Value = 0 Then      'Cells (103,3)
       Cells(i, period + 10 + periodFactor * (period - 2)) = "N/A"  'Cells (103,12)
''''''
       End If  'temporary
       
    'kontrola poctu mesiacov v Cells (2,13)

'    ElseIf Cells(2, period + 11 + periodFactor * (period - 2)) = 12 Then
'    'ak je pocet mesiacov 12
'    'Cells(12, 12) = Cells(12, 5) / Cells(12, 3) - 1
    Cells(i, period + 10 + periodFactor * (period - 2)) = _
        Cells(i, period + 3 + (periodFactor - 3) * (period - 2)) / Cells(i, period + 1 + (periodFactor - 3) * (period - 2)) - 1
'    Else
'    'ak je pocet mesiacov iny ako 12
'    'Cells(12, 12) = (Cells(12, 5)/Cells(2,13)*12 / Cells(12, 3) - 1
'    Cells(i, period + 10 + periodFactor * (period - 2)) = _
'        (Cells(i, period + 3 + (periodFactor - 3) * (period - 2)) / Cells(2, period + 11 + periodFactor * (period - 2)) * 12) _
'        / Cells(i, period + 1 + (periodFactor - 3) * (period - 2)) - 1
'    End If
    
    Cells(i, period + 10 + periodFactor * (period - 2)).NumberFormat = "0.0%"   'Cells(12, 12)
    
    'share comparison in column M: Cells(i, 13) = Cells(i, 5) / Cells(103, 5)
    'set 2nd e103 to $ (alias fix) through fixed row index
    Cells(i, period + 11 + periodFactor * (period - 2)) = _
    Cells(i, period + 3 + (periodFactor - 3) * (period - 2)) / Cells(103, period + 3 + (periodFactor - 3) * (period - 2))
    Cells(i, period + 11 + periodFactor * (period - 2)).NumberFormat = "0.0%"
        '-------

    'delta comparison in column N: Cells(i, 14) = Cells(i, 5) - Cells(i, 3)
''''''
'    'kontrola poctu mesiacov v Cells (2,13)
'    If Cells(2, period + 11 + periodFactor * (period - 2)) = 12 Then
'    'ak je pocet mesiacov 12
'    'Cells(i, 14) = Cells(i, 5) - Cells(i, 3)
    Cells(i, period + 12 + periodFactor * (period - 2)) = _
    Cells(i, period + 3 + (periodFactor - 3) * (period - 2)) - Cells(i, period + 1 + (periodFactor - 3) * (period - 2))
'    Else
'    'ak je pocet mesiacov iny ako 12
'    'Cells(i, 14) = Cells(i, 5) - (Cells(i, 3)/12*Cells(2,13))
'    Cells(i, period + 12 + periodFactor * (period - 2)) = _
'    Cells(i, period + 3 + (periodFactor - 3) * (period - 2)) _
'    - (Cells(i, period + 1 + (periodFactor - 3) * (period - 2)) / 12 * Cells(2, period + 11 + periodFactor * (period - 2)))
'    End If
    Cells(i, period + 12 + periodFactor * (period - 2)).NumberFormat = "#,##0"
    
    'delta pp comparison in column O: Cells(i, 15) = Cells(i, 13) - Cells(i, 3) / Cells(101, 3)
    Cells(i, period + 13 + periodFactor * (period - 2)) = _
    Cells(i, period + 11 + periodFactor * (period - 2)) - Cells(i, period + 1 + (periodFactor - 3) * (period - 2)) _
    / Cells(103, period + 1 + (periodFactor - 3) * (period - 2)) 'set 2nd c103 to $ (alias fix) through fixed row index
    Cells(i, period + 13 + periodFactor * (period - 2)).NumberFormat = "0.0%"
    
    Next
    
Next

'Range("L12:O12").Select
'    Selection.AutoFill Destination:=Range("L12:O144"), Type:=xlFillDefault
'-----------------------------------------------------------------------------------------------------------------------------

'formatovanie
ActiveWorkbook.Sheets("FA_Súvaha a VZaS").Activate
'Cells.Select 'vyber celeho harku
Range("A1:Z150").Select
With Selection.Font
        .Name = "Arial"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With

'hlavicka
Range("L1:Y4").Select
With Selection
    .Font.Bold = True
    .WrapText = True
End With
Range("M3,M4,O3,O4,R3,R4,T3,T4,W3,W4,Y3,Y4").Select
Selection.NumberFormat = "#,##0"

'format zony pre porovnanie
'ciary tabulky
Range("L13:Y146").Select
    Selection.Font.Bold = False
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

'podhlavicka (subheader)
Range("L13:Y13").Select
Selection.Font.Bold = True
Selection.Interior.Color = 5296274

Range("L102:Y102").Select
Selection.Font.Bold = True
Selection.Interior.Color = 5296274
'With Selection.Interior
'        .Pattern = xlSolid
'        .PatternColorIndex = xlAutomatic
'        .Color = 5296274
'        .TintAndShade = 0
'        .PatternTintAndShade = 0
'    End With

'podmieneny format zony porovnania - Balance sheet
For i = 14 To 100
    If periodNumber < 2 Then
    Exit For
    End If
    
    For period = 2 To periodNumber
    
    'stlpec M: share on Balance sheet
    If Cells(i, period + 11 + periodFactor * (period - 2)) >= 0.05 Then
    Cells(i, period + 11 + periodFactor * (period - 2)).Font.Bold = True
    Cells(i, period + 11 + periodFactor * (period - 2)).Interior.Color = 65535
    ElseIf Cells(i, period + 11 + periodFactor * (period - 2)) >= 0.02 Then
    Cells(i, period + 11 + periodFactor * (period - 2)).Interior.ThemeColor = xlThemeColorAccent2
    Cells(i, period + 11 + periodFactor * (period - 2)).Interior.TintAndShade = 0.799981688894314
    End If
    'stlpec N: delta on previous Balance sheet
    If Cells(i, period + 12 + periodFactor * (period - 2)) >= Cells(3, 13) Or _
    Cells(i, period + 12 + periodFactor * (period - 2)) <= -Cells(3, 13) Then
    Cells(i, period + 12 + periodFactor * (period - 2)).Font.Bold = True
    Cells(i, period + 12 + periodFactor * (period - 2)).Interior.Color = 65535
    ElseIf Cells(i, period + 12 + periodFactor * (period - 2)) >= Cells(4, 13) Or _
    Cells(i, period + 12 + periodFactor * (period - 2)) <= -Cells(4, 13) Then
    Cells(i, period + 12 + periodFactor * (period - 2)).Interior.ThemeColor = xlThemeColorAccent2
    Cells(i, period + 12 + periodFactor * (period - 2)).Interior.TintAndShade = 0.799981688894314
    End If
    
    Next
Next

'podmieneny format zony porovnania - P&L
For i = 103 To 146
    If periodNumber < 2 Then
    Exit For
    End If
    
    For period = 2 To periodNumber
    
    'stlpec M: share on Revenues
    If Cells(i, period + 11 + periodFactor * (period - 2)) >= 0.05 Then
    Cells(i, period + 11 + periodFactor * (period - 2)).Font.Bold = True
    Cells(i, period + 11 + periodFactor * (period - 2)).Interior.Color = 65535
    ElseIf Cells(i, period + 11 + periodFactor * (period - 2)) >= 0.02 Then
    Cells(i, period + 11 + periodFactor * (period - 2)).Interior.ThemeColor = xlThemeColorAccent2
    Cells(i, period + 11 + periodFactor * (period - 2)).Interior.TintAndShade = 0.799981688894314
    End If
    'stlpec N: delta on previous P&L
    If Cells(i, period + 12 + periodFactor * (period - 2)) >= Cells(3, 15) Or _
    Cells(i, period + 12 + periodFactor * (period - 2)) <= -Cells(3, 15) Then
    Cells(i, period + 12 + periodFactor * (period - 2)).Font.Bold = True
    Cells(i, period + 12 + periodFactor * (period - 2)).Interior.Color = 65535
    ElseIf Cells(i, period + 12 + periodFactor * (period - 2)) >= Cells(4, 15) Or _
    Cells(i, period + 12 + periodFactor * (period - 2)) <= -Cells(4, 15) Then
    Cells(i, period + 12 + periodFactor * (period - 2)).Interior.ThemeColor = xlThemeColorAccent2
    Cells(i, period + 12 + periodFactor * (period - 2)).Interior.TintAndShade = 0.799981688894314
    End If
    
    Next
Next

'skratenie popisu riadkov v stlpci B, t.j. vypustenie anglickych ekvivalentov
Call XSplitLabelShort

'skrytie stlpca A (kody poloziek v CRIS
Columns("A:A").Select
Selection.EntireColumn.Hidden = True

'uprava sirky stlpcov tabulky B:J
Columns("B:B").Select
With Selection
        .ColumnWidth = 30
End With
Columns("D:J").EntireColumn.AutoFit
'uprava sirky stlpcov porovnania M:Z
Columns("M:Z").EntireColumn.AutoFit

'zoskupenie a skrytie riadkov 5-12
Range("L5:L12").Select
    Selection.Rows.Group
    Selection.Rows.Hidden = True
    
'ukotvenie priecok
    Range("D14").Select
    ActiveWindow.FreezePanes = True
    
'uprava medzistlpcov L, Q a V
    Range("L:L, Q:Q, V:V").Select
    With Selection
        .ColumnWidth = 1.11
        .Borders(xlInsideHorizontal).LineStyle = xlNone
        .Interior.Pattern = xlNone
        .Interior.TintAndShade = 0
        .Interior.PatternTintAndShade = 0
    End With
    
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

End Sub


Sub XSplitLabelShort()

Range("C1").Select
Selection.EntireColumn.Insert

Range("B14:B146").Select
Selection.Replace What:=" / ", Replacement:="@"
Selection.TextToColumns , DataType:=xlDelimited, Other:=True, OtherChar:="@"
Columns("C:C").Select
Selection.EntireColumn.Hidden = True
Range("B1").Select
End Sub


Sub XHideZeroRows()

For i = 14 To 146
    If Cells(i, 4) + Cells(i, 6) + Cells(i, 8) + Cells(i, 10) = 0 Then
        Rows(i).Hidden = True
    End If
Next

Rows(102).Hidden = False

End Sub
Sub XHideZeroRowsVstupy()

For i = 16 To 220
    If Cells(i, 3) + Cells(i, 4) + Cells(i, 5) + Cells(i, 6) = 0 Then
        Rows(i).Hidden = True
    End If
Next

Rows(93).Hidden = False
Rows(158).Hidden = False
Rows(159).Hidden = False

End Sub

Sub XFormatVstupy()

For i = 16 To 157
    For j = 3 To 6
        If Cells(i, j) >= (0.05 * Cells(16, j)) Then
        Cells(i, j).Interior.Color = 65535
        End If
    Next
Next

For i = 160 To 220
    For j = 3 To 6
        If Cells(i, j) >= (0.05 * (Cells(160, j) + Cells(163, j))) Then
        Cells(i, j).Interior.Color = 65535
        End If
    Next
Next

End Sub



Sub XShowAllRows()

Range("A14:A146").EntireRow.Hidden = False

End Sub

Sub CRIS_RatiosFlags()

SalesRed = 0        'max
SalesYellow1 = 3    'max
SalesYellow2 = 30   'min
EbitdaRed = 5       'max
EbitdaYellow = 10   'max
RoeRed = 5          'max
RoeYellow = 10      'max
DebtorDaysRed = 90  'min
DebtorDaysYellow = 60   'min
StockDaysRed = 60       'min
StockDaysYellow = 30    'min
CreditorDaysRed = 90    'min
CreditorDaysYellow = 60 'min
DebtEbitdaRed = 5       'min
DebtEbitdaYellow = 3    'min
InterestCoverageRed = 2 'max
InterestCoverageYellow = 5 'max
DscrRed = 1         'max
DscrYellow = 1.5    'max
EquityRed = 10      'max
EquityYellow = 20      'max
GearingRed = 200    'min
GearingYellow = 150    'min
LiquidityRed = 1    'max
LiquidityYellow = 1.3   'max

  
If Cells(13, 1) = "FAU001" Then
   For i = 3 To 6
        If Cells(13, i) < SalesRed Then
            Cells(13, i).Interior.Color = 255
        ElseIf Cells(13, i) < SalesYellow1 Then
            Cells(13, i).Interior.Color = 65535
        ElseIf Cells(13, i) > SalesYellow2 Then
            Cells(13, i).Interior.Color = 65535
        End If
        If Cells(15, i) < EbitdaRed Then
            Cells(15, i).Interior.Color = 255
        ElseIf Cells(15, i) < EbitdaYellow Then
            Cells(15, i).Interior.Color = 65535
        End If
        If Cells(16, i) < RoeRed Then
            Cells(16, i).Interior.Color = 255
        ElseIf Cells(16, i) < RoeYellow Then
            Cells(16, i).Interior.Color = 65535
        End If
        If Cells(18, i) > DebtorDaysRed Then
            Cells(18, i).Interior.Color = 255
        ElseIf Cells(18, i) > DebtorDaysYellow Then
            Cells(18, i).Interior.Color = 65535
        End If
        If Cells(19, i) > StockDaysRed Then
            Cells(19, i).Interior.Color = 255
        ElseIf Cells(19, i) > StockDaysYellow Then
            Cells(19, i).Interior.Color = 65535
        End If
        If Cells(20, i) > CreditorDaysRed Then
            Cells(20, i).Interior.Color = 255
        ElseIf Cells(20, i) > CreditorDaysYellow Then
            Cells(20, i).Interior.Color = 65535
        End If
        If Cells(22, i) > DebtEbitdaRed Then
            Cells(22, i).Interior.Color = 255
        ElseIf Cells(22, i) > DebtEbitdaYellow Then
            Cells(22, i).Interior.Color = 65535
        End If
        If Cells(23, i) < InterestCoverageRed Then
            Cells(23, i).Interior.Color = 255
        ElseIf Cells(23, i) < InterestCoverageYellow Then
            Cells(23, i).Interior.Color = 65535
        End If
        If Cells(24, i) < DscrRed Then
            Cells(24, i).Interior.Color = 255
        ElseIf Cells(24, i) < DscrYellow Then
            Cells(24, i).Interior.Color = 65535
        End If
        If Cells(25, i) < EquityRed Then
            Cells(25, i).Interior.Color = 255
        ElseIf Cells(25, i) < EquityYellow Then
            Cells(25, i).Interior.Color = 65535
        End If
        If Cells(26, i) > GearingRed Then
            Cells(26, i).Interior.Color = 255
        ElseIf Cells(26, i) > GearingYellow Then
            Cells(26, i).Interior.Color = 65535
        End If
        If Cells(28, i) < LiquidityRed Then
            Cells(28, i).Interior.Color = 255
        ElseIf Cells(28, i) < LiquidityYellow Then
            Cells(28, i).Interior.Color = 65535
        End If
                
    Next
Else
MsgBox ("FAU001 not found")
End If
   
End Sub


Sub CRIS_RatiosFlags2()

'verzia pre vystup z CRISu od 15.4.2016

'Definition of treshold values for Ratios:
SalesRed = 0        'max
SalesYellow1 = 3    'max
SalesYellow2 = 30   'min
EbitdaRed = 5       'max
EbitdaYellow = 10   'max
RoeRed = 5          'max
RoeYellow = 10      'max
DebtorDaysRed = 90  'min
DebtorDaysYellow = 60   'min
StockDaysRed = 60       'min
StockDaysYellow = 30    'min
CreditorDaysRed = 90    'min
CreditorDaysYellow = 60 'min
DebtEbitdaRed = 5       'min
DebtEbitdaYellow = 3    'min
InterestCoverageRed = 2 'max
InterestCoverageYellow = 5 'max
DscrRed = 1         'max
DscrYellow = 1.5    'max
EquityRed = 10      'max
EquityYellow = 20      'max
GearingRed = 200    'min
GearingYellow = 150    'min
LiquidityRed = 1    'max
LiquidityYellow = 1.3   'max

'kontrola poctu obdobi v Cris vystupe
'Ukazovatele
If Cells(5, 7) <> 0 Then
periodNumber = 5
ElseIf Cells(5, 6) <> 0 Then
periodNumber = 4
ElseIf Cells(5, 5) <> 0 Then
periodNumber = 3
ElseIf Cells(5, 4) <> 0 Then
periodNumber = 2
Else
periodNumber = 1
End If

'Convert text data to number format:
Dim MyNumberValue As Double

Sheets("FA_Ukazovatele").Activate

    For i = 13 To 60
        For j = 3 To periodNumber + 2
            MyNumberValue = Cells(i, j).Value
            Cells(i, j).Value = MyNumberValue
        Next j
    Next i
'MsgBox ("Numbers ready!")

'Flag values according to defined tresholds above:
If Cells(15, 1) = "FAU001" Then
   For i = 3 To periodNumber + 2
        If Cells(15, i) < SalesRed Then
            Cells(15, i).Interior.Color = 255
        ElseIf Cells(15, i) < SalesYellow1 Then
            Cells(15, i).Interior.Color = 65535
        ElseIf Cells(15, i) > SalesYellow2 Then
            Cells(15, i).Interior.Color = 65535
        End If
        If Cells(17, i) < EbitdaRed Then
            Cells(17, i).Interior.Color = 255
        ElseIf Cells(17, i) < EbitdaYellow Then
            Cells(17, i).Interior.Color = 65535
        End If
        If Cells(18, i) < RoeRed Then
            Cells(18, i).Interior.Color = 255
        ElseIf Cells(18, i) < RoeYellow Then
            Cells(18, i).Interior.Color = 65535
        End If
        If Cells(20, i) > DebtorDaysRed Then
            Cells(20, i).Interior.Color = 255
        ElseIf Cells(20, i) > DebtorDaysYellow Then
            Cells(20, i).Interior.Color = 65535
        End If
        If Cells(21, i) > StockDaysRed Then
            Cells(21, i).Interior.Color = 255
        ElseIf Cells(21, i) > StockDaysYellow Then
            Cells(21, i).Interior.Color = 65535
        End If
        If Cells(22, i) > CreditorDaysRed Then
            Cells(22, i).Interior.Color = 255
        ElseIf Cells(22, i) > CreditorDaysYellow Then
            Cells(22, i).Interior.Color = 65535
        End If
        If Cells(24, i) > DebtEbitdaRed Then
            Cells(24, i).Interior.Color = 255
        ElseIf Cells(24, i) > DebtEbitdaYellow Then
            Cells(24, i).Interior.Color = 65535
        End If
        If Cells(25, i) < InterestCoverageRed Then
            Cells(25, i).Interior.Color = 255
        ElseIf Cells(25, i) < InterestCoverageYellow Then
            Cells(25, i).Interior.Color = 65535
        End If
        If Cells(26, i) < DscrRed Then
            Cells(26, i).Interior.Color = 255
        ElseIf Cells(26, i) < DscrYellow Then
            Cells(26, i).Interior.Color = 65535
        End If
        If Cells(27, i) < EquityRed Then
            Cells(27, i).Interior.Color = 255
        ElseIf Cells(27, i) < EquityYellow Then
            Cells(27, i).Interior.Color = 65535
        End If
        If Cells(28, i) > GearingRed Then
            Cells(28, i).Interior.Color = 255
        ElseIf Cells(28, i) > GearingYellow Then
            Cells(28, i).Interior.Color = 65535
        End If
        If Cells(30, i) < LiquidityRed Then
            Cells(30, i).Interior.Color = 255
        ElseIf Cells(30, i) < LiquidityYellow Then
            Cells(30, i).Interior.Color = 65535
        End If
                
    Next
Else
MsgBox ("FAU001 not found")
End If
   
End Sub

---


Sub SaveCopyAsBackUp_BezMakier_Word()

ActiveDocument.SaveCopyAs "L:\Docs\9 Back up\" & Left(ActiveDocument.Name, Len(ActiveDocument.Name) - 5) & " backup.docx"
                                                                          
MsgBox "Copy Saved as: " & Left(ActiveDocument.Name, _
                           Len(ActiveDocument.Name) - 5) & " backup.docx" _
                           & " in L:\Docs\9 Back up\"
                                     
End Sub

'Sub SaveCopyAsBackUp_SMakrami()
'
'ActiveWorkbook.SaveCopyAs "L:\Docs\9 Back up\" & Left(ActiveWorkbook.Name, Len(ActiveWorkbook.Name) - 5) & " backup.xlsm"
'
'MsgBox "Copy Saved as: " & Left(ActiveWorkbook.Name, _
'                                     Len(ActiveWorkbook.Name) - 5) & " backup.xlsm" _
'                                     & " in L:\Docs\9 Back up\"
'
'End Sub


