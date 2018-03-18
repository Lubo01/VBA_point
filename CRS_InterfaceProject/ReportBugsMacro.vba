Sub Report_Bugs()

Dim BugsValues As String
Dim BugsRange As Range
Dim ReportCell As Range

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

BugsValues = "Pri kontrole vystupu FA Suvaha VZaS financnej analyzy sme nasli nasledovne chyby vo vypocte:"

Set BugsRange = ActiveWorkbook.Sheets("LL check FA Vystup").Range("g7:g130")
Set ReportCell = ActiveWorkbook.Sheets("LL check FA Vystup").Range("g6")

For Each i In BugsRange
    
    If i <> "" Then
        
    BugsValues = BugsValues & Chr(10) & i
    
    End If
    
    Next i
    
    Columns("G:G").EntireColumn.AutoFit
    ReportCell = BugsValues
    ReportCell.WrapText = True
    Rows("6:6").EntireRow.AutoFit
    

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

End Sub
