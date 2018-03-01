Sub ConcatenateAnalysis()

Dim OutputTextEN As String
Dim AnalysisPeriod_9EN As Range
Dim TargetCellEN As Range
Dim OutputTextSK As String
Dim AnalysisPeriod_9SK As Range
Dim TargetCellSK As Range

Dim OutputTextEN8 As String
Dim AnalysisPeriod_8EN As Range
Dim TargetCellEN8 As Range
Dim OutputTextSK8 As String
Dim AnalysisPeriod_8SK As Range
Dim TargetCellSK8 As Range

Dim OutputTextEN7 As String
Dim AnalysisPeriod_7EN As Range
Dim TargetCellEN7 As Range
Dim OutputTextSK7 As String
Dim AnalysisPeriod_7SK As Range
Dim TargetCellSK7 As Range

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

'Analysis with English labels, quarter 9

OutputTextEN = "Period"

Set AnalysisPeriod_9EN = ActiveWorkbook.Sheets("Analysis").Range("K17:K75")
Set TargetCellEN = ActiveWorkbook.Sheets("Analysis").Range("K3")

For Each i In AnalysisPeriod_9EN
    
    If i <> "" Then
        
    OutputTextEN = OutputTextEN & Chr(10) & i
    
    End If
    
    Next i
    
    TargetCellEN = OutputTextEN
    TargetCellEN.WrapText = True
    Rows("3:3").EntireRow.AutoFit
    
'Analysis EN, quarter 8

OutputTextEN8 = "Period"

Set AnalysisPeriod_8EN = ActiveWorkbook.Sheets("Analysis").Range("I17:I75")
Set TargetCellEN8 = ActiveWorkbook.Sheets("Analysis").Range("I3")

For Each i In AnalysisPeriod_8EN
    
    If i <> "" Then
        
    OutputTextEN8 = OutputTextEN8 & Chr(10) & i
    
    End If
    
    Next i
    
    TargetCellEN8 = OutputTextEN8
    TargetCellEN8.WrapText = True
    Rows("3:3").EntireRow.AutoFit
    
'Analysis EN, quarter 7

OutputTextEN7 = "Period"

Set AnalysisPeriod_7EN = ActiveWorkbook.Sheets("Analysis").Range("G17:G75")
Set TargetCellEN7 = ActiveWorkbook.Sheets("Analysis").Range("G3")

For Each i In AnalysisPeriod_7EN
    
    If i <> "" Then
        
    OutputTextEN7 = OutputTextEN7 & Chr(10) & i
    
    End If
    
    Next i
    
    TargetCellEN7 = OutputTextEN7
    TargetCellEN7.WrapText = True
    Rows("3:3").EntireRow.AutoFit
     
'Analysis with Slovak labels, quarter 9

OutputTextSK = "Period"

Set AnalysisPeriod_9SK = ActiveWorkbook.Sheets("Analysis").Range("L17:L75")
Set TargetCellSK = ActiveWorkbook.Sheets("Analysis").Range("L3")

For Each i In AnalysisPeriod_9SK
    
    If i <> "" Then
        
    OutputTextSK = OutputTextSK & Chr(10) & i
    
    End If
    
    Next i
    
    TargetCellSK = OutputTextSK
    TargetCellSK.WrapText = True
    Rows("3:3").EntireRow.AutoFit
    

'Analysis SK, quarter 8

OutputTextSK8 = "Period"

Set AnalysisPeriod_8SK = ActiveWorkbook.Sheets("Analysis").Range("J17:J75")
Set TargetCellSK8 = ActiveWorkbook.Sheets("Analysis").Range("J3")

For Each i In AnalysisPeriod_8SK
    
    If i <> "" Then
        
    OutputTextSK8 = OutputTextSK8 & Chr(10) & i
    
    End If
    
    Next i
    
    TargetCellSK8 = OutputTextSK8
    TargetCellSK8.WrapText = True
    Rows("3:3").EntireRow.AutoFit
    
'Analysis SK, quarter 7

OutputTextSK7 = "Period"

Set AnalysisPeriod_7SK = ActiveWorkbook.Sheets("Analysis").Range("H17:H75")
Set TargetCellSK7 = ActiveWorkbook.Sheets("Analysis").Range("H3")

For Each i In AnalysisPeriod_7SK
    
    If i <> "" Then
        
    OutputTextSK7 = OutputTextSK7 & Chr(10) & i
    
    End If
    
    Next i
    
    TargetCellSK7 = OutputTextSK7
    TargetCellSK7.WrapText = True
    Rows("3:3").EntireRow.AutoFit

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

End Sub




Sub ImportConverterDataCA()

Dim ClientName
Dim ConverterZdroj_1 As Range
Dim ImportData_1 As Range
Dim ConverterSubor_1
Dim Interface As Workbook
Dim ConverterZdroj_2 As Range
Dim ImportData_2 As Range
Dim ConverterSubor_2


Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual


ClientName = InputBox("Zadaj nazov klienta:")

'zadame cestu a skopirujeme Converter 1

Set ConverterSubor_1 = Workbooks.Open("P:\03_COMMERCIAL BANKING_\3100_Podpora predaja\1021 Kreditne analyzy\Kreditne Analyzy Klientov\" & ClientName & "\IFRS Converter 1 " & ClientName & ".xlsm")
MsgBox "Vybrany subor je: " & vbNewLine & (ActiveWorkbook.FullName)

Set Interface = Workbooks.Open("P:\03_COMMERCIAL BANKING_\3100_Podpora predaja\1021 Kreditne analyzy\Kreditne Analyzy Klientov\" & ClientName & "\!Financials interface for CA " & ClientName & ".xlsm")

'zadame zdroj, ciel a skopirujeme vykazy do interface
Set ConverterZdroj_1 = ConverterSubor_1.Sheets("Slovakia").Range("B5:F215")
Set ImportData_1 = Interface.Sheets("Converter Input").Range("B5:F215")
ConverterZdroj_1.Copy ImportData_1

'Interface.Save
'CAzdroj.Close
'Interface.Close

MsgBox "Converter 1 input ready!"

'zadame cestu a skopirujeme Converter 2

Set ConverterSubor_2 = Workbooks.Open("P:\03_COMMERCIAL BANKING_\3100_Podpora predaja\1021 Kreditne analyzy\Kreditne Analyzy Klientov\" & ClientName & "\IFRS Converter 2 " & ClientName & ".xlsm")
MsgBox "Vybrany subor je: " & vbNewLine & (ActiveWorkbook.FullName)

'zadame zdroj, ciel a skopirujeme vykazy do interface
Set ConverterZdroj_2 = ConverterSubor_2.Sheets("Slovakia").Range("B5:F215")
Set ImportData_2 = Interface.Sheets("Converter Input").Range("H5:L215")
ConverterZdroj_2.Copy ImportData_2

MsgBox "Ready! Finally!"


Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

End Sub

Sub ImportConverterDataAddon()

Dim BorrowerName
Dim AddonClientName
Dim ConverterZdroj_1 As Range
Dim ImportData_1 As Range
Dim ConverterSubor_1
Dim Interface As Workbook
Dim ConverterZdroj_2 As Range
Dim ImportData_2 As Range
Dim ConverterSubor_2

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual


BorrowerName = InputBox("Zadaj nazov klienta - dlznika:")
AddonClientName = InputBox("Zadaj nazov Add-on klienta:")

'zadame cestu a skopirujeme Converter 1

Set ConverterSubor_1 = Workbooks.Open("P:\03_COMMERCIAL BANKING_\3100_Podpora predaja\1021 Kreditne analyzy\Kreditne Analyzy Klientov\" & BorrowerName & "\IFRS Converter 1 " & AddonClientName & ".xlsm")
MsgBox "Vybrany subor je: " & vbNewLine & (ActiveWorkbook.FullName)

Set Interface = Workbooks.Open("P:\03_COMMERCIAL BANKING_\3100_Podpora predaja\1021 Kreditne analyzy\Kreditne Analyzy Klientov\" & BorrowerName & "\!Financials interface for Addon " & AddonClientName & ".xlsm")

'zadame zdroj, ciel a skopirujeme vykazy do interface
Set ConverterZdroj_1 = ConverterSubor_1.Sheets("Slovakia").Range("B5:F215")
Set ImportData_1 = Interface.Sheets("Converter Input").Range("B5:F215")
ConverterZdroj_1.Copy ImportData_1

'Interface.Save
'CAzdroj.Close
'Interface.Close

MsgBox "Converter 1 input ready!"

'zadame cestu a skopirujeme Converter 2

Set ConverterSubor_2 = Workbooks.Open("P:\03_COMMERCIAL BANKING_\3100_Podpora predaja\1021 Kreditne analyzy\Kreditne Analyzy Klientov\" & BorrowerName & "\IFRS Converter 2 " & AddonClientName & ".xlsm")
MsgBox "Vybrany subor je: " & vbNewLine & (ActiveWorkbook.FullName)

'zadame zdroj, ciel a skopirujeme vykazy do interface
Set ConverterZdroj_2 = ConverterSubor_2.Sheets("Slovakia").Range("B5:F215")
Set ImportData_2 = Interface.Sheets("Converter Input").Range("H5:L215")
ConverterZdroj_2.Copy ImportData_2

MsgBox "Ready! Finally!"


Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

End Sub


Sub ImportConverterCA_Quarterly()

Dim ClientName
Dim ConverterZdroj_Q As Range
Dim ImportData_Q As Range
Dim ConverterSubor_Q
Dim Interface As Workbook

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

ClientName = InputBox("Zadaj nazov klienta:")

'zadame cestu a skopirujeme Converter 1

Set ConverterSubor_Q = Workbooks.Open("P:\03_COMMERCIAL BANKING_\3100_Podpora predaja\1021 Kreditne analyzy\Kreditne Analyzy Klientov\" & ClientName & "\Converter Q " & ClientName & ".xlsm")
MsgBox "Vybrany subor je: " & vbNewLine & (ActiveWorkbook.FullName)

Set Interface = Workbooks.Open("P:\03_COMMERCIAL BANKING_\3100_Podpora predaja\1021 Kreditne analyzy\Kreditne Analyzy Klientov\" & ClientName & "\!Financials interface for CA " & ClientName & ".xlsm")

'zadame zdroj, ciel a skopirujeme vykazy do interface
Set ConverterZdroj_Q = ConverterSubor_Q.Sheets("Slovakia").Range("P16:T215")
Set ImportData_Q = Interface.Sheets("Converter Input").Range("N16:R215")
ConverterZdroj_Q.Copy ImportData_Q

'Interface.Save
'CAzdroj.Close
'Interface.Close

MsgBox "Converter Q input ready!"

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

End Sub





Sub OneColumnData_CopyAndTranspose()

Dim shift As Integer
Dim myRange As Range
Dim xRow As Integer
Dim yColumn As Integer
Dim NumRows As Integer

'
'Application.ScreenUpdating = False
'Application.Calculation = xlCalculationManual

'oblast dat ktore chcem kopirovat
Set myRange = Range("A1").CurrentRegion

'zaciatok relevantnych dat po nadpise
shift = 3
NumRows = myRange.Rows.Count

For i = shift To NumRows
    'textove polia ostavaju v prvom stlpci, nekopiruju sa
            
        If yColumn = 0 Then xRow = i
            
        yColumn = yColumn + 1
        
        myRange.Cells(xRow, yColumn) = myRange.Cells(i, 1)
        
        If yColumn = 9 Then yColumn = 0
          
Next
        


'Sub skuska()

 
'
'Dim myRange As Range
'
'Dim bunka As Range
'
'
'
'Set myRange = Range("a1").CurrentRegion
'
'pocet = myRange.Rows.Count
'
'Perioda = 9
'
'posun = 3

 
'For i = posun To pocet
'
'    stlpec = stlpec + 1
'
'    myRange.Cells(i, stlpec).Value = myRange.Cells(i, 1).Value
'
'    If stlpec = 9 Then stlpec = 0
'
'Next

 'End Sub
 
 
'Ak to ma ist do jedneho riadku napr. prveho tak mala uprava
'
'
'For i = posun To pocet
'
'    If stlpec = 0 Then riadok = i
'
'    stlpec = stlpec + 1
'
'    myRange.Cells(riadok, stlpec).Value = myRange.Cells(i, 1).Value
'
'    If stlpec = 9 Then stlpec = 0
'
'Next
'
'
End Sub



Sub OpenFileAndImportDataToNewFileCA()

Dim CAzdroj As Range
Dim ImportData As Range
Dim CA As Workbook
Dim Interface As Workbook
Dim ClientName As String
Dim Period As String


Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

'newCA1 As TextFrame
'otvorime uverovy navrh s vykazmi po zadani cesty
'MsgBox "Zadaj cestu pre otvorenie navrhu"
'CA = InputBox("zadaj cestu k navrhu:")
'Application.GetOpenFilename ("")
ClientName = InputBox("Zadaj nazov klienta:")
Period = InputBox("Zadaj obdobie navrhu: (syntax: 2014_04)")
Set CA = Workbooks.Open("P:\03_COMMERCIAL BANKING_\3100_Podpora predaja\1021 Kreditne analyzy\Kreditne Analyzy Klientov\" & ClientName & "\Credit Application_" & Period & "_" & ClientName & "_v18.xlsm")
'ActiveWorkbook.ChangeFileAccess Mode:=xlReadWrite, Notify:=True

MsgBox "Vybrany subor je: " & vbNewLine & (ActiveWorkbook.FullName)

'otvorime novy adresar pre analyzu
'ChDir ("P:\03_COMMERCIAL BANKING_\3100_Podpora predaja\1021 Kreditne analyzy\Kreditne Analyzy Klientov\newCA1")

'zalozime novy interface a ulozime ho v novom adresari
Set Interface = Workbooks.Open("L:\Docs\1 Templates\11C Financials interface\!Financials interface for CA Template updated.xlsm")
Interface.SaveAs "P:\03_COMMERCIAL BANKING_\3100_Podpora predaja\1021 Kreditne analyzy\Kreditne Analyzy Klientov\" & ClientName & "\!Financials interface for CA " & ClientName & ".xlsm"
MsgBox "Novy Interface ulozeny na P:\...\Kreditne Analyzy Klientov\" & ClientName & "\"

'zadame zdroj, ciel a skopirujeme vykazy do interface
Set CAzdroj = CA.Sheets("Financials").Range("A3:N164")
Set ImportData = Interface.Sheets("First Data Input CA").Range("A1:N162")
CAzdroj.Copy ImportData

Interface.Close SaveChanges:=True
CA.Close SaveChanges:=False

MsgBox "Ready! Finally!"

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

End Sub


Sub OpenFileAndImportDataToNewFileAddon()

Dim CAzdroj As Range
Dim ImportData As Range
Dim CA_Addon As Workbook
Dim CA_AddonName
Dim AddonClientName
Dim Interface As Workbook
Dim BorrowerName

Application.ScreenUpdating = False
'Application.Calculation = xlCalculationManual

'newCA1 As TextFrame
'otvorime uverovy navrh s vykazmi po zadani cesty
'MsgBox "Zadaj cestu pre otvorenie navrhu"
BorrowerName = InputBox("Zadaj nazov klienta - dlznika:")
AddonClientName = InputBox("Zadaj nazov Add-on klienta:")
CA_AddonName = InputBox("Zadaj nazov pre Addon subor (bez pripony):")

'Application.GetOpenFilename ("P:\03_COMMERCIAL BANKING_\3100_Podpora predaja\1021 Kreditne analyzy\Kreditne Analyzy Klientov\" & BorrowerName & "\" & CA_Addon & ".xlsm")
Set CA_Addon = Workbooks.Open("P:\03_COMMERCIAL BANKING_\3100_Podpora predaja\1021 Kreditne analyzy\Kreditne Analyzy Klientov\" & BorrowerName & "\" & CA_AddonName & ".xlsm")
'ActiveWorkbook.ChangeFileAccess Mode:=xlReadWrite, Notify:=True

MsgBox "Vybrany subor je: " & vbNewLine & (ActiveWorkbook.FullName)

'otvorime novy adresar pre analyzu
'ChDir ("P:\03_COMMERCIAL BANKING_\3100_Podpora predaja\1021 Kreditne analyzy\Kreditne Analyzy Klientov\newCA1")

'zalozime novy interface a ulozime ho v novom adresari
Set Interface = Workbooks.Open("L:\Docs\1 Templates\11C Financials interface\!Financials interface for CA Template updated.xlsm")
Interface.SaveAs "P:\03_COMMERCIAL BANKING_\3100_Podpora predaja\1021 Kreditne analyzy\Kreditne Analyzy Klientov\" & BorrowerName & "\!Financials interface for Addon " & AddonClientName & ".xlsm"
MsgBox "Novy Interface ulozeny na P:\...\Kreditne Analyzy Klientov\" & ClientName & "\"

'zadame zdroj, ciel a skopirujeme vykazy do interface
Set CAzdroj = CA_Addon.Sheets("Financials").Range("A3:N164")
Set ImportData = Interface.Sheets("First Data Input CA").Range("A1:N162")
CAzdroj.Copy ImportData

Interface.Close SaveChanges:=True
CA_Addon.Close SaveChanges:=False

MsgBox "Ready! Finally!"

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

End Sub


Sub OpenFileAndImportDataToNewFileCUSM()

Dim CAzdroj As Range
Dim ImportData As Range
Dim CA_Addon As Workbook
Dim CA_AddonName
Dim AddonClientName
Dim Interface As Workbook
Dim BorrowerName

Application.ScreenUpdating = False
'Application.Calculation = xlCalculationManual

'newCA1 As TextFrame
'otvorime uverovy navrh s vykazmi po zadani cesty
'MsgBox "Zadaj cestu pre otvorenie navrhu"
BorrowerName = InputBox("Zadaj nazov klienta - dlznika:")
AddonClientName = InputBox("Zadaj nazov Customer monitoring klienta:")
CA_AddonName = InputBox("Zadaj nazov pre Customer monitoring subor (bez pripony):")

'Application.GetOpenFilename ("P:\03_COMMERCIAL BANKING_\3100_Podpora predaja\1021 Kreditne analyzy\Kreditne Analyzy Klientov\" & BorrowerName & "\" & CA_Addon & ".xlsm")
Set CA_Addon = Workbooks.Open("P:\03_COMMERCIAL BANKING_\3100_Podpora predaja\1021 Kreditne analyzy\Kreditne Analyzy Klientov\" & BorrowerName & "\" & CA_AddonName & ".xlsm")
'ActiveWorkbook.ChangeFileAccess Mode:=xlReadWrite, Notify:=True

MsgBox "Vybrany subor je: " & vbNewLine & (ActiveWorkbook.FullName)

'otvorime novy adresar pre analyzu
'ChDir ("P:\03_COMMERCIAL BANKING_\3100_Podpora predaja\1021 Kreditne analyzy\Kreditne Analyzy Klientov\newCA1")

'zalozime novy interface a ulozime ho v novom adresari
Set Interface = Workbooks.Open("L:\Docs\1 Templates\11C Financials interface\!Financials interface for CUSM Template updated.xlsm")
Interface.SaveAs "P:\03_COMMERCIAL BANKING_\3100_Podpora predaja\1021 Kreditne analyzy\Kreditne Analyzy Klientov\" & BorrowerName & "\!Financials interface for CUSM " & AddonClientName & ".xlsm"
MsgBox "Novy Interface ulozeny na P:\...\Kreditne Analyzy Klientov\" & ClientName & "\"

'zadame zdroj, ciel a skopirujeme vykazy do interface
Set CAzdroj = CA_Addon.Sheets("Financial statements").Range("A3:N159")
Set ImportData = Interface.Sheets("First Data Input CA").Range("A1:N157")
CAzdroj.Copy ImportData

Interface.Close SaveChanges:=True
CA_Addon.Close SaveChanges:=False

MsgBox "Ready! Finally!"

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

End Sub



Sub SaveCopyAsBackUp_BezMakier()

      ActiveWorkbook.SaveCopyAs "P:\03_COMMERCIAL BANKING_\3100_Podpora predaja\1021 Kreditne analyzy\Kreditne Analyzy Klientov\BACKUP\" & Left(ActiveWorkbook.Name, Len(ActiveWorkbook.Name) - 5) & " backup.xlsx"
                                                                          
      MsgBox "Copy Saved as: " & Left(ActiveWorkbook.Name, _
                                     Len(ActiveWorkbook.Name) - 5) & " backup.xlsx" _
                                     & " in P:\...\Kreditne Analyzy Klientov\BACKUP\"
                                     
End Sub


Sub SaveCopyAsBackUp_SMakrami()

      ActiveWorkbook.SaveCopyAs "P:\03_COMMERCIAL BANKING_\3100_Podpora predaja\1021 Kreditne analyzy\Kreditne Analyzy Klientov\BACKUP\" & Left(ActiveWorkbook.Name, Len(ActiveWorkbook.Name) - 5) & " backup.xlsm"
                                                                          
      MsgBox "Copy Saved as: " & Left(ActiveWorkbook.Name, _
                                     Len(ActiveWorkbook.Name) - 5) & " backup.xlsm" _
                                     & " in P:\...\Kreditne Analyzy Klientov\BACKUP\"
                                     
End Sub


Public Sub UnprotectSheetAndCells()

ActiveSheet.Unprotect

Cells.Select
Selection.Locked = False
Range("A1").Select

End Sub

