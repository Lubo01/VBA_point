Sub XBankStatementDataFeed()


mesiac = InputBox("Vyber mesiac na CostTracking v tvare MMYY:")

'!!!check if name already exist, if yes, question and instructions
Sheets.Add(, After:=Sheets("Export")).Name = "Cost " & mesiac

Set target = Sheets("Cost " & mesiac)
Set Source = Sheets("export")
'MsgBox "target is " & target.Name & " and source is " & Source.Name

'write Header for Cost Sheet
Cells(1, 2) = "Datum"
Cells(1, 3) = "Debet/Credit"
Cells(1, 4) = "Druh nakladu"
Cells(1, 5) = "Suma"
Cells(1, 6) = "Popis"
Cells(1, 7) = "Poznamka"
Columns("B:E").ColumnWidth = 15
Columns("F:G").ColumnWidth = 45

'select Sheet export, cancel previous filter and filter transactions of selected month
Sheets("export").Select
If (ActiveSheet.AutoFilterMode And ActiveSheet.FilterMode) Or ActiveSheet.FilterMode Then
  ActiveSheet.ShowAllData
End If
Selection.AutoFilter Field:=6, Criteria1:="=*" & mesiac & "*"

Cells(1, 6).Select
rowscount = ActiveCell.CurrentRegion.Rows.Count
'MsgBox rowscount
cost = 2

'check valid data and data transfer to target cost sheet
For Row = 2 To rowscount
    ActiveCell.Offset(rowOffset:=1).Activate
    If ActiveCell.EntireRow.Hidden = False Then
        target.Cells(cost, 2).Value = Source.Cells(Row, 6).Value    'datum
'        If Source.Cells(Row, 6).Value > 0 Then
'                target.Cells(cost, 3).Value = "Credit"   'Debet/Credit
'        Else: target.Cells(cost, 3).Value = "Debet"
'        End If
                
        target.Cells(cost, 5).Value = Source.Cells(Row, 8).Value    'suma
        target.Cells(cost, 6).Value = Source.Cells(Row, 16).Value    'popis
        target.Cells(cost, 7).Value = Source.Cells(Row, 20).Value    'ucel
        
        cost = cost + 1
    End If
    
Next

target.Select
Call XCheckExpenseType    'druh nakladu
'ActiveWorkbook.Save


End Sub

Sub XCheckExpenseType()


Set byvanie = New Collection 'Byvanie
    byvanie.Add "Nas Dom"
    byvanie.Add "ZSE"
    byvanie.Add "SPP"
    byvanie.Add "Orange"
    
    
Set potraviny = New Collection  'Potraviny (vratane drogerie)
    potraviny.Add "Terno"
    potraviny.Add "TERNO"
    potraviny.Add "COOP"
    potraviny.Add "TESCO,KE"
    potraviny.Add "TESCO,PN"
    potraviny.Add "EXP TESCO"
    potraviny.Add "TESCO,BA"
    potraviny.Add "BILLA"
    potraviny.Add "KAUFLAND"
    potraviny.Add "LIDL"
    potraviny.Add "ds6"
    potraviny.Add "MICHAL RANIN"
    potraviny.Add "Diligentia"
    potraviny.Add "DM"
    potraviny.Add "101 DROGERIE"
    potraviny.Add "Potraviny"


Set doprava = New Collection 'Doprava
    doprava.Add "SHELL"
    doprava.Add "LUKOIL"
    doprava.Add "OMV"
    doprava.Add "SLOVNAFT"
    doprava.Add "CS ENI"
    doprava.Add "TESCO CS"
    doprava.Add "DPB"
    doprava.Add "ZSR"
    doprava.Add "STUDENT AGEN"
    doprava.Add "ZELEZNICNA"
    doprava.Add "eznamka"
    doprava.Add "ContiTrade"
    doprava.Add "ZSSK"
    doprava.Add "Komunalna poistovna"
    doprava.Add "Kooperativa poistovn"
    doprava.Add "AUTO - IMPEX"
    doprava.Add "Auto - Impex"
    
        
Set volny = New Collection  'Volny cas
    volny.Add "REST"    'Restauracia a MOTOREST
    volny.Add "HOTEL"
    volny.Add "PENZION"
    volny.Add "PUB"
    volny.Add "PIZZA"
    volny.Add "PUB"
    volny.Add "PLAVAREN"
    volny.Add "KSP Zralok"
    volny.Add "Kvety"
    volny.Add "KIKA RESTAUR"    'check duplicity KIKA RESTAUR s REST
    volny.Add "WWW.MOBIL.IF"
    volny.Add "FUNCITY"
    volny.Add "LD BALNEA"
    volny.Add "SAROSPAT"
    volny.Add "FCPA"
    volny.Add "MASARYKOV DV"
 
     
Set oblecenie = New Collection 'Oblecenie
    oblecenie.Add "H&M"
    oblecenie.Add "CCC"
    oblecenie.Add "MARKS"
    oblecenie.Add "SPORTISIMO"
    oblecenie.Add "SPORTSDIRECT"
    oblecenie.Add "PREDAJNA PRO"

    
Set ostKD = New Collection  'Ostatne KD
    ostKD.Add "LEKA" 'vratane LEKAREN
    ostKD.Add "SUNPHARMA"
    ostKD.Add "HORNBACH"
    ostKD.Add "JEDALEN LACHOVA"
    ostKD.Add "MS Lachova"
    ostKD.Add "WWW.ORANGE"
    ostKD.Add "DU Bratislava"
    ostKD.Add "Rodicovske zdruzenie"
    ostKD.Add "DRUG PHARMA"
    ostKD.Add "HLAVNE MESTO SR Dan"
    ostKD.Add "MEDICATE"
    ostKD.Add "BENU"
    ostKD.Add "OBI"
    ostKD.Add "IMUNOVA"
    ostKD.Add "LEK."
    ostKD.Add "Allianz"
   
    
Set ostDD = New Collection  'Ostatne DD
    ostDD.Add "KIKA,GALVANI"
    ostDD.Add "IKEA"
    ostDD.Add "BILINGVI"
    ostDD.Add "Bilingvi"
    ostDD.Add "DRACIK"
    ostDD.Add "Panta"
    ostDD.Add "MARTINUS"
    ostDD.Add "ALZA"
    ostDD.Add "INTERNET MAL"
    ostDD.Add "ELEKTROSPED"
    ostDD.Add "SIKO"
    ostDD.Add "NAY"
    ostDD.Add "AMAZON"
    ostDD.Add "KNIHY"

    
'Set vyplata, sporenie, vybery - ATM, splatky - kreditna karta KR.KAR, uvery SPL.UVERU, SPL.UR
    
Cells(1, 6).Select
rowscount = ActiveCell.CurrentRegion.Rows.Count
    For i = 2 To rowscount
        data = Cells(i, 6).Value
        Cells(i, 6).Activate
         
            If InStr(data, "Doplatok mzdy") > 0 Then
                naklad = "Vyplata"
                Cells(i, 4).Value = naklad
            ElseIf InStr(data, "Sporiaci") > 0 Then
                naklad = "Sporenie"
                Cells(i, 4).Value = naklad
            ElseIf InStr(data, "STAVEBNA") > 0 Then
                naklad = "Sporenie"
                Cells(i, 4).Value = naklad
            ElseIf InStr(data, "SPL.UVERU") > 0 Then
                naklad = "Splatky"
                Cells(i, 4).Value = naklad
            ElseIf InStr(data, "SPL.UR") > 0 Then
                naklad = "Splatky"
                Cells(i, 4).Value = naklad
            ElseIf InStr(data, "KR.KAR") > 0 Then
                naklad = "Splatky"
                Cells(i, 4).Value = naklad
            ElseIf InStr(data, "ATM") > 0 Then
                naklad = "Vybery"
                Cells(i, 4).Value = naklad

            End If
        
        For Each Item In byvanie
            Search = InStr(data, Item)
            If Search > 0 Then
                naklad = "Byvanie"
                Cells(i, 4).Value = naklad
        Exit For
            End If
        Next
        
        For Each Item In potraviny
            Search = InStr(data, Item)
            If Search > 0 Then
                naklad = "Potraviny"
                Cells(i, 4).Value = naklad
        Exit For
            End If
        Next
        
'        For Each Item In drogeria
'            Search = InStr(data, Item)
'            If Search > 0 Then
'                naklad = "Drogeria"
'                Cells(i, 4).Value = naklad
'        Exit For
'            End If
'        Next

        For Each Item In doprava
            Search = InStr(data, Item)
            If Search > 0 Then
                naklad = "Doprava"
                Cells(i, 4).Value = naklad
        Exit For
            End If
        Next

        For Each Item In volny
            Search = InStr(data, Item)
            If Search > 0 Then
                naklad = "Volny cas"
                Cells(i, 4).Value = naklad
        Exit For
            End If
        Next

        For Each Item In oblecenie
            Search = InStr(data, Item)
            If Search > 0 Then
                naklad = "Oblecenie"
                Cells(i, 4).Value = naklad
        Exit For
            End If
        Next

        For Each Item In ostKD
            Search = InStr(data, Item)
            If Search > 0 Then
                naklad = "Ostatne KD"
                Cells(i, 4).Value = naklad
        Exit For
            End If
        Next

        For Each Item In ostDD
            Search = InStr(data, Item)
            If Search > 0 Then
                naklad = "Ostatne DD"
                Cells(i, 4).Value = naklad
        Exit For
            End If
        Next
        
        naklad = ""     'druh nakladu
    Next


End Sub



Sub XExpenseTotals()


'write Header for Expense Report
Cells(1, 8) = "Expense Report"
Cells(2, 8) = "Naklad"
Cells(2, 9) = "Suma"
Columns("H:I").ColumnWidth = 15

Cells(3, 8) = "Byvanie"     'sumif
Cells(4, 8) = "Potraviny"   'sumif
Cells(5, 8) = "Doprava"     'sumif
Cells(6, 8) = "Ostatne KD"  'sumif
Cells(7, 8) = "Volny cas"   'sumif
Cells(8, 8) = "Subtotal KD" 'calc sum
Cells(9, 8) = "Oblecenie"   'sumif
Cells(10, 8) = "Ostatne DD" 'sumif
Cells(11, 8) = "Subtotal DD"    'calc sum
Cells(12, 8) = "Total Expense"  'calc sum

Cells(14, 8) = "Vyplata"    'sumif
Cells(15, 8) = "Sporenie"   'sumif
Cells(16, 8) = "Splatky"    'sumif
Cells(17, 8) = "Vybery"     'sumif
Cells(18, 8) = "Prevody"    'sumif
Cells(19, 8) = "Cash flow"  'calc sum

'calculate sumif via expense types
With ActiveSheet
        .Range("I3").Value = Application.SumIf(.Range("D2:D100"), _
                              "*" & .Range("H3").Value & "*", _
                            .Range("E2:E100"))
    End With

With ActiveSheet
        .Range("I4").Value = Application.SumIf(.Range("D2:D100"), _
                              "*" & .Range("H4").Value & "*", _
                            .Range("E2:E100"))
    End With

With ActiveSheet
        .Range("I5").Value = Application.SumIf(.Range("D2:D100"), _
                              "*" & .Range("H5").Value & "*", _
                            .Range("E2:E100"))
    End With

With ActiveSheet
        .Range("I6").Value = Application.SumIf(.Range("D2:D100"), _
                              "*" & .Range("H6").Value & "*", _
                            .Range("E2:E100"))
    End With

With ActiveSheet
        .Range("I7").Value = Application.SumIf(.Range("D2:D100"), _
                              "*" & .Range("H7").Value & "*", _
                            .Range("E2:E100"))
    End With

With ActiveSheet
        .Range("I9").Value = Application.SumIf(.Range("D2:D100"), _
                              "*" & .Range("H9").Value & "*", _
                            .Range("E2:E100"))
    End With

With ActiveSheet
        .Range("I10").Value = Application.SumIf(.Range("D2:D100"), _
                              "*" & .Range("H10").Value & "*", _
                            .Range("E2:E100"))
    End With

With ActiveSheet
        .Range("I14").Value = Application.SumIf(.Range("D2:D100"), _
                              "*" & .Range("H14").Value & "*", _
                            .Range("E2:E100"))
    End With
    
With ActiveSheet
        .Range("I15").Value = Application.SumIf(.Range("D2:D100"), _
                              "*" & .Range("H15").Value & "*", _
                            .Range("E2:E100"))
    End With

With ActiveSheet
        .Range("I16").Value = Application.SumIf(.Range("D2:D100"), _
                              "*" & .Range("H16").Value & "*", _
                            .Range("E2:E100"))
    End With

With ActiveSheet
        .Range("I17").Value = Application.SumIf(.Range("D2:D100"), _
                              "*" & .Range("H17").Value & "*", _
                            .Range("E2:E100"))
    End With

With ActiveSheet
        .Range("I18").Value = Application.SumIf(.Range("D2:D100"), _
                              "*" & .Range("H18").Value & "*", _
                            .Range("E2:E100"))
    End With

Range("I3").Value = WorksheetFunction.Round(Range("I3"), 0)
Range("I4").Value = WorksheetFunction.Round(Range("I4"), 0)
Range("I5").Value = WorksheetFunction.Round(Range("I5"), 0)
Range("I6").Value = WorksheetFunction.Round(Range("I6"), 0)
Range("I7").Value = WorksheetFunction.Round(Range("I7"), 0)

'calculate subtotals and total
Cells(8, 9).Value = Application.Sum(Range("I3:I7"))
Cells(11, 9).Value = Application.Sum(Range("I9:I10"))
Cells(12, 9).Value = Application.Sum(Range("I8") + Range("I11"))

Cells(19, 9).Value = Range("I14") + Range("I12") + Range("I15") + Range("I16") + Range("I17") + Range("I18")

'format header and totals to bold font
Cells(1, 8).Font.Bold = True
Cells(2, 8).Font.Bold = True
Cells(2, 9).Font.Bold = True
Cells(8, 8).Font.Bold = True
Cells(8, 9).Font.Bold = True
Cells(11, 8).Font.Bold = True
Cells(11, 9).Font.Bold = True
Cells(12, 8).Font.Bold = True
Cells(12, 9).Font.Bold = True
Cells(19, 8).Font.Bold = True
Cells(19, 9).Font.Bold = True

 
End Sub
  
