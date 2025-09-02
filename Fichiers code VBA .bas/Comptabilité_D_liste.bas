Attribute VB_Name = "Comptabilité_D_liste"


Sub Final_Comptabilité_D()


Application.ScreenUpdating = False
Application.StatusBar = "D"

Sheets.Add.Name = "D"
    Cells.Select
    With Selection.Font
        .Name = "Times New Roman"
        .Size = 10
    End With
    With ActiveSheet.PageSetup
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.25)
        .BottomMargin = Application.InchesToPoints(0.25)
        .CenterHorizontally = True
        .CenterVertically = True
        .Order = xlDownThenOver
        .Zoom = 95
    End With
ActiveWindow.View = xlPageLayoutView
    Call Complément_Comptabilité_D
Range("N7,N75,N143,N211,N279,N347,N415,N483,N551,N619,N687,N755,N823,N891,N959,N1027,N1095,N1163,N1231,N1299,N1367,N1435,N1503,N1571,N1639") = "Janvier"
Range("N1707,N1775,N1843,N1911,N1979,N2047,N2115,N2183,N2251,N2319,N2387,N2455,N2523,N2591,N2659,N2727,N2795,N2863,N2931,N2999,N3067") = "Janvier"
Columns("A:AA").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Comptabilité_D
Range("N7,N75,N143,N211,N279,N347,N415,N483,N551,N619,N687,N755,N823,N891,N959,N1027,N1095,N1163,N1231,N1299,N1367,N1435,N1503,N1571,N1639") = "Février"
Range("N1707,N1775,N1843,N1911,N1979,N2047,N2115,N2183,N2251,N2319,N2387,N2455,N2523,N2591,N2659,N2727,N2795,N2863,N2931,N2999,N3067") = "Février"
Columns("A:AA").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Comptabilité_D
Range("N7,N75,N143,N211,N279,N347,N415,N483,N551,N619,N687,N755,N823,N891,N959,N1027,N1095,N1163,N1231,N1299,N1367,N1435,N1503,N1571,N1639") = "Mars"
Range("N1707,N1775,N1843,N1911,N1979,N2047,N2115,N2183,N2251,N2319,N2387,N2455,N2523,N2591,N2659,N2727,N2795,N2863,N2931,N2999,N3067") = "Mars"
Columns("A:AA").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Comptabilité_D
Range("N7,N75,N143,N211,N279,N347,N415,N483,N551,N619,N687,N755,N823,N891,N959,N1027,N1095,N1163,N1231,N1299,N1367,N1435,N1503,N1571,N1639") = "Avril"
Range("N1707,N1775,N1843,N1911,N1979,N2047,N2115,N2183,N2251,N2319,N2387,N2455,N2523,N2591,N2659,N2727,N2795,N2863,N2931,N2999,N3067") = "Avril"
Columns("A:AA").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Comptabilité_D
Range("N7,N75,N143,N211,N279,N347,N415,N483,N551,N619,N687,N755,N823,N891,N959,N1027,N1095,N1163,N1231,N1299,N1367,N1435,N1503,N1571,N1639") = "Mai"
Range("N1707,N1775,N1843,N1911,N1979,N2047,N2115,N2183,N2251,N2319,N2387,N2455,N2523,N2591,N2659,N2727,N2795,N2863,N2931,N2999,N3067") = "Mai"
Columns("A:AA").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Comptabilité_D
Range("N7,N75,N143,N211,N279,N347,N415,N483,N551,N619,N687,N755,N823,N891,N959,N1027,N1095,N1163,N1231,N1299,N1367,N1435,N1503,N1571,N1639") = "Juin"
Range("N1707,N1775,N1843,N1911,N1979,N2047,N2115,N2183,N2251,N2319,N2387,N2455,N2523,N2591,N2659,N2727,N2795,N2863,N2931,N2999,N3067") = "Juin"
Columns("A:AA").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Comptabilité_D
Range("N7,N75,N143,N211,N279,N347,N415,N483,N551,N619,N687,N755,N823,N891,N959,N1027,N1095,N1163,N1231,N1299,N1367,N1435,N1503,N1571,N1639") = "Juillet"
Range("N1707,N1775,N1843,N1911,N1979,N2047,N2115,N2183,N2251,N2319,N2387,N2455,N2523,N2591,N2659,N2727,N2795,N2863,N2931,N2999,N3067") = "Juillet"
Columns("A:AA").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Comptabilité_D
Range("N7,N75,N143,N211,N279,N347,N415,N483,N551,N619,N687,N755,N823,N891,N959,N1027,N1095,N1163,N1231,N1299,N1367,N1435,N1503,N1571,N1639") = "Août"
Range("N1707,N1775,N1843,N1911,N1979,N2047,N2115,N2183,N2251,N2319,N2387,N2455,N2523,N2591,N2659,N2727,N2795,N2863,N2931,N2999,N3067") = "Août"
Columns("A:S").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Comptabilité_D
Range("N7,N75,N143,N211,N279,N347,N415,N483,N551,N619,N687,N755,N823,N891,N959,N1027,N1095,N1163,N1231,N1299,N1367,N1435,N1503,N1571,N1639") = "Septembre"
Range("N1707,N1775,N1843,N1911,N1979,N2047,N2115,N2183,N2251,N2319,N2387,N2455,N2523,N2591,N2659,N2727,N2795,N2863,N2931,N2999,N3067") = "Septembre"
Columns("A:S").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Comptabilité_D
Range("N7,N75,N143,N211,N279,N347,N415,N483,N551,N619,N687,N755,N823,N891,N959,N1027,N1095,N1163,N1231,N1299,N1367,N1435,N1503,N1571,N1639") = "Octobre"
Range("N1707,N1775,N1843,N1911,N1979,N2047,N2115,N2183,N2251,N2319,N2387,N2455,N2523,N2591,N2659,N2727,N2795,N2863,N2931,N2999,N3067") = "Octobre"
Columns("A:S").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Comptabilité_D
Range("N7,N75,N143,N211,N279,N347,N415,N483,N551,N619,N687,N755,N823,N891,N959,N1027,N1095,N1163,N1231,N1299,N1367,N1435,N1503,N1571,N1639") = "Novembre"
Range("N1707,N1775,N1843,N1911,N1979,N2047,N2115,N2183,N2251,N2319,N2387,N2455,N2523,N2591,N2659,N2727,N2795,N2863,N2931,N2999,N3067") = "Novembre"
Columns("A:S").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Comptabilité_D
Range("N7,N75,N143,N211,N279,N347,N415,N483,N551,N619,N687,N755,N823,N891,N959,N1027,N1095,N1163,N1231,N1299,N1367,N1435,N1503,N1571,N1639") = "Décembre"
Range("N1707,N1775,N1843,N1911,N1979,N2047,N2115,N2183,N2251,N2319,N2387,N2455,N2523,N2591,N2659,N2727,N2795,N2863,N2931,N2999,N3067") = "Décembre"
Rows("69").Select
    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
Rows("137").Select
    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
Rows("205").Select
    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
Rows("273").Select
    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
Rows("341").Select
    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
Rows("409").Select
    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
Rows("477").Select
    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
Rows("545").Select
    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
Rows("613").Select
    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
Rows("681").Select
    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
Rows("749").Select
    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
Rows("817").Select
    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
Rows("885").Select
    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
Rows("953").Select
    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
Rows("1021").Select
    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
Rows("1089").Select
    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
Rows("1157").Select
    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
Rows("1225").Select
    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
Rows("1293").Select
    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
Rows("1361").Select
    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
Rows("1429").Select
    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
Rows("1497").Select
    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
Rows("1565").Select
    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
Rows("1633").Select
    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
Rows("1701").Select
    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
Rows("1769").Select
    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
Rows("1837").Select
    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
Rows("1905").Select
    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
Rows("1973").Select
    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
Rows("2041").Select
    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
Rows("2109").Select
    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
Rows("2177").Select
    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
Rows("2245").Select
    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
Rows("2313").Select
    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
Rows("2381").Select
    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
Rows("2449").Select
    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
Rows("2517").Select
    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
Rows("2585").Select
    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
Rows("2653").Select
    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
Rows("2721").Select
    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
Rows("2789").Select
    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
Rows("2857").Select
    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
Rows("2925").Select
    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
Rows("2993").Select
    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
Rows("3061").Select
    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
Rows("3129").Select
    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell

Application.StatusBar = False
    
    
End Sub


Sub Complément_Comptabilité_D()
    
    
    Call Mise_en_page_Comptabilité_D
Range("A1").Select
    Call Fiche_Comptabilité_D
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("T12").Copy Workbooks("Comptabilité.xlsx").Worksheets("D").Range("F7")
Range("A69").Select
    Call Fiche_Comptabilité_D
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("T13").Copy Workbooks("Comptabilité.xlsx").Worksheets("D").Range("F75")
Range("A137").Select
    Call Fiche_Comptabilité_D
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("T14").Copy Workbooks("Comptabilité.xlsx").Worksheets("D").Range("F143")
Range("A205").Select
    Call Fiche_Comptabilité_D
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("T15").Copy Workbooks("Comptabilité.xlsx").Worksheets("D").Range("F211")
Range("A273").Select
    Call Fiche_Comptabilité_D
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("T16").Copy Workbooks("Comptabilité.xlsx").Worksheets("D").Range("F279")
Range("A341").Select
    Call Fiche_Comptabilité_D
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("T17").Copy Workbooks("Comptabilité.xlsx").Worksheets("D").Range("F347")
Range("A409").Select
    Call Fiche_Comptabilité_D
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("T18").Copy Workbooks("Comptabilité.xlsx").Worksheets("D").Range("F415")
Range("A477").Select
    Call Fiche_Comptabilité_D
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("T19").Copy Workbooks("Comptabilité.xlsx").Worksheets("D").Range("F483")
Range("A545").Select
    Call Fiche_Comptabilité_D
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("T20").Copy Workbooks("Comptabilité.xlsx").Worksheets("D").Range("F551")
Range("A613").Select
    Call Fiche_Comptabilité_D
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("T21").Copy Workbooks("Comptabilité.xlsx").Worksheets("D").Range("F619")
Range("A681").Select
    Call Fiche_Comptabilité_D
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("T22").Copy Workbooks("Comptabilité.xlsx").Worksheets("D").Range("F687")
Range("A749").Select
    Call Fiche_Comptabilité_D
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("T23").Copy Workbooks("Comptabilité.xlsx").Worksheets("D").Range("F755")
Range("A817").Select
    Call Fiche_Comptabilité_D
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("T24").Copy Workbooks("Comptabilité.xlsx").Worksheets("D").Range("F823")
Range("A885").Select
    Call Fiche_Comptabilité_D
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("T25").Copy Workbooks("Comptabilité.xlsx").Worksheets("D").Range("F891")
Range("A953").Select
    Call Fiche_Comptabilité_D
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("T26").Copy Workbooks("Comptabilité.xlsx").Worksheets("D").Range("F959")
Range("A1021").Select
    Call Fiche_Comptabilité_D
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("T27").Copy Workbooks("Comptabilité.xlsx").Worksheets("D").Range("F1027")
Range("A1089").Select
    Call Fiche_Comptabilité_D
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("T28").Copy Workbooks("Comptabilité.xlsx").Worksheets("D").Range("F1095")
Range("A1157").Select
    Call Fiche_Comptabilité_D
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("T29").Copy Workbooks("Comptabilité.xlsx").Worksheets("D").Range("F1163")
Range("A1225").Select
    Call Fiche_Comptabilité_D
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("T30").Copy Workbooks("Comptabilité.xlsx").Worksheets("D").Range("F1231")
Range("A1293").Select
    Call Fiche_Comptabilité_D
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("T31").Copy Workbooks("Comptabilité.xlsx").Worksheets("D").Range("F1299")
Range("A1361").Select
    Call Fiche_Comptabilité_D
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("T32").Copy Workbooks("Comptabilité.xlsx").Worksheets("D").Range("F1367")
Range("A1429").Select
    Call Fiche_Comptabilité_D
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("T33").Copy Workbooks("Comptabilité.xlsx").Worksheets("D").Range("F1435")
Range("A1497").Select
    Call Fiche_Comptabilité_D
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("T34").Copy Workbooks("Comptabilité.xlsx").Worksheets("D").Range("F1503")
Range("A1565").Select
    Call Fiche_Comptabilité_D
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("T35").Copy Workbooks("Comptabilité.xlsx").Worksheets("D").Range("F1571")
Range("A1633").Select
    Call Fiche_Comptabilité_D
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("T36").Copy Workbooks("Comptabilité.xlsx").Worksheets("D").Range("F1639")
Range("A1701").Select
    Call Fiche_Comptabilité_D
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("T37").Copy Workbooks("Comptabilité.xlsx").Worksheets("D").Range("F1707")
Range("A1769").Select
    Call Fiche_Comptabilité_D
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("T38").Copy Workbooks("Comptabilité.xlsx").Worksheets("D").Range("F1775")
Range("A1837").Select
    Call Fiche_Comptabilité_D
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("T39").Copy Workbooks("Comptabilité.xlsx").Worksheets("D").Range("F1843")
Range("A1905").Select
    Call Fiche_Comptabilité_D
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("T40").Copy Workbooks("Comptabilité.xlsx").Worksheets("D").Range("F1911")
Range("A1973").Select
    Call Fiche_Comptabilité_D
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("T41").Copy Workbooks("Comptabilité.xlsx").Worksheets("D").Range("F1979")
Range("A2041").Select
    Call Fiche_Comptabilité_D
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("T42").Copy Workbooks("Comptabilité.xlsx").Worksheets("D").Range("F2047")
Range("A2109").Select
    Call Fiche_Comptabilité_D
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("T43").Copy Workbooks("Comptabilité.xlsx").Worksheets("D").Range("F2115")
Range("A2177").Select
    Call Fiche_Comptabilité_D
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("T44").Copy Workbooks("Comptabilité.xlsx").Worksheets("D").Range("F2183")
Range("A2245").Select
    Call Fiche_Comptabilité_D
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("T45").Copy Workbooks("Comptabilité.xlsx").Worksheets("D").Range("F2251")
Range("A2313").Select
    Call Fiche_Comptabilité_D
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("T46").Copy Workbooks("Comptabilité.xlsx").Worksheets("D").Range("F2319")
Range("A2381").Select
    Call Fiche_Comptabilité_D
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("T47").Copy Workbooks("Comptabilité.xlsx").Worksheets("D").Range("F2387")
Range("A2449").Select
    Call Fiche_Comptabilité_D
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("T48").Copy Workbooks("Comptabilité.xlsx").Worksheets("D").Range("F2455")
Range("A2517").Select
    Call Fiche_Comptabilité_D
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("T49").Copy Workbooks("Comptabilité.xlsx").Worksheets("D").Range("F2523")
Range("A2585").Select
    Call Fiche_Comptabilité_D
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("T50").Copy Workbooks("Comptabilité.xlsx").Worksheets("D").Range("F2591")
Range("A2653").Select
    Call Fiche_Comptabilité_D
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("T51").Copy Workbooks("Comptabilité.xlsx").Worksheets("D").Range("F2659")
Range("A2721").Select
    Call Fiche_Comptabilité_D
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("T52").Copy Workbooks("Comptabilité.xlsx").Worksheets("D").Range("F2727")
Range("A2789").Select
    Call Fiche_Comptabilité_D
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("T53").Copy Workbooks("Comptabilité.xlsx").Worksheets("D").Range("F2795")
Range("A2857").Select
    Call Fiche_Comptabilité_D
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("T54").Copy Workbooks("Comptabilité.xlsx").Worksheets("D").Range("F2863")
Range("A2925").Select
    Call Fiche_Comptabilité_D
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("T55").Copy Workbooks("Comptabilité.xlsx").Worksheets("D").Range("F2931")
Range("A2993").Select
    Call Fiche_Comptabilité_D
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("T56").Copy Workbooks("Comptabilité.xlsx").Worksheets("D").Range("F2999")
Range("A3061").Select
    Call Fiche_Comptabilité_D
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("T57").Copy Workbooks("Comptabilité.xlsx").Worksheets("D").Range("F3067")
Columns("AB").Select
    ActiveWindow.SelectedSheets.VPageBreaks.Add Before:=ActiveCell


End Sub



Sub Fiche_Comptabilité_D()


ActiveCell.Offset(0, 0).Range("A1:AA1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Comptabilité"
ActiveCell.Offset(1, 0).Range("A1").Select
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
ActiveCell.Offset(0, 25).Range("A1").Select
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
ActiveCell.Offset(1, -24).Range("A1:Z1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    ActiveCell.FormulaR1C1 = "D"
ActiveCell.Offset(2, -1).Range("A1:K1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Compte"
ActiveCell.Offset(0, 2).Range("A1:G1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Mois"
ActiveCell.Offset(0, 2).Range("A1:G1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Année"
ActiveCell.Offset(1, -19).Range("A1").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
ActiveCell.Offset(0, 8).Range("A1").Select
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
ActiveCell.Offset(0, 3).Range("A1").Select
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
ActiveCell.Offset(0, 5).Range("A1").Select
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
ActiveCell.Offset(0, 3).Range("A1").Select
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
ActiveCell.Offset(0, 5).Range("A1").Select
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
ActiveCell.Offset(1, -12).Range("A1:E1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
ActiveCell.Offset(0, 4).Range("A1:E1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.NumberFormat = "yyyy"
    ActiveCell.FormulaR1C1 = Date
ActiveCell.Offset(3, -21).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Facture"
ActiveCell.Offset(0, 2).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Info"
ActiveCell.Offset(0, 2).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Date"
ActiveCell.Offset(0, 2).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Vente march."
ActiveCell.Offset(0, 2).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Export march."
ActiveCell.Offset(0, 2).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Déduction"
ActiveCell.Offset(0, 2).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "TVA"
ActiveCell.Offset(0, -24).Range("A1:AA1").Select
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
ActiveCell.Offset(2, 11).Range("A1:A53").Select
    Selection.NumberFormat = "#,##0.00"
ActiveCell.Offset(0, 4).Range("A1:A53").Select
    Selection.NumberFormat = "#,##0.00"
ActiveCell.Offset(0, 4).Range("A1:A53").Select
    Selection.NumberFormat = "#,##0.00"
ActiveCell.Offset(0, 4).Range("A1:A53").Select
    Selection.NumberFormat = "#,##0.00"
ActiveCell.Offset(54, -25).Range("A1:K1").Select
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
ActiveCell.Offset(0, 12).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Vente March."
ActiveCell.Offset(0, 2).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Export march."
ActiveCell.Offset(0, 2).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Déduction"
ActiveCell.Offset(0, 2).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "TVA"
ActiveCell.Offset(1, -11).Range("A1").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
ActiveCell.Offset(0, 4).Range("A1").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
ActiveCell.Offset(0, 4).Range("A1").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
ActiveCell.Offset(0, 4).Range("A1").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
ActiveCell.Offset(1, -12).Range("A1").Select
    Selection.NumberFormat = "#,##0.00"
    ActiveCell.FormulaR1C1 = "=SUM(R[-56]C:R[-4]C)"
ActiveCell.Offset(0, 4).Range("A1").Select
    Selection.NumberFormat = "#,##0.00"
    ActiveCell.FormulaR1C1 = "=SUM(R[-56]C:R[-4]C)"
ActiveCell.Offset(0, 4).Range("A1").Select
    Selection.NumberFormat = "#,##0.00"
    ActiveCell.FormulaR1C1 = "=SUM(R[-56]C:R[-4]C)"
ActiveCell.Offset(0, 4).Range("A1").Select
    Selection.NumberFormat = "#,##0.00"
    ActiveCell.FormulaR1C1 = "=SUM(R[-56]C:R[-4]C)"


End Sub


Sub Mise_en_page_Comptabilité_D()


ActiveCell.Columns("A:A").EntireColumn.ColumnWidth = 0.5
ActiveCell.Offset(0, 1).Columns("A:A").EntireColumn.ColumnWidth = 8
ActiveCell.Offset(0, 2).Columns("A:A").EntireColumn.ColumnWidth = 0.5
ActiveCell.Offset(0, 3).Columns("A:A").EntireColumn.ColumnWidth = 1
ActiveCell.Offset(0, 4).Columns("A:A").EntireColumn.ColumnWidth = 0.5
ActiveCell.Offset(0, 5).Columns("A:A").EntireColumn.ColumnWidth = 4
ActiveCell.Offset(0, 6).Columns("A:A").EntireColumn.ColumnWidth = 0.5
ActiveCell.Offset(0, 7).Columns("A:A").EntireColumn.ColumnWidth = 1
ActiveCell.Offset(0, 8).Columns("A:A").EntireColumn.ColumnWidth = 0.5
ActiveCell.Offset(0, 9).Columns("A:A").EntireColumn.ColumnWidth = 10.67
ActiveCell.Offset(0, 10).Columns("A:A").EntireColumn.ColumnWidth = 0.5
ActiveCell.Offset(0, 11).Columns("A:A").EntireColumn.ColumnWidth = 1
ActiveCell.Offset(0, 12).Columns("A:A").EntireColumn.ColumnWidth = 0.5
ActiveCell.Offset(0, 13).Columns("A:A").EntireColumn.ColumnWidth = 10.67
ActiveCell.Offset(0, 14).Columns("A:A").EntireColumn.ColumnWidth = 0.5
ActiveCell.Offset(0, 15).Columns("A:A").EntireColumn.ColumnWidth = 1
ActiveCell.Offset(0, 16).Columns("A:A").EntireColumn.ColumnWidth = 0.5
ActiveCell.Offset(0, 17).Columns("A:A").EntireColumn.ColumnWidth = 10.67
ActiveCell.Offset(0, 18).Columns("A:A").EntireColumn.ColumnWidth = 0.5
ActiveCell.Offset(0, 19).Columns("A:A").EntireColumn.ColumnWidth = 1
ActiveCell.Offset(0, 20).Columns("A:A").EntireColumn.ColumnWidth = 0.5
ActiveCell.Offset(0, 21).Columns("A:A").EntireColumn.ColumnWidth = 10.67
ActiveCell.Offset(0, 22).Columns("A:A").EntireColumn.ColumnWidth = 0.5
ActiveCell.Offset(0, 23).Columns("A:A").EntireColumn.ColumnWidth = 1
ActiveCell.Offset(0, 24).Columns("A:A").EntireColumn.ColumnWidth = 0.5
ActiveCell.Offset(0, 25).Columns("A:A").EntireColumn.ColumnWidth = 10.67
ActiveCell.Offset(0, 26).Columns("A:A").EntireColumn.ColumnWidth = 0.5


End Sub



