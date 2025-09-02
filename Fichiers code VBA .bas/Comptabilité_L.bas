Attribute VB_Name = "Comptabilité_L"


Sub Final_Comptabilité_L()


Application.ScreenUpdating = False
Application.StatusBar = "Feuil L"

Sheets.Add.Name = "L"
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
    Call Complément_Comptabilité_L
Range("J7,J75,J143,J211,J279,J347,J415,J483,J551,J619,J687,J755,J823,J891,J959,J1027,J1095,J1163,J1231,J1299,J1367,J1435,J1503,J1571,J1639") = "Janvier"
Range("J1707,J1775,J1843,J1911,J1979,J2047,J2115,J2183,J2251,J2319,J2387,J2455,J2523,J2591,J2659,J2727,J2795,J2863,J2931,J2999,J3067") = "Janvier"
Columns("A:S").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Comptabilité_L
Range("J7,J75,J143,J211,J279,J347,J415,J483,J551,J619,J687,J755,J823,J891,J959,J1027,J1095,J1163,J1231,J1299,J1367,J1435,J1503,J1571,J1639") = "Février"
Range("J1707,J1775,J1843,J1911,J1979,J2047,J2115,J2183,J2251,J2319,J2387,J2455,J2523,J2591,J2659,J2727,J2795,J2863,J2931,J2999,J3067") = "Février"
Columns("A:S").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Comptabilité_L
Range("J7,J75,J143,J211,J279,J347,J415,J483,J551,J619,J687,J755,J823,J891,J959,J1027,J1095,J1163,J1231,J1299,J1367,J1435,J1503,J1571,J1639") = "Mars"
Range("J1707,J1775,J1843,J1911,J1979,J2047,J2115,J2183,J2251,J2319,J2387,J2455,J2523,J2591,J2659,J2727,J2795,J2863,J2931,J2999,J3067") = "Mars"
Columns("A:S").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Comptabilité_L
Range("J7,J75,J143,J211,J279,J347,J415,J483,J551,J619,J687,J755,J823,J891,J959,J1027,J1095,J1163,J1231,J1299,J1367,J1435,J1503,J1571,J1639") = "Avril"
Range("J1707,J1775,J1843,J1911,J1979,J2047,J2115,J2183,J2251,J2319,J2387,J2455,J2523,J2591,J2659,J2727,J2795,J2863,J2931,J2999,J3067") = "Avril"
Columns("A:S").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Comptabilité_L
Range("J7,J75,J143,J211,J279,J347,J415,J483,J551,J619,J687,J755,J823,J891,J959,J1027,J1095,J1163,J1231,J1299,J1367,J1435,J1503,J1571,J1639") = "Mai"
Range("J1707,J1775,J1843,J1911,J1979,J2047,J2115,J2183,J2251,J2319,J2387,J2455,J2523,J2591,J2659,J2727,J2795,J2863,J2931,J2999,J3067") = "Mai"
Columns("A:S").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Comptabilité_L
Range("J7,J75,J143,J211,J279,J347,J415,J483,J551,J619,J687,J755,J823,J891,J959,J1027,J1095,J1163,J1231,J1299,J1367,J1435,J1503,J1571,J1639") = "Juin"
Range("J1707,J1775,J1843,J1911,J1979,J2047,J2115,J2183,J2251,J2319,J2387,J2455,J2523,J2591,J2659,J2727,J2795,J2863,J2931,J2999,J3067") = "Juin"
Columns("A:S").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Comptabilité_L
Range("J7,J75,J143,J211,J279,J347,J415,J483,J551,J619,J687,J755,J823,J891,J959,J1027,J1095,J1163,J1231,J1299,J1367,J1435,J1503,J1571,J1639") = "Juillet"
Range("J1707,J1775,J1843,J1911,J1979,J2047,J2115,J2183,J2251,J2319,J2387,J2455,J2523,J2591,J2659,J2727,J2795,J2863,J2931,J2999,J3067") = "Juillet"
Columns("A:S").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Comptabilité_L
Range("J7,J75,J143,J211,J279,J347,J415,J483,J551,J619,J687,J755,J823,J891,J959,J1027,J1095,J1163,J1231,J1299,J1367,J1435,J1503,J1571,J1639") = "Août"
Range("J1707,J1775,J1843,J1911,J1979,J2047,J2115,J2183,J2251,J2319,J2387,J2455,J2523,J2591,J2659,J2727,J2795,J2863,J2931,J2999,J3067") = "Août"
Columns("A:S").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Comptabilité_L
Range("J7,J75,J143,J211,J279,J347,J415,J483,J551,J619,J687,J755,J823,J891,J959,J1027,J1095,J1163,J1231,J1299,J1367,J1435,J1503,J1571,J1639") = "Septembre"
Range("J1707,J1775,J1843,J1911,J1979,J2047,J2115,J2183,J2251,J2319,J2387,J2455,J2523,J2591,J2659,J2727,J2795,J2863,J2931,J2999,J3067") = "Septembre"
Columns("A:S").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Comptabilité_L
Range("J7,J75,J143,J211,J279,J347,J415,J483,J551,J619,J687,J755,J823,J891,J959,J1027,J1095,J1163,J1231,J1299,J1367,J1435,J1503,J1571,J1639") = "Octobre"
Range("J1707,J1775,J1843,J1911,J1979,J2047,J2115,J2183,J2251,J2319,J2387,J2455,J2523,J2591,J2659,J2727,J2795,J2863,J2931,J2999,J3067") = "Octobre"
Columns("A:S").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Comptabilité_L
Range("J7,J75,J143,J211,J279,J347,J415,J483,J551,J619,J687,J755,J823,J891,J959,J1027,J1095,J1163,J1231,J1299,J1367,J1435,J1503,J1571,J1639") = "Novembre"
Range("J1707,J1775,J1843,J1911,J1979,J2047,J2115,J2183,J2251,J2319,J2387,J2455,J2523,J2591,J2659,J2727,J2795,J2863,J2931,J2999,J3067") = "Novembre"
Columns("A:S").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Comptabilité_L
Range("J7,J75,J143,J211,J279,J347,J415,J483,J551,J619,J687,J755,J823,J891,J959,J1027,J1095,J1163,J1231,J1299,J1367,J1435,J1503,J1571,J1639") = "Décembre"
Range("J1707,J1775,J1843,J1911,J1979,J2047,J2115,J2183,J2251,J2319,J2387,J2455,J2523,J2591,J2659,J2727,J2795,J2863,J2931,J2999,J3067") = "Décembre"
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
 
    
Sub Complément_Comptabilité_L()
    
    
    Call Mise_en_page_Comptabilité_L
Range("A1").Select
    Call Fiche_Comptabilité_L
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BS12").Copy Workbooks("Comptabilité.xlsx").Worksheets("L").Range("B7")
Range("A69").Select
    Call Fiche_Comptabilité_L
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BS13").Copy Workbooks("Comptabilité.xlsx").Worksheets("L").Range("B75")
Range("A137").Select
    Call Fiche_Comptabilité_L
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BS14").Copy Workbooks("Comptabilité.xlsx").Worksheets("L").Range("B143")
Range("A205").Select
    Call Fiche_Comptabilité_L
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BS15").Copy Workbooks("Comptabilité.xlsx").Worksheets("L").Range("B211")
Range("A273").Select
    Call Fiche_Comptabilité_L
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BS16").Copy Workbooks("Comptabilité.xlsx").Worksheets("L").Range("B279")
Range("A341").Select
    Call Fiche_Comptabilité_L
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BS17").Copy Workbooks("Comptabilité.xlsx").Worksheets("L").Range("B347")
Range("A409").Select
    Call Fiche_Comptabilité_L
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BS18").Copy Workbooks("Comptabilité.xlsx").Worksheets("L").Range("B415")
Range("A477").Select
    Call Fiche_Comptabilité_L
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BS19").Copy Workbooks("Comptabilité.xlsx").Worksheets("L").Range("B483")
Range("A545").Select
    Call Fiche_Comptabilité_L
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BS20").Copy Workbooks("Comptabilité.xlsx").Worksheets("L").Range("B551")
Range("A613").Select
    Call Fiche_Comptabilité_L
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BS21").Copy Workbooks("Comptabilité.xlsx").Worksheets("L").Range("B619")
Range("A681").Select
    Call Fiche_Comptabilité_L
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BS22").Copy Workbooks("Comptabilité.xlsx").Worksheets("L").Range("B687")
Range("A749").Select
    Call Fiche_Comptabilité_L
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BS23").Copy Workbooks("Comptabilité.xlsx").Worksheets("L").Range("B755")
Range("A817").Select
    Call Fiche_Comptabilité_L
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BS24").Copy Workbooks("Comptabilité.xlsx").Worksheets("L").Range("B823")
Range("A885").Select
    Call Fiche_Comptabilité_L
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BS25").Copy Workbooks("Comptabilité.xlsx").Worksheets("L").Range("B891")
Range("A953").Select
    Call Fiche_Comptabilité_L
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BS26").Copy Workbooks("Comptabilité.xlsx").Worksheets("L").Range("B959")
Range("A1021").Select
    Call Fiche_Comptabilité_L
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BS27").Copy Workbooks("Comptabilité.xlsx").Worksheets("L").Range("B1027")
Range("A1089").Select
    Call Fiche_Comptabilité_L
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BS28").Copy Workbooks("Comptabilité.xlsx").Worksheets("L").Range("B1095")
Range("A1157").Select
    Call Fiche_Comptabilité_L
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BS29").Copy Workbooks("Comptabilité.xlsx").Worksheets("L").Range("B1163")
Range("A1225").Select
    Call Fiche_Comptabilité_L
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BS30").Copy Workbooks("Comptabilité.xlsx").Worksheets("L").Range("B1231")
Range("A1293").Select
    Call Fiche_Comptabilité_L
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BS31").Copy Workbooks("Comptabilité.xlsx").Worksheets("L").Range("B1299")
Range("A1361").Select
    Call Fiche_Comptabilité_L
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BS32").Copy Workbooks("Comptabilité.xlsx").Worksheets("L").Range("B1367")
Range("A1429").Select
    Call Fiche_Comptabilité_L
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BS33").Copy Workbooks("Comptabilité.xlsx").Worksheets("L").Range("B1435")
Range("A1497").Select
    Call Fiche_Comptabilité_L
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BS34").Copy Workbooks("Comptabilité.xlsx").Worksheets("L").Range("B1503")
Range("A1565").Select
    Call Fiche_Comptabilité_L
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BS35").Copy Workbooks("Comptabilité.xlsx").Worksheets("L").Range("B1571")
Range("A1633").Select
    Call Fiche_Comptabilité_L
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BS36").Copy Workbooks("Comptabilité.xlsx").Worksheets("L").Range("B1639")
Range("A1701").Select
    Call Fiche_Comptabilité_L
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BS37").Copy Workbooks("Comptabilité.xlsx").Worksheets("L").Range("B1707")
Range("A1769").Select
    Call Fiche_Comptabilité_L
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BS38").Copy Workbooks("Comptabilité.xlsx").Worksheets("L").Range("B1775")
Range("A1837").Select
    Call Fiche_Comptabilité_L
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BS39").Copy Workbooks("Comptabilité.xlsx").Worksheets("L").Range("B1843")
Range("A1905").Select
    Call Fiche_Comptabilité_L
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BS40").Copy Workbooks("Comptabilité.xlsx").Worksheets("L").Range("B1911")
Range("A1973").Select
    Call Fiche_Comptabilité_L
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BS41").Copy Workbooks("Comptabilité.xlsx").Worksheets("L").Range("B1979")
Range("A2041").Select
    Call Fiche_Comptabilité_L
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BS42").Copy Workbooks("Comptabilité.xlsx").Worksheets("L").Range("B2047")
Range("A2109").Select
    Call Fiche_Comptabilité_L
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BS43").Copy Workbooks("Comptabilité.xlsx").Worksheets("L").Range("B2115")
Range("A2177").Select
    Call Fiche_Comptabilité_L
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BS44").Copy Workbooks("Comptabilité.xlsx").Worksheets("L").Range("B2183")
Range("A2245").Select
    Call Fiche_Comptabilité_L
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BS45").Copy Workbooks("Comptabilité.xlsx").Worksheets("L").Range("B2251")
Range("A2313").Select
    Call Fiche_Comptabilité_L
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BS46").Copy Workbooks("Comptabilité.xlsx").Worksheets("L").Range("B2319")
Range("A2381").Select
    Call Fiche_Comptabilité_L
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BS47").Copy Workbooks("Comptabilité.xlsx").Worksheets("L").Range("B2387")
Range("A2449").Select
    Call Fiche_Comptabilité_L
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BS48").Copy Workbooks("Comptabilité.xlsx").Worksheets("L").Range("B2455")
Range("A2517").Select
    Call Fiche_Comptabilité_L
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BS49").Copy Workbooks("Comptabilité.xlsx").Worksheets("L").Range("B2523")
Range("A2585").Select
    Call Fiche_Comptabilité_L
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BS50").Copy Workbooks("Comptabilité.xlsx").Worksheets("L").Range("B2591")
Range("A2653").Select
    Call Fiche_Comptabilité_L
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BS51").Copy Workbooks("Comptabilité.xlsx").Worksheets("L").Range("B2659")
Range("A2721").Select
    Call Fiche_Comptabilité_L
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BS52").Copy Workbooks("Comptabilité.xlsx").Worksheets("L").Range("B2727")
Range("A2789").Select
    Call Fiche_Comptabilité_L
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BS53").Copy Workbooks("Comptabilité.xlsx").Worksheets("L").Range("B2795")
Range("A2857").Select
    Call Fiche_Comptabilité_L
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BS54").Copy Workbooks("Comptabilité.xlsx").Worksheets("L").Range("B2863")
Range("A2925").Select
    Call Fiche_Comptabilité_L
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BS55").Copy Workbooks("Comptabilité.xlsx").Worksheets("L").Range("B2931")
Range("A2993").Select
    Call Fiche_Comptabilité_L
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BS56").Copy Workbooks("Comptabilité.xlsx").Worksheets("L").Range("B2999")
Range("A3061").Select
    Call Fiche_Comptabilité_L
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BS57").Copy Workbooks("Comptabilité.xlsx").Worksheets("L").Range("B3067")
Columns("T").Select
    ActiveWindow.SelectedSheets.VPageBreaks.Add Before:=ActiveCell


End Sub
    
    
Sub Fiche_Comptabilité_L()
    
    
ActiveCell.Offset(0, 0).Range("A1:S1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Comptabilité"
ActiveCell.Offset(1, 0).Range("A1").Select
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
ActiveCell.Offset(0, 17).Range("A1").Select
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
ActiveCell.Offset(1, -16).Range("A1:Q1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    ActiveCell.FormulaR1C1 = "L"
ActiveCell.Offset(2, -1).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Compte"
ActiveCell.Offset(0, 2).Range("A1:K1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Mois"
ActiveCell.Offset(0, 2).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Année"
ActiveCell.Offset(1, -15).Range("A1").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
ActiveCell.Offset(0, 3).Range("A1").Select
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
ActiveCell.Offset(0, 9).Range("A1").Select
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
ActiveCell.Offset(0, 3).Range("A1").Select
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
ActiveCell.Offset(0, 1).Range("A1").Select
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
ActiveCell.Offset(1, -16).Range("A1").Select
    Selection.HorizontalAlignment = xlCenter
ActiveCell.Offset(0, 8).Range("A1").Select
    Selection.HorizontalAlignment = xlCenter
ActiveCell.Offset(0, 8).Range("A1").Select
    Selection.HorizontalAlignment = xlCenter
    Selection.NumberFormat = "yyyy"
    ActiveCell.FormulaR1C1 = Date
ActiveCell.Offset(3, -13).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Date"
ActiveCell.Offset(0, 2).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Débit"
ActiveCell.Offset(0, 2).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Crédit"
ActiveCell.Offset(0, 2).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Solde"
ActiveCell.Offset(0, -16).Range("A1:S1").Select
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
ActiveCell.Offset(2, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Solde à nouveau"
ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Débit"
ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Credit"
ActiveCell.Offset(-2, 8).Range("A1:A53").Select
    Selection.NumberFormat = "#,##0.00"
ActiveCell.Offset(0, 4).Range("A1:A53").Select
    Selection.NumberFormat = "#,##0.00"
ActiveCell.Offset(0, 4).Range("A1:A53").Select
    Selection.NumberFormat = "#,##0.00"
ActiveCell.Offset(54, -17).Range("A1:G1").Select
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
ActiveCell.Offset(0, 8).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Débit"
ActiveCell.Offset(0, 2).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Crédit"
ActiveCell.Offset(0, 2).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Solde"
ActiveCell.Offset(1, -7).Range("A1").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
ActiveCell.Offset(0, 4).Range("A1").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
ActiveCell.Offset(0, 4).Range("A1").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
ActiveCell.Offset(1, -8).Range("A1").Select
    Selection.NumberFormat = "#,##0.00"
    ActiveCell.FormulaR1C1 = "=R[-55]C"
ActiveCell.Offset(0, 4).Range("A1").Select
    Selection.NumberFormat = "#,##0.00"
    ActiveCell.FormulaR1C1 = "=R[-54]C"
ActiveCell.Offset(0, 4).Range("A1").Select
    Selection.NumberFormat = "#,##0.00"
   ActiveCell.FormulaR1C1 = "=R[-56]C+R[-55]C[-8]-R[-54]C[-4]"
   
    
    


End Sub


Sub Mise_en_page_Comptabilité_L()


ActiveCell.Columns("A:A").EntireColumn.ColumnWidth = 0.5
ActiveCell.Offset(0, 1).Columns("A:A").EntireColumn.ColumnWidth = 22.14
ActiveCell.Offset(0, 2).Columns("A:A").EntireColumn.ColumnWidth = 0.5
ActiveCell.Offset(0, 3).Columns("A:A").EntireColumn.ColumnWidth = 1
ActiveCell.Offset(0, 4).Columns("A:A").EntireColumn.ColumnWidth = 0.5
ActiveCell.Offset(0, 5).Columns("A:A").EntireColumn.ColumnWidth = 10.67
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
ActiveCell.Offset(0, 17).Columns("A:A").EntireColumn.ColumnWidth = 10.71
ActiveCell.Offset(0, 18).Columns("A:A").EntireColumn.ColumnWidth = 0.5


End Sub


