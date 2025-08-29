Attribute VB_Name = "Comptabilité_F_liste"


Sub Final_Comptabilité_F()


Application.ScreenUpdating = False
Application.StatusBar = "Feuil F"

Sheets.Add.Name = "F"
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
    Call Complément_Comptabilité_F
Range("J7,J75,J143,J211,J279,J347,J415,J483,J551,J619,J687,J755,J823,J891,J959,J1027,J1095,J1163,J1231,J1299,J1367,J1435,J1503,J1571,J1639") = "Janvier"
Range("J1707,J1775,J1843,J1911,J1979,J2047,J2115,J2183,J2251,J2319,J2387,J2455,J2523,J2591,J2659,J2727,J2795,J2863,J2931,J2999,J3067") = "Janvier"
Columns("A:S").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Comptabilité_F
Range("J7,J75,J143,J211,J279,J347,J415,J483,J551,J619,J687,J755,J823,J891,J959,J1027,J1095,J1163,J1231,J1299,J1367,J1435,J1503,J1571,J1639") = "Février"
Range("J1707,J1775,J1843,J1911,J1979,J2047,J2115,J2183,J2251,J2319,J2387,J2455,J2523,J2591,J2659,J2727,J2795,J2863,J2931,J2999,J3067") = "Février"
Columns("A:S").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Comptabilité_F
Range("J7,J75,J143,J211,J279,J347,J415,J483,J551,J619,J687,J755,J823,J891,J959,J1027,J1095,J1163,J1231,J1299,J1367,J1435,J1503,J1571,J1639") = "Mars"
Range("J1707,J1775,J1843,J1911,J1979,J2047,J2115,J2183,J2251,J2319,J2387,J2455,J2523,J2591,J2659,J2727,J2795,J2863,J2931,J2999,J3067") = "Mars"
Columns("A:S").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Comptabilité_F
Range("J7,J75,J143,J211,J279,J347,J415,J483,J551,J619,J687,J755,J823,J891,J959,J1027,J1095,J1163,J1231,J1299,J1367,J1435,J1503,J1571,J1639") = "Avril"
Range("J1707,J1775,J1843,J1911,J1979,J2047,J2115,J2183,J2251,J2319,J2387,J2455,J2523,J2591,J2659,J2727,J2795,J2863,J2931,J2999,J3067") = "Avril"
Columns("A:S").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Comptabilité_F
Range("J7,J75,J143,J211,J279,J347,J415,J483,J551,J619,J687,J755,J823,J891,J959,J1027,J1095,J1163,J1231,J1299,J1367,J1435,J1503,J1571,J1639") = "Mai"
Range("J1707,J1775,J1843,J1911,J1979,J2047,J2115,J2183,J2251,J2319,J2387,J2455,J2523,J2591,J2659,J2727,J2795,J2863,J2931,J2999,J3067") = "Mai"
Columns("A:S").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Comptabilité_F
Range("J7,J75,J143,J211,J279,J347,J415,J483,J551,J619,J687,J755,J823,J891,J959,J1027,J1095,J1163,J1231,J1299,J1367,J1435,J1503,J1571,J1639") = "Juin"
Range("J1707,J1775,J1843,J1911,J1979,J2047,J2115,J2183,J2251,J2319,J2387,J2455,J2523,J2591,J2659,J2727,J2795,J2863,J2931,J2999,J3067") = "Juin"
Columns("A:S").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Comptabilité_F
Range("J7,J75,J143,J211,J279,J347,J415,J483,J551,J619,J687,J755,J823,J891,J959,J1027,J1095,J1163,J1231,J1299,J1367,J1435,J1503,J1571,J1639") = "Juillet"
Range("J1707,J1775,J1843,J1911,J1979,J2047,J2115,J2183,J2251,J2319,J2387,J2455,J2523,J2591,J2659,J2727,J2795,J2863,J2931,J2999,J3067") = "Juillet"
Columns("A:S").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Comptabilité_F
Range("J7,J75,J143,J211,J279,J347,J415,J483,J551,J619,J687,J755,J823,J891,J959,J1027,J1095,J1163,J1231,J1299,J1367,J1435,J1503,J1571,J1639") = "Août"
Range("J1707,J1775,J1843,J1911,J1979,J2047,J2115,J2183,J2251,J2319,J2387,J2455,J2523,J2591,J2659,J2727,J2795,J2863,J2931,J2999,J3067") = "Août"
Columns("A:S").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Comptabilité_F
Range("J7,J75,J143,J211,J279,J347,J415,J483,J551,J619,J687,J755,J823,J891,J959,J1027,J1095,J1163,J1231,J1299,J1367,J1435,J1503,J1571,J1639") = "Septembre"
Range("J1707,J1775,J1843,J1911,J1979,J2047,J2115,J2183,J2251,J2319,J2387,J2455,J2523,J2591,J2659,J2727,J2795,J2863,J2931,J2999,J3067") = "Septembre"
Columns("A:S").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Comptabilité_F
Range("J7,J75,J143,J211,J279,J347,J415,J483,J551,J619,J687,J755,J823,J891,J959,J1027,J1095,J1163,J1231,J1299,J1367,J1435,J1503,J1571,J1639") = "Octobre"
Range("J1707,J1775,J1843,J1911,J1979,J2047,J2115,J2183,J2251,J2319,J2387,J2455,J2523,J2591,J2659,J2727,J2795,J2863,J2931,J2999,J3067") = "Octobre"
Columns("A:S").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Comptabilité_F
Range("J7,J75,J143,J211,J279,J347,J415,J483,J551,J619,J687,J755,J823,J891,J959,J1027,J1095,J1163,J1231,J1299,J1367,J1435,J1503,J1571,J1639") = "Novembre"
Range("J1707,J1775,J1843,J1911,J1979,J2047,J2115,J2183,J2251,J2319,J2387,J2455,J2523,J2591,J2659,J2727,J2795,J2863,J2931,J2999,J3067") = "Novembre"
Columns("A:S").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Comptabilité_F
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
 
    
Sub Complément_Comptabilité_F()
    
    
    Call Mise_en_page_Comptabilité_F
Range("A1").Select
    Call Fiche_Comptabilité_F
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("AK12").Copy Workbooks("Comptabilité.xlsx").Worksheets("F").Range("B7")
Range("B7").Select
    Selection.HorizontalAlignment = xlCenter
Range("A69").Select
    Call Fiche_Comptabilité_F
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("AK13").Copy Workbooks("Comptabilité.xlsx").Worksheets("F").Range("B75")
Range("B75").Select
    Selection.HorizontalAlignment = xlCenter
Range("A137").Select
    Call Fiche_Comptabilité_F
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("AK14").Copy Workbooks("Comptabilité.xlsx").Worksheets("F").Range("B143")
Range("B143").Select
    Selection.HorizontalAlignment = xlCenter
Range("A205").Select
    Call Fiche_Comptabilité_F
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("AK15").Copy Workbooks("Comptabilité.xlsx").Worksheets("F").Range("B211")
Range("B211").Select
    Selection.HorizontalAlignment = xlCenter
Range("A273").Select
    Call Fiche_Comptabilité_F
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("AK16").Copy Workbooks("Comptabilité.xlsx").Worksheets("F").Range("B279")
Range("B279").Select
    Selection.HorizontalAlignment = xlCenter
Range("A341").Select
    Call Fiche_Comptabilité_F
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("AK17").Copy Workbooks("Comptabilité.xlsx").Worksheets("F").Range("B347")
Range("B347").Select
    Selection.HorizontalAlignment = xlCenter
Range("A409").Select
    Call Fiche_Comptabilité_F
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("AK18").Copy Workbooks("Comptabilité.xlsx").Worksheets("F").Range("B415")
Range("B415").Select
    Selection.HorizontalAlignment = xlCenter
Range("A477").Select
    Call Fiche_Comptabilité_F
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("AK19").Copy Workbooks("Comptabilité.xlsx").Worksheets("F").Range("B483")
Range("B483").Select
    Selection.HorizontalAlignment = xlCenter
Range("A545").Select
    Call Fiche_Comptabilité_F
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("AK20").Copy Workbooks("Comptabilité.xlsx").Worksheets("F").Range("B551")
Range("B551").Select
    Selection.HorizontalAlignment = xlCenter
Range("A613").Select
    Call Fiche_Comptabilité_F
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("AK21").Copy Workbooks("Comptabilité.xlsx").Worksheets("F").Range("B619")
Range("B619").Select
    Selection.HorizontalAlignment = xlCenter
Range("A681").Select
    Call Fiche_Comptabilité_F
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("AK22").Copy Workbooks("Comptabilité.xlsx").Worksheets("F").Range("B687")
Range("B687").Select
    Selection.HorizontalAlignment = xlCenter
Range("A749").Select
    Call Fiche_Comptabilité_F
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("AK23").Copy Workbooks("Comptabilité.xlsx").Worksheets("F").Range("B755")
Range("B755").Select
    Selection.HorizontalAlignment = xlCenter
Range("A817").Select
    Call Fiche_Comptabilité_F
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("AK24").Copy Workbooks("Comptabilité.xlsx").Worksheets("F").Range("B823")
Range("B823").Select
    Selection.HorizontalAlignment = xlCenter
Range("A885").Select
    Call Fiche_Comptabilité_F
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("AK25").Copy Workbooks("Comptabilité.xlsx").Worksheets("F").Range("B891")
Range("B891").Select
    Selection.HorizontalAlignment = xlCenter
Range("A953").Select
    Call Fiche_Comptabilité_F
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("AK26").Copy Workbooks("Comptabilité.xlsx").Worksheets("F").Range("B959")
Range("B959").Select
    Selection.HorizontalAlignment = xlCenter
Range("A1021").Select
    Call Fiche_Comptabilité_F
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("AK27").Copy Workbooks("Comptabilité.xlsx").Worksheets("F").Range("B1027")
Range("B1027").Select
    Selection.HorizontalAlignment = xlCenter
Range("A1089").Select
    Call Fiche_Comptabilité_F
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("AK28").Copy Workbooks("Comptabilité.xlsx").Worksheets("F").Range("B1095")
Range("B1095").Select
    Selection.HorizontalAlignment = xlCenter
Range("A1157").Select
    Call Fiche_Comptabilité_F
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("AK29").Copy Workbooks("Comptabilité.xlsx").Worksheets("F").Range("B1163")
Range("B1163").Select
    Selection.HorizontalAlignment = xlCenter
Range("A1225").Select
    Call Fiche_Comptabilité_F
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("AK30").Copy Workbooks("Comptabilité.xlsx").Worksheets("F").Range("B1231")
Range("B1231").Select
    Selection.HorizontalAlignment = xlCenter
Range("A1293").Select
    Call Fiche_Comptabilité_F
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("AK31").Copy Workbooks("Comptabilité.xlsx").Worksheets("F").Range("B1299")
Range("B1299").Select
    Selection.HorizontalAlignment = xlCenter
Range("A1361").Select
    Call Fiche_Comptabilité_F
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("AK32").Copy Workbooks("Comptabilité.xlsx").Worksheets("F").Range("B1367")
Range("B1367").Select
    Selection.HorizontalAlignment = xlCenter
Range("A1429").Select
    Call Fiche_Comptabilité_F
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("AK33").Copy Workbooks("Comptabilité.xlsx").Worksheets("F").Range("B1435")
Range("B1435").Select
    Selection.HorizontalAlignment = xlCenter
Range("A1497").Select
    Call Fiche_Comptabilité_F
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("AK34").Copy Workbooks("Comptabilité.xlsx").Worksheets("F").Range("B1503")
Range("B1503").Select
    Selection.HorizontalAlignment = xlCenter
Range("A1565").Select
    Call Fiche_Comptabilité_F
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("AK35").Copy Workbooks("Comptabilité.xlsx").Worksheets("F").Range("B1571")
Range("B1571").Select
    Selection.HorizontalAlignment = xlCenter
Range("A1633").Select
    Call Fiche_Comptabilité_F
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("AK36").Copy Workbooks("Comptabilité.xlsx").Worksheets("F").Range("B1639")
Range("B1639").Select
    Selection.HorizontalAlignment = xlCenter
Range("A1701").Select
    Call Fiche_Comptabilité_F
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("AK37").Copy Workbooks("Comptabilité.xlsx").Worksheets("F").Range("B1707")
Range("B1707").Select
    Selection.HorizontalAlignment = xlCenter
Range("A1769").Select
    Call Fiche_Comptabilité_F
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("AK38").Copy Workbooks("Comptabilité.xlsx").Worksheets("F").Range("B1775")
Range("B1775").Select
    Selection.HorizontalAlignment = xlCenter
Range("A1837").Select
    Call Fiche_Comptabilité_F
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("AK39").Copy Workbooks("Comptabilité.xlsx").Worksheets("F").Range("B1843")
Range("B1843").Select
    Selection.HorizontalAlignment = xlCenter
Range("A1905").Select
    Call Fiche_Comptabilité_F
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("AK40").Copy Workbooks("Comptabilité.xlsx").Worksheets("F").Range("B1911")
Range("B1911").Select
    Selection.HorizontalAlignment = xlCenter
Range("A1973").Select
    Call Fiche_Comptabilité_F
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("AK41").Copy Workbooks("Comptabilité.xlsx").Worksheets("F").Range("B1979")
Range("B1979").Select
    Selection.HorizontalAlignment = xlCenter
Range("A2041").Select
    Call Fiche_Comptabilité_F
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("AK42").Copy Workbooks("Comptabilité.xlsx").Worksheets("F").Range("B2047")
Range("B2047").Select
    Selection.HorizontalAlignment = xlCenter
Range("A2109").Select
    Call Fiche_Comptabilité_F
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("AK43").Copy Workbooks("Comptabilité.xlsx").Worksheets("F").Range("B2115")
Range("B2115").Select
    Selection.HorizontalAlignment = xlCenter
Range("A2177").Select
    Call Fiche_Comptabilité_F
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("AK44").Copy Workbooks("Comptabilité.xlsx").Worksheets("F").Range("B2183")
Range("B2183").Select
    Selection.HorizontalAlignment = xlCenter
Range("A2245").Select
    Call Fiche_Comptabilité_F
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("AK45").Copy Workbooks("Comptabilité.xlsx").Worksheets("F").Range("B2251")
Range("B2251").Select
    Selection.HorizontalAlignment = xlCenter
Range("A2313").Select
    Call Fiche_Comptabilité_F
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("AK46").Copy Workbooks("Comptabilité.xlsx").Worksheets("F").Range("B2319")
Range("B2319").Select
    Selection.HorizontalAlignment = xlCenter
Range("A2381").Select
    Call Fiche_Comptabilité_F
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("AK47").Copy Workbooks("Comptabilité.xlsx").Worksheets("F").Range("B2387")
Range("B2387").Select
    Selection.HorizontalAlignment = xlCenter
Range("A2449").Select
    Call Fiche_Comptabilité_F
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("AK48").Copy Workbooks("Comptabilité.xlsx").Worksheets("F").Range("B2455")
Range("B2455").Select
    Selection.HorizontalAlignment = xlCenter
Range("A2517").Select
    Call Fiche_Comptabilité_F
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("AK49").Copy Workbooks("Comptabilité.xlsx").Worksheets("F").Range("B2523")
Range("B2523").Select
    Selection.HorizontalAlignment = xlCenter
Range("A2585").Select
    Call Fiche_Comptabilité_F
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("AK50").Copy Workbooks("Comptabilité.xlsx").Worksheets("F").Range("B2591")
Range("B2591").Select
    Selection.HorizontalAlignment = xlCenter
Range("A2653").Select
    Call Fiche_Comptabilité_F
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("AK51").Copy Workbooks("Comptabilité.xlsx").Worksheets("F").Range("B2659")
Range("B2659").Select
    Selection.HorizontalAlignment = xlCenter
Range("A2721").Select
    Call Fiche_Comptabilité_F
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("AK52").Copy Workbooks("Comptabilité.xlsx").Worksheets("F").Range("B2727")
Range("B2727").Select
    Selection.HorizontalAlignment = xlCenter
Range("A2789").Select
    Call Fiche_Comptabilité_F
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("AK53").Copy Workbooks("Comptabilité.xlsx").Worksheets("F").Range("B2795")
Range("B2795").Select
    Selection.HorizontalAlignment = xlCenter
Range("A2857").Select
    Call Fiche_Comptabilité_F
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("AK54").Copy Workbooks("Comptabilité.xlsx").Worksheets("F").Range("B2863")
Range("B2863").Select
    Selection.HorizontalAlignment = xlCenter
Range("A2925").Select
    Call Fiche_Comptabilité_F
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("AK55").Copy Workbooks("Comptabilité.xlsx").Worksheets("F").Range("B2931")
Range("B2931").Select
    Selection.HorizontalAlignment = xlCenter
Range("A2993").Select
    Call Fiche_Comptabilité_F
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("AK56").Copy Workbooks("Comptabilité.xlsx").Worksheets("F").Range("B2999")
Range("B2999").Select
    Selection.HorizontalAlignment = xlCenter
Range("A3061").Select
    Call Fiche_Comptabilité_F
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("AK57").Copy Workbooks("Comptabilité.xlsx").Worksheets("F").Range("B3067")
Range("B3067").Select
    Selection.HorizontalAlignment = xlCenter
Columns("T").Select
    ActiveWindow.SelectedSheets.VPageBreaks.Add Before:=ActiveCell


End Sub
    
    
Sub Fiche_Comptabilité_F()
    
    
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
    ActiveCell.FormulaR1C1 = "F"
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
ActiveCell.Offset(1, -8).Range("A1").Select
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
    ActiveCell.FormulaR1C1 = "Achat march."
ActiveCell.Offset(0, 2).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Frais d'achat"
ActiveCell.Offset(0, 2).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "TVA"
ActiveCell.Offset(0, -16).Range("A1:S1").Select
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
ActiveCell.Offset(2, 9).Range("A1:A53").Select
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
    ActiveCell.FormulaR1C1 = "Achat march."
ActiveCell.Offset(0, 2).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Frais d'achat"
ActiveCell.Offset(0, 2).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "TVA"
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
    ActiveCell.FormulaR1C1 = "=SUM(R[-56]C:R[-4]C)"
ActiveCell.Offset(0, 4).Range("A1").Select
    Selection.NumberFormat = "#,##0.00"
    ActiveCell.FormulaR1C1 = "=SUM(R[-56]C:R[-4]C)"
ActiveCell.Offset(0, 4).Range("A1").Select
    Selection.NumberFormat = "#,##0.00"
    ActiveCell.FormulaR1C1 = "=SUM(R[-56]C:R[-4]C)"


End Sub


Sub Mise_en_page_Comptabilité_F()


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

