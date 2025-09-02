Attribute VB_Name = "Comptabilité_I_liste"


Sub Final_Comptabilité_I()


Application.ScreenUpdating = False
Application.StatusBar = "Feuil I"

Sheets.Add.Name = "I"
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
    Call Complément_Comptabilité_I
Range("I7,I75,I143,I211,I279,I347,I415,I483,I551,I619,I687,I755,I823,I891,I959,I1027,I1095,I1163,I1231,I1299,I1367,I1435,I1503,I1571,I1639") = "Janvier"
Range("I1707,I1775,I1843,I1911,I1979,I2047,I2115,I2183,I2251,I2319,I2387,I2455,I2523,I2591,I2659,I2727,I2795,I2863,I2931,I2999,I3067") = "Janvier"
Columns("A:Q").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Comptabilité_I
Range("I7,I75,I143,I211,I279,I347,I415,I483,I551,I619,I687,I755,I823,I891,I959,I1027,I1095,I1163,I1231,I1299,I1367,I1435,I1503,I1571,I1639") = "Février"
Range("I1707,I1775,I1843,I1911,I1979,I2047,I2115,I2183,I2251,I2319,I2387,I2455,I2523,I2591,I2659,I2727,I2795,I2863,I2931,I2999,I3067") = "Février"
Columns("A:Q").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Comptabilité_I
Range("I7,I75,I143,I211,I279,I347,I415,I483,I551,I619,I687,I755,I823,I891,I959,I1027,I1095,I1163,I1231,I1299,I1367,I1435,I1503,I1571,I1639") = "Mars"
Range("I1707,I1775,I1843,I1911,I1979,I2047,I2115,I2183,I2251,I2319,I2387,I2455,I2523,I2591,I2659,I2727,I2795,I2863,I2931,I2999,I3067") = "Mars"
Columns("A:Q").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Comptabilité_I
Range("I7,I75,I143,I211,I279,I347,I415,I483,I551,I619,I687,I755,I823,I891,I959,I1027,I1095,I1163,I1231,I1299,I1367,I1435,I1503,I1571,I1639") = "Avril"
Range("I1707,I1775,I1843,I1911,I1979,I2047,I2115,I2183,I2251,I2319,I2387,I2455,I2523,I2591,I2659,I2727,I2795,I2863,I2931,I2999,I3067") = "Avril"
Columns("A:Q").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Comptabilité_I
Range("I7,I75,I143,I211,I279,I347,I415,I483,I551,I619,I687,I755,I823,I891,I959,I1027,I1095,I1163,I1231,I1299,I1367,I1435,I1503,I1571,I1639") = "Mai"
Range("I1707,I1775,I1843,I1911,I1979,I2047,I2115,I2183,I2251,I2319,I2387,I2455,I2523,I2591,I2659,I2727,I2795,I2863,I2931,I2999,I3067") = "Mai"
Columns("A:Q").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Comptabilité_I
Range("I7,I75,I143,I211,I279,I347,I415,I483,I551,I619,I687,I755,I823,I891,I959,I1027,I1095,I1163,I1231,I1299,I1367,I1435,I1503,I1571,I1639") = "Juin"
Range("I1707,I1775,I1843,I1911,I1979,I2047,I2115,I2183,I2251,I2319,I2387,I2455,I2523,I2591,I2659,I2727,I2795,I2863,I2931,I2999,I3067") = "juin"
Columns("A:Q").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Comptabilité_I
Range("I7,I75,I143,I211,I279,I347,I415,I483,I551,I619,I687,I755,I823,I891,I959,I1027,I1095,I1163,I1231,I1299,I1367,I1435,I1503,I1571,I1639") = "Juillet"
Range("I1707,I1775,I1843,I1911,I1979,I2047,I2115,I2183,I2251,I2319,I2387,I2455,I2523,I2591,I2659,I2727,I2795,I2863,I2931,I2999,I3067") = "Juillet"
Columns("A:Q").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Comptabilité_I
Range("I7,I75,I143,I211,I279,I347,I415,I483,I551,I619,I687,I755,I823,I891,I959,I1027,I1095,I1163,I1231,I1299,I1367,I1435,I1503,I1571,I1639") = "Août"
Range("I1707,I1775,I1843,I1911,I1979,I2047,I2115,I2183,I2251,I2319,I2387,I2455,I2523,I2591,I2659,I2727,I2795,I2863,I2931,I2999,I3067") = "Août"
Columns("A:Q").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Comptabilité_I
Range("I7,I75,I143,I211,I279,I347,I415,I483,I551,I619,I687,I755,I823,I891,I959,I1027,I1095,I1163,I1231,I1299,I1367,I1435,I1503,I1571,I1639") = "Septembre"
Range("I1707,I1775,I1843,I1911,I1979,I2047,I2115,I2183,I2251,I2319,I2387,I2455,I2523,I2591,I2659,I2727,I2795,I2863,I2931,I2999,I3067") = "Septembre"
Columns("A:Q").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Comptabilité_I
Range("I7,I75,I143,I211,I279,I347,I415,I483,I551,I619,I687,I755,I823,I891,I959,I1027,I1095,I1163,I1231,I1299,I1367,I1435,I1503,I1571,I1639") = "Octobre"
Range("I1707,I1775,I1843,I1911,I1979,I2047,I2115,I2183,I2251,I2319,I2387,I2455,I2523,I2591,I2659,I2727,I2795,I2863,I2931,I2999,I3067") = "Octobre"
Columns("A:Q").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Comptabilité_I
Range("I7,I75,I143,I211,I279,I347,I415,I483,I551,I619,I687,I755,I823,I891,I959,I1027,I1095,I1163,I1231,I1299,I1367,I1435,I1503,I1571,I1639") = "Novembre"
Range("I1707,I1775,I1843,I1911,I1979,I2047,I2115,I2183,I2251,I2319,I2387,I2455,I2523,I2591,I2659,I2727,I2795,I2863,I2931,I2999,I3067") = "Novembre"
Columns("A:Q").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Comptabilité_I
Range("I7,I75,I143,I211,I279,I347,I415,I483,I551,I619,I687,I755,I823,I891,I959,I1027,I1095,I1163,I1231,I1299,I1367,I1435,I1503,I1571,I1639") = "Décembre"
Range("I1707,I1775,I1843,I1911,I1979,I2047,I2115,I2183,I2251,I2319,I2387,I2455,I2523,I2591,I2659,I2727,I2795,I2863,I2931,I2999,I3067") = "Décembre"
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
 
    
Sub Complément_Comptabilité_I()
    
    
    Call Mise_en_page_Comptabilité_I
Range("A1").Select
    Call Fiche_Comptabilité_I
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BB12").Copy Workbooks("Comptabilité.xlsx").Worksheets("I").Range("C7")
Range("C7").Select
    Selection.HorizontalAlignment = xlCenter
Range("A69").Select
    Call Fiche_Comptabilité_I
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BB13").Copy Workbooks("Comptabilité.xlsx").Worksheets("I").Range("C75")
Range("C75").Select
Selection.HorizontalAlignment = xlCenter
Range("A137").Select
    Call Fiche_Comptabilité_I
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BB14").Copy Workbooks("Comptabilité.xlsx").Worksheets("I").Range("C143")
Range("C143").Select
    Selection.HorizontalAlignment = xlCenter
Range("A205").Select
    Call Fiche_Comptabilité_I
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BB15").Copy Workbooks("Comptabilité.xlsx").Worksheets("I").Range("C211")
Range("C211").Select
    Selection.HorizontalAlignment = xlCenter
Range("A273").Select
    Call Fiche_Comptabilité_I
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BB16").Copy Workbooks("Comptabilité.xlsx").Worksheets("I").Range("C279")
Range("C279").Select
    Selection.HorizontalAlignment = xlCenter
Range("A341").Select
    Call Fiche_Comptabilité_I
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BB17").Copy Workbooks("Comptabilité.xlsx").Worksheets("I").Range("C347")
Range("C347").Select
    Selection.HorizontalAlignment = xlCenter
Range("A409").Select
    Call Fiche_Comptabilité_I
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BB18").Copy Workbooks("Comptabilité.xlsx").Worksheets("I").Range("C415")
Range("C415").Select
    Selection.HorizontalAlignment = xlCenter
Range("A477").Select
    Call Fiche_Comptabilité_I
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BB19").Copy Workbooks("Comptabilité.xlsx").Worksheets("I").Range("C483")
Range("C483").Select
    Selection.HorizontalAlignment = xlCenter
Range("A545").Select
    Call Fiche_Comptabilité_I
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BB20").Copy Workbooks("Comptabilité.xlsx").Worksheets("I").Range("C551")
Range("C551").Select
    Selection.HorizontalAlignment = xlCenter
Range("A613").Select
    Call Fiche_Comptabilité_I
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BB21").Copy Workbooks("Comptabilité.xlsx").Worksheets("I").Range("C619")
Range("C619").Select
    Selection.HorizontalAlignment = xlCenter
Range("A681").Select
    Call Fiche_Comptabilité_I
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BB22").Copy Workbooks("Comptabilité.xlsx").Worksheets("I").Range("C687")
Range("C687").Select
    Selection.HorizontalAlignment = xlCenter
Range("A749").Select
    Call Fiche_Comptabilité_I
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BB23").Copy Workbooks("Comptabilité.xlsx").Worksheets("I").Range("C755")
Range("C755").Select
    Selection.HorizontalAlignment = xlCenter
Range("A817").Select
    Call Fiche_Comptabilité_I
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BB24").Copy Workbooks("Comptabilité.xlsx").Worksheets("I").Range("C823")
Range("C823").Select
    Selection.HorizontalAlignment = xlCenter
Range("A885").Select
    Call Fiche_Comptabilité_I
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BB25").Copy Workbooks("Comptabilité.xlsx").Worksheets("I").Range("C891")
Range("C891").Select
    Selection.HorizontalAlignment = xlCenter
Range("A953").Select
    Call Fiche_Comptabilité_I
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BB26").Copy Workbooks("Comptabilité.xlsx").Worksheets("I").Range("C959")
Range("C959").Select
    Selection.HorizontalAlignment = xlCenter
Range("A1021").Select
    Call Fiche_Comptabilité_I
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BB27").Copy Workbooks("Comptabilité.xlsx").Worksheets("I").Range("C1027")
Range("C1027").Select
    Selection.HorizontalAlignment = xlCenter
Range("A1089").Select
    Call Fiche_Comptabilité_I
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BB28").Copy Workbooks("Comptabilité.xlsx").Worksheets("I").Range("C1095")
Range("C1095").Select
    Selection.HorizontalAlignment = xlCenter
Range("A1157").Select
    Call Fiche_Comptabilité_I
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BB29").Copy Workbooks("Comptabilité.xlsx").Worksheets("I").Range("C1163")
Range("C1163").Select
    Selection.HorizontalAlignment = xlCenter
Range("A1225").Select
    Call Fiche_Comptabilité_I
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BB30").Copy Workbooks("Comptabilité.xlsx").Worksheets("I").Range("C1231")
Range("C1231").Select
    Selection.HorizontalAlignment = xlCenter
Range("A1293").Select
    Call Fiche_Comptabilité_I
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BB31").Copy Workbooks("Comptabilité.xlsx").Worksheets("I").Range("C1299")
Range("C1299").Select
    Selection.HorizontalAlignment = xlCenter
Range("A1361").Select
    Call Fiche_Comptabilité_I
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BB32").Copy Workbooks("Comptabilité.xlsx").Worksheets("I").Range("C1367")
Range("C1367").Select
    Selection.HorizontalAlignment = xlCenter
Range("A1429").Select
    Call Fiche_Comptabilité_I
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BB33").Copy Workbooks("Comptabilité.xlsx").Worksheets("I").Range("C1435")
Range("C1435").Select
    Selection.HorizontalAlignment = xlCenter
Range("A1497").Select
    Call Fiche_Comptabilité_I
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BB34").Copy Workbooks("Comptabilité.xlsx").Worksheets("I").Range("C1503")
Range("C1503").Select
    Selection.HorizontalAlignment = xlCenter
Range("A1565").Select
    Call Fiche_Comptabilité_I
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BB35").Copy Workbooks("Comptabilité.xlsx").Worksheets("I").Range("C1571")
Range("C1571").Select
    Selection.HorizontalAlignment = xlCenter
Range("A1633").Select
    Call Fiche_Comptabilité_I
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BB36").Copy Workbooks("Comptabilité.xlsx").Worksheets("I").Range("C1639")
Range("C1639").Select
    Selection.HorizontalAlignment = xlCenter
Range("A1701").Select
    Call Fiche_Comptabilité_I
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BB37").Copy Workbooks("Comptabilité.xlsx").Worksheets("I").Range("C1707")
Range("C1707").Select
    Selection.HorizontalAlignment = xlCenter
Range("A1769").Select
    Call Fiche_Comptabilité_I
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BB38").Copy Workbooks("Comptabilité.xlsx").Worksheets("I").Range("C1775")
Range("C1775").Select
    Selection.HorizontalAlignment = xlCenter
Range("A1837").Select
    Call Fiche_Comptabilité_I
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BB39").Copy Workbooks("Comptabilité.xlsx").Worksheets("I").Range("C1843")
Range("C1843").Select
    Selection.HorizontalAlignment = xlCenter
Range("A1905").Select
    Call Fiche_Comptabilité_I
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BB40").Copy Workbooks("Comptabilité.xlsx").Worksheets("I").Range("C1911")
Range("C1911").Select
    Selection.HorizontalAlignment = xlCenter
Range("A1973").Select
    Call Fiche_Comptabilité_I
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BB41").Copy Workbooks("Comptabilité.xlsx").Worksheets("I").Range("C1979")
Range("C1979").Select
    Selection.HorizontalAlignment = xlCenter
Range("A2041").Select
    Call Fiche_Comptabilité_I
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BB42").Copy Workbooks("Comptabilité.xlsx").Worksheets("I").Range("C2047")
Range("C2047").Select
    Selection.HorizontalAlignment = xlCenter
Range("A2109").Select
    Call Fiche_Comptabilité_I
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BB43").Copy Workbooks("Comptabilité.xlsx").Worksheets("I").Range("C2115")
Range("C2115").Select
    Selection.HorizontalAlignment = xlCenter
Range("A2177").Select
    Call Fiche_Comptabilité_I
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BB44").Copy Workbooks("Comptabilité.xlsx").Worksheets("I").Range("C2183")
Range("C2183").Select
    Selection.HorizontalAlignment = xlCenter
Range("A2245").Select
    Call Fiche_Comptabilité_I
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BB45").Copy Workbooks("Comptabilité.xlsx").Worksheets("I").Range("C2251")
Range("C2251").Select
    Selection.HorizontalAlignment = xlCenter
Range("A2313").Select
    Call Fiche_Comptabilité_I
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BB46").Copy Workbooks("Comptabilité.xlsx").Worksheets("I").Range("C2319")
Range("C2319").Select
    Selection.HorizontalAlignment = xlCenter
Range("A2381").Select
    Call Fiche_Comptabilité_I
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BB47").Copy Workbooks("Comptabilité.xlsx").Worksheets("I").Range("C2387")
Range("C2387").Select
    Selection.HorizontalAlignment = xlCenter
Range("A2449").Select
    Call Fiche_Comptabilité_I
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BB48").Copy Workbooks("Comptabilité.xlsx").Worksheets("I").Range("C2455")
Range("C2455").Select
    Selection.HorizontalAlignment = xlCenter
Range("A2517").Select
    Call Fiche_Comptabilité_I
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BB49").Copy Workbooks("Comptabilité.xlsx").Worksheets("I").Range("C2523")
Range("C2523").Select
    Selection.HorizontalAlignment = xlCenter
Range("A2585").Select
    Call Fiche_Comptabilité_I
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BB50").Copy Workbooks("Comptabilité.xlsx").Worksheets("I").Range("C2591")
Range("C2591").Select
    Selection.HorizontalAlignment = xlCenter
Range("A2653").Select
    Call Fiche_Comptabilité_I
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BB51").Copy Workbooks("Comptabilité.xlsx").Worksheets("I").Range("C2659")
Range("C2659").Select
    Selection.HorizontalAlignment = xlCenter
Range("A2721").Select
    Call Fiche_Comptabilité_I
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BB52").Copy Workbooks("Comptabilité.xlsx").Worksheets("I").Range("C2727")
Range("C2727").Select
    Selection.HorizontalAlignment = xlCenter
Range("A2789").Select
    Call Fiche_Comptabilité_I
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BB53").Copy Workbooks("Comptabilité.xlsx").Worksheets("I").Range("C2795")
Range("C2795").Select
    Selection.HorizontalAlignment = xlCenter
Range("A2857").Select
    Call Fiche_Comptabilité_I
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BB54").Copy Workbooks("Comptabilité.xlsx").Worksheets("I").Range("C2863")
Range("C2863").Select
    Selection.HorizontalAlignment = xlCenter
Range("A2925").Select
    Call Fiche_Comptabilité_I
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BB55").Copy Workbooks("Comptabilité.xlsx").Worksheets("I").Range("C2931")
Range("C2931").Select
    Selection.HorizontalAlignment = xlCenter
Range("A2993").Select
    Call Fiche_Comptabilité_I
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BB56").Copy Workbooks("Comptabilité.xlsx").Worksheets("I").Range("C2999")
Range("C2999").Select
    Selection.HorizontalAlignment = xlCenter
Range("A3061").Select
    Call Fiche_Comptabilité_I
Workbooks("Comptes.xlsx").Worksheets("Liste").Range("BB57").Copy Workbooks("Comptabilité.xlsx").Worksheets("I").Range("C3067")
Range("C3067").Select
    Selection.HorizontalAlignment = xlCenter
Columns("R").Select
    ActiveWindow.SelectedSheets.VPageBreaks.Add Before:=ActiveCell


End Sub
    
    
Sub Fiche_Comptabilité_I()
    
    
ActiveCell.Offset(0, 1).Range("A1:O1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Comptabilité"
ActiveCell.Offset(1, 0).Range("A1").Select
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
ActiveCell.Offset(0, 13).Range("A1").Select
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
ActiveCell.Offset(1, -12).Range("A1:M1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    ActiveCell.FormulaR1C1 = "I"
ActiveCell.Offset(2, -1).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Compte"
ActiveCell.Offset(0, 2).Range("A1:G1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Mois"
ActiveCell.Offset(0, 2).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Année"
ActiveCell.Offset(1, -11).Range("A1").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
ActiveCell.Offset(0, 3).Range("A1").Select
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
ActiveCell.Offset(0, 5).Range("A1").Select
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
ActiveCell.Offset(0, 3).Range("A1").Select
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
ActiveCell.Offset(0, 1).Range("A1").Select
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
ActiveCell.Offset(1, -6).Range("A1").Select
    Selection.HorizontalAlignment = xlCenter
ActiveCell.Offset(0, 6).Range("A1").Select
    Selection.HorizontalAlignment = xlCenter
    Selection.NumberFormat = "yyyy"
    ActiveCell.FormulaR1C1 = Date
ActiveCell.Offset(3, -9).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Date"
ActiveCell.Offset(0, 2).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Payement"
ActiveCell.Offset(0, 2).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "TVA"
ActiveCell.Offset(0, -12).Range("A1:O1").Select
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
ActiveCell.Offset(2, 9).Range("A1:A53").Select
    Selection.NumberFormat = "#,##0.00"
ActiveCell.Offset(0, 4).Range("A1:A53").Select
    Selection.NumberFormat = "#,##0.00"
ActiveCell.Offset(54, -13).Range("A1:G1").Select
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
ActiveCell.Offset(0, 8).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Payement"
ActiveCell.Offset(0, 2).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "TVA"
ActiveCell.Offset(1, -3).Range("A1").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
ActiveCell.Offset(0, 4).Range("A1").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
ActiveCell.Offset(1, -4).Range("A1").Select
    Selection.NumberFormat = "#,##0.00"
    ActiveCell.FormulaR1C1 = "=SUM(R[-56]C:R[-4]C)"
ActiveCell.Offset(0, 4).Range("A1").Select
    Selection.NumberFormat = "#,##0.00"
    ActiveCell.FormulaR1C1 = "=SUM(R[-56]C:R[-4]C)"
'Columns("R").Select
    'ActiveWindow.SelectedSheets.VPageBreaks.Add Before:=ActiveCell
'Rows("64").Select
    'ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell


End Sub


Sub Mise_en_page_Comptabilité_I()


ActiveCell.Columns("A:A").EntireColumn.ColumnWidth = 10
ActiveCell.Offset(0, 1).Columns("A:A").EntireColumn.ColumnWidth = 0.5
ActiveCell.Offset(0, 2).Columns("A:A").EntireColumn.ColumnWidth = 22.17
ActiveCell.Offset(0, 3).Columns("A:A").EntireColumn.ColumnWidth = 0.5
ActiveCell.Offset(0, 4).Columns("A:A").EntireColumn.ColumnWidth = 1
ActiveCell.Offset(0, 5).Columns("A:A").EntireColumn.ColumnWidth = 0.5
ActiveCell.Offset(0, 6).Columns("A:A").EntireColumn.ColumnWidth = 10.67
ActiveCell.Offset(0, 7).Columns("A:A").EntireColumn.ColumnWidth = 0.5
ActiveCell.Offset(0, 8).Columns("A:A").EntireColumn.ColumnWidth = 1
ActiveCell.Offset(0, 9).Columns("A:A").EntireColumn.ColumnWidth = 0.5
ActiveCell.Offset(0, 10).Columns("A:A").EntireColumn.ColumnWidth = 10.67
ActiveCell.Offset(0, 11).Columns("A:A").EntireColumn.ColumnWidth = 0.5
ActiveCell.Offset(0, 12).Columns("A:A").EntireColumn.ColumnWidth = 1
ActiveCell.Offset(0, 13).Columns("A:A").EntireColumn.ColumnWidth = 0.5
ActiveCell.Offset(0, 14).Columns("A:A").EntireColumn.ColumnWidth = 10.67
ActiveCell.Offset(0, 15).Columns("A:A").EntireColumn.ColumnWidth = 0.5
ActiveCell.Offset(0, 16).Columns("A:A").EntireColumn.ColumnWidth = 10


End Sub

