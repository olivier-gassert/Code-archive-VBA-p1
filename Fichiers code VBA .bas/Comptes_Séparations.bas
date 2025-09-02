Attribute VB_Name = "Comptes_Séparations"


Sub Final_Séparations_Des_Comptes()


Application.ScreenUpdating = False
Application.StatusBar = "Feuil Séparations des comptes"

Sheets.Add.Name = "Séparations"
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
    Call Complément_Séparations_Des_Comptes
Range("B3,B71,B139,B207,B275,B343,B411,B479,B547,B615,B683,B751,B819,B887,B955,B1023,B1091,B1159,B1227,B1295,B1363,B1431,B1499,B1567,B1635") = "L"
Range("B1703,B1771,B1839,B1907,B1975,B2043,B2111,B2179,B2247,B2315,B2383,B2451,B2519,B2587,B2655,B2723,B2791,B2859,B2927,B2995,B3063") = "L"
Columns("A:Q").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Séparations_Des_Comptes
Range("B3,B71,B139,B207,B275,B343,B411,B479,B547,B615,B683,B751,B819,B887,B955,B1023,B1091,B1159,B1227,B1295,B1363,B1431,B1499,B1567,B1635") = "I"
Range("B1703,B1771,B1839,B1907,B1975,B2043,B2111,B2179,B2247,B2315,B2383,B2451,B2519,B2587,B2655,B2723,B2791,B2859,B2927,B2995,B3063") = "I"
Columns("A:Q").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Séparations_Des_Comptes
Range("B3,B71,B139,B207,B275,B343,B411,B479,B547,B615,B683,B751,B819,B887,B955,B1023,B1091,B1159,B1227,B1295,B1363,B1431,B1499,B1567,B1635") = "F"
Range("B1703,B1771,B1839,B1907,B1975,B2043,B2111,B2179,B2247,B2315,B2383,B2451,B2519,B2587,B2655,B2723,B2791,B2859,B2927,B2995,B3063") = "F"
Columns("A:Q").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Séparations_Des_Comptes
Range("B3,B71,B139,B207,B275,B343,B411,B479,B547,B615,B683,B751,B819,B887,B955,B1023,B1091,B1159,B1227,B1295,B1363,B1431,B1499,B1567,B1635") = "D"
Range("B1703,B1771,B1839,B1907,B1975,B2043,B2111,B2179,B2247,B2315,B2383,B2451,B2519,B2587,B2655,B2723,B2791,B2859,B2927,B2995,B3063") = "D"
Columns("A:Q").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Séparations_Des_Comptes
Range("B3,B71,B139,B207,B275,B343,B411,B479,B547,B615,B683,B751,B819,B887,B955,B1023,B1091,B1159,B1227,B1295,B1363,B1431,B1499,B1567,B1635") = "C"
Range("B1703,B1771,B1839,B1907,B1975,B2043,B2111,B2179,B2247,B2315,B2383,B2451,B2519,B2587,B2655,B2723,B2791,B2859,B2927,B2995,B3063") = "C"

    
End Sub


Sub Complément_Séparations_Des_Comptes()
    
    
    Call Mise_en_page_Séparations_Des_Comptes
Range("A1").Select
    Call Fiche_Séparations_Des_Comptes
Range("A69").Select
    Call Fiche_Séparations_Des_Comptes
Range("A137").Select
    Call Fiche_Séparations_Des_Comptes
Range("A205").Select
    Call Fiche_Séparations_Des_Comptes
Range("A273").Select
    Call Fiche_Séparations_Des_Comptes
Range("A341").Select
    Call Fiche_Séparations_Des_Comptes
Range("A409").Select
    Call Fiche_Séparations_Des_Comptes
Range("A477").Select
    Call Fiche_Séparations_Des_Comptes
Range("A545").Select
    Call Fiche_Séparations_Des_Comptes
Range("A613").Select
    Call Fiche_Séparations_Des_Comptes
Range("A681").Select
    Call Fiche_Séparations_Des_Comptes
Range("A749").Select
    Call Fiche_Séparations_Des_Comptes
Range("A817").Select
    Call Fiche_Séparations_Des_Comptes
Range("A885").Select
    Call Fiche_Séparations_Des_Comptes
Range("A953").Select
    Call Fiche_Séparations_Des_Comptes
Range("A1021").Select
    Call Fiche_Séparations_Des_Comptes
Range("A1089").Select
    Call Fiche_Séparations_Des_Comptes
Range("A1157").Select
    Call Fiche_Séparations_Des_Comptes
Range("A1225").Select
    Call Fiche_Séparations_Des_Comptes
Range("A1293").Select
    Call Fiche_Séparations_Des_Comptes
Range("A1361").Select
    Call Fiche_Séparations_Des_Comptes
Range("A1429").Select
    Call Fiche_Séparations_Des_Comptes
Range("A1497").Select
    Call Fiche_Séparations_Des_Comptes
Range("A1565").Select
    Call Fiche_Séparations_Des_Comptes
Range("A1633").Select
    Call Fiche_Séparations_Des_Comptes
Range("A1701").Select
    Call Fiche_Séparations_Des_Comptes
Range("A1769").Select
    Call Fiche_Séparations_Des_Comptes
Range("A1837").Select
    Call Fiche_Séparations_Des_Comptes
Range("A1905").Select
    Call Fiche_Séparations_Des_Comptes
Range("A1973").Select
    Call Fiche_Séparations_Des_Comptes
Range("A2041").Select
    Call Fiche_Séparations_Des_Comptes
Range("A2109").Select
    Call Fiche_Séparations_Des_Comptes
Range("A2177").Select
    Call Fiche_Séparations_Des_Comptes
Range("A2245").Select
    Call Fiche_Séparations_Des_Comptes
Range("A2313").Select
    Call Fiche_Séparations_Des_Comptes
Range("A2381").Select
    Call Fiche_Séparations_Des_Comptes
Range("A2449").Select
    Call Fiche_Séparations_Des_Comptes
Range("A2517").Select
    Call Fiche_Séparations_Des_Comptes
Range("A2585").Select
    Call Fiche_Séparations_Des_Comptes
Range("A2653").Select
    Call Fiche_Séparations_Des_Comptes
Range("A2721").Select
    Call Fiche_Séparations_Des_Comptes
Range("A2789").Select
    Call Fiche_Séparations_Des_Comptes
Range("A2857").Select
    Call Fiche_Séparations_Des_Comptes
Range("A2925").Select
    Call Fiche_Séparations_Des_Comptes
Range("A2993").Select
    Call Fiche_Séparations_Des_Comptes
Range("A3061").Select
    Call Fiche_Séparations_Des_Comptes
Columns("AB").Select
    ActiveWindow.SelectedSheets.VPageBreaks.Add Before:=ActiveCell


End Sub


Sub Fiche_Séparations_Des_Comptes()


ActiveCell.Offset(0, 0).Range("A1:AA1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Comptabilité"
ActiveCell.Offset(1, 0).Range("A1").Select
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
ActiveCell.Offset(0, 25).Range("A1").Select
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
ActiveCell.Offset(1, -24).Range("A1:Y1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
ActiveCell.Offset(2, -1).Range("A1:AA1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Compte"
ActiveCell.Offset(1, 0).Range("A1").Select
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
ActiveCell.Offset(0, 25).Range("A1").Select
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
ActiveCell.Offset(1, -24).Range("A1:Y2").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.VerticalAlignment = xlCenter
    Selection.Font.Size = 20


End Sub


Sub Mise_en_page_Séparations_Des_Comptes()


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

