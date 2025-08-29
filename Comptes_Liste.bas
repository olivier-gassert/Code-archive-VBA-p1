Attribute VB_Name = "Comptes_Liste"


Sub Final_Liste_Des_Comptes()


Application.ScreenUpdating = False
Application.StatusBar = "Liste des comptes"

Sheets.Add.Name = "Liste"
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
Call Complément_Liste_Des_Comptes
Range("C3") = "L"
Columns("A:Q").Select
    Selection.Insert Shift:=xlToRight
Call Complément_Liste_Des_Comptes
Range("C3") = "I"
Columns("A:Q").Select
    Selection.Insert Shift:=xlToRight
Call Complément_Liste_Des_Comptes
Range("C3") = "F"
Columns("A:Q").Select
    Selection.Insert Shift:=xlToRight
Call Complément_Liste_Des_Comptes
Range("C3") = "D"
Columns("A:Q").Select
    Selection.Insert Shift:=xlToRight
 Call Complément_Liste_Des_Comptes
Range("C3") = "C"
Rows("69").Select
    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell


End Sub


Sub Complément_Liste_Des_Comptes()


    Call Mise_en_Page_Liste_Des_Comptes
Range("A1").Select
    Call Fiche_Liste_Des_Comptes
Columns("R").Select
    ActiveWindow.SelectedSheets.VPageBreaks.Add Before:=ActiveCell


End Sub


Sub Fiche_Liste_Des_Comptes()


    Call Fiche_Comptabilité_C
ActiveCell.Offset(0, 0).Range("A1").Select
    Selection.ClearContents
ActiveCell.Offset(0, -4).Range("A1").Select
    Selection.ClearContents
ActiveCell.Offset(-2, 0).Range("A1").Select
    Selection.ClearContents
 ActiveCell.Offset(0, 2).Range("A1").Select
    Selection.ClearContents
ActiveCell.Offset(-56, 0).Range("A1").Select
    Selection.ClearContents
ActiveCell.Offset(0, -2).Range("A1").Select
    Selection.ClearContents
ActiveCell.Offset(0, -2).Range("A1").Select
    Selection.ClearContents
  ActiveCell.Offset(-3, 7).Range("A1").Select
    Selection.ClearContents
 ActiveCell.Offset(-2, 0).Range("A1").Select
    Selection.ClearContents
ActiveCell.Offset(0, -2).Range("A1").Select
    Selection.ClearContents
 ActiveCell.Offset(0, -2).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Comptes"
 

End Sub


Sub Mise_en_Page_Liste_Des_Comptes()


    Call Mise_en_page_Comptabilité_C


End Sub
