Attribute VB_Name = "Mensuel_I_zoom"
Sub Final_Mensuel_I()


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
        .Order = xlOverThenDown
        .Zoom = 95
    End With
ActiveWindow.View = xlPageLayoutView
    Call Complément_Mensuel_I
Range("I7") = "Janvier"
Columns("A:Q").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Mensuel_I
Range("I7") = "Février"
Columns("A:Q").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Mensuel_I
Range("I7") = "Mars"
Columns("A:Q").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Mensuel_I
Range("I7") = "Avril"
Columns("A:Q").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Mensuel_I
Range("I7") = "Mai"
Columns("A:Q").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Mensuel_I
Range("I7") = "Juin"
Columns("A:Q").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Mensuel_I
Range("I7") = "Juillet"
Columns("A:Q").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Mensuel_I
Range("I7") = "Août"
Columns("A:Q").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Mensuel_I
Range("I7") = "Septembre"
Columns("A:Q").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Mensuel_I
Range("I7") = "Octobre"
Columns("A:Q").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Mensuel_I
Range("I7") = "Novembre"
Columns("A:Q").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Mensuel_I
Range("I7") = "Décembre"
Rows("69").Select
    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
    
Application.StatusBar = False


End Sub


Sub Complément_Mensuel_I()
    
    
    Call Mise_en_page_Comptabilité_I
Range("A1").Select
    Call Fiche_Mensuel_I
Columns("R").Select
    ActiveWindow.SelectedSheets.VPageBreaks.Add Before:=ActiveCell


End Sub
    
    
Sub Fiche_Mensuel_I()

    
    Call Fiche_Comptabilité_I
ActiveCell.Offset(-67, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Mensuel"


End Sub




