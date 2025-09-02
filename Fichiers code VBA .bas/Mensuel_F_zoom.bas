Attribute VB_Name = "Mensuel_F_zoom"


Sub Final_Mensuel_F()


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
        .Order = xlOverThenDown
        .Zoom = 95
    End With
ActiveWindow.View = xlPageLayoutView
    Call Complément_Mensuel_F
Range("J7") = "Janvier"
Columns("A:S").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Mensuel_F
Range("J7") = "Février"
Columns("A:S").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Mensuel_F
Range("J7") = "Mars"
Columns("A:S").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Mensuel_F
Range("J7") = "Avril"
Columns("A:S").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Mensuel_F
Range("J7") = "Mai"
Columns("A:S").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Mensuel_F
Range("J7") = "Juin"
Columns("A:S").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Mensuel_F
Range("J7") = "Juillet"
Columns("A:S").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Mensuel_F
Range("J7") = "Août"
Columns("A:S").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Mensuel_F
Range("J7") = "Septembre"
Columns("A:S").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Mensuel_F
Range("J7") = "Octobre"
Columns("A:S").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Mensuel_F
Range("J7") = "Novembre"
Columns("A:S").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Mensuel_F
Range("J7") = "Décembre"
Rows("69").Select
    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
    
Application.StatusBar = False


End Sub


Sub Complément_Mensuel_F()
    
    
    Call Mise_en_page_Comptabilité_F
Range("A1").Select
    Call Fiche_Mensuel_F
Columns("T").Select
    ActiveWindow.SelectedSheets.VPageBreaks.Add Before:=ActiveCell


End Sub
    
    
Sub Fiche_Mensuel_F()

    
    Call Fiche_Comptabilité_F
ActiveCell.Offset(-67, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Mensuel"
    

End Sub







