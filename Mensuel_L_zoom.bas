Attribute VB_Name = "Mensuel_L_zoom"


Sub Final_Mensuel_L()


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
        .Order = xlOverThenDown
        .Zoom = 95
    End With
ActiveWindow.View = xlPageLayoutView
    Call Complément_Mensuel_L
Range("J7") = "Janvier"
Columns("A:S").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Mensuel_L
Range("J7") = "Février"
Columns("A:S").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Mensuel_L
Range("J7") = "Mars"
Columns("A:S").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Mensuel_L
Range("J7") = "Avril"
Columns("A:S").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Mensuel_L
Range("J7") = "Mai"
Columns("A:S").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Mensuel_L
Range("J7") = "Juin"
Columns("A:S").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Mensuel_L
Range("J7") = "Juillet"
Columns("A:S").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Mensuel_L
Range("J7") = "Août"
Columns("A:S").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Mensuel_L
Range("J7") = "Septembre"
Columns("A:S").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Mensuel_L
Range("J7") = "Octobre"
Columns("A:S").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Mensuel_L
Range("J7") = "Novembre"
Columns("A:S").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Mensuel_L
Range("J7") = "Décembre"
Rows("69").Select
    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
    
Application.StatusBar = False


End Sub


Sub Complément_Mensuel_L()
    
    
    Call Mise_en_page_Comptabilité_L
Range("A1").Select
    Call Fiche_Mensuel_L
Columns("T").Select
    ActiveWindow.SelectedSheets.VPageBreaks.Add Before:=ActiveCell


End Sub
    
    
Sub Fiche_Mensuel_L()

    
    Call Fiche_Comptabilité_L
ActiveCell.Offset(-67, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Mensuel"
ActiveCell.Offset(11, 0).Range("A1").Select
ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = ""
ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = ""
ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = ""
ActiveCell.Offset(54, 8).Range("A1").Select
   ActiveCell.FormulaR1C1 = "=SUM(R[-56]C:R[-4]C)"
ActiveCell.Offset(0, 4).Range("A1").Select
   ActiveCell.FormulaR1C1 = "=SUM(R[-56]C:R[-4]C)"
   ActiveCell.Offset(0, 4).Range("A1").Select
   ActiveCell.FormulaR1C1 = "=SUM(R[-56]C:R[-4]C)"
 
 
End Sub





