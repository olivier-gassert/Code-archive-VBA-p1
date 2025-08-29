Attribute VB_Name = "Mensuel_D_zoom"
Sub Final_Mensuel_D()


Application.ScreenUpdating = False
Application.StatusBar = "Feuil D"

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
        .Order = xlOverThenDown
        .Zoom = 95
    End With
ActiveWindow.View = xlPageLayoutView
    Call Complément_Mensuel_D
Range("N7") = "Janvier"
Columns("A:AA").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Mensuel_D
Range("N7") = "Février"
Columns("A:AA").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Mensuel_D
Range("N7") = "Mars"
Columns("A:AA").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Mensuel_D
Range("N7") = "Avril"
Columns("A:AA").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Mensuel_D
Range("N7") = "Mai"
Columns("A:AA").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Mensuel_D
Range("N7") = "Juin"
Columns("A:AA").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Mensuel_D
Range("N7") = "Juillet"
Columns("A:AA").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Mensuel_D
Range("N7") = "Août"
Columns("A:AA").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Mensuel_D
Range("N7") = "Septembre"
Columns("A:AA").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Mensuel_D
Range("N7") = "Octobre"
Columns("A:AA").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Mensuel_D
Range("N7") = "Novembre"
Columns("A:AA").Select
    Selection.Insert Shift:=xlToRight
    Call Complément_Mensuel_D
Range("N7") = "Décembre"
Rows("69").Select
    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
    
Application.StatusBar = False


End Sub


Sub Complément_Mensuel_D()
    
    
    Call Mise_en_page_Comptabilité_D
Range("A1").Select
    Call Fiche_Mensuel_D
Columns("AB").Select
    ActiveWindow.SelectedSheets.VPageBreaks.Add Before:=ActiveCell


End Sub
    
    
Sub Fiche_Mensuel_D()

    
    Call Fiche_Comptabilité_D
ActiveCell.Offset(-67, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Mensuel"
ActiveCell.Offset(9, 0).Range("A1").Select
    Selection.UnMerge
    Selection.Interior.ColorIndex = False
    Selection.ClearContents
ActiveCell.Offset(0, 4).Range("A1").Select
    Selection.UnMerge
    Selection.Interior.ColorIndex = False
    Selection.ClearContents
ActiveCell.Offset(0, 4).Range("A1").Select
    Selection.UnMerge
    Selection.Interior.ColorIndex = False
    Selection.ClearContents

End Sub



