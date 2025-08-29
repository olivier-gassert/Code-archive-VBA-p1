Attribute VB_Name = "Fiduciaire_"
Sub Bouton_Nouveau_Fichier_Comptabilité()


Application.ScreenUpdating = False
Application.DisplayAlerts = False

Workbooks.Add
ActiveWorkbook.SaveAs Filename:="Macintosh HD:Users:bureaucentral:Documents:Elisa Gassert:Programmes:Fiduciaire:Comptabilité.xlsx", FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Call Final_Comptabilité_L
    Call Final_Comptabilité_I
    Call Final_Comptabilité_F
    Call Final_Comptabilité_D
    Call Final_Comptabilité_C
Sheets(Array("Feuil1")).Select
    ActiveWindow.SelectedSheets.Delete
ActiveWorkbook.Save
    
End Sub


Sub Bouton_Transfert_Comptabilité_à_Séparations()


Application.ScreenUpdating = False
Application.DisplayAlerts = False

'Sheets("Séparations des comptes").Select
    'Call Transfert_Liste_a_Separat

    
End Sub


Sub Bouton_Nouveau_Fichier_Mensuel()


Application.ScreenUpdating = False
Application.DisplayAlerts = False

Workbooks.Add
    Call Final_Mensuel_L
    Call Final_Mensuel_I
    Call Final_Mensuel_F
    Call Final_Mensuel_D
    Call Final_Mensuel_C
Sheets(Array("Feuil1")).Select
    ActiveWindow.SelectedSheets.Delete
    
    
End Sub

Sub Bouton_Nouveau_Fichier_Liste_et_Séparations()


Application.ScreenUpdating = False
Application.DisplayAlerts = False

Workbooks.Add
    Call Final_Séparations_Des_Comptes
    Call Final_Liste_Des_Comptes
Sheets(Array("Feuil1")).Select
    ActiveWindow.SelectedSheets.Delete
    
    
End Sub

Sub Bouton_Transfert_Liste_à_Séparations()


Application.ScreenUpdating = False
Application.DisplayAlerts = False

Sheets("Séparations").Select
    Call Transfert_Liste_a_Separation_L
    Call Transfert_Liste_a_Separation_I
    Call Transfert_Liste_a_Separation_F
    Call Transfert_Liste_a_Separation_D
    Call Transfert_Liste_a_Separation_C

    
End Sub



