# VBA_Gen_PDF
Macro VBA pour générer un pdf à partir d'un onglet

## Infos
	```bash

 
Sub gen_pdf(nom_feuille As String, nom_pdf As String, Optional chemin As String = "", Optional chemin_relatif As Boolean = False)
    'déclaration des variables
    Dim feuille_actuelle As String
   
    'génère le chemin
    If chemin = "" Then
        chemin_pdf = ThisWorkbook.Path & Application.PathSeparator & nom_pdf & ".pdf"
    Else
        If chemin_relatif = False Then
            chemin_pdf = chemin & Application.PathSeparator & nom_pdf & ".pdf"
        Else
            chemin_pdf = ThisWorkbook.Path & Application.PathSeparator & chemin & Application.PathSeparator & nom_pdf & ".pdf"
        End If
    End If
   
    'sélection l'onglet, génère le pdf, et reviens à l'onglet de départ
    Application.ScreenUpdating = False
    feuille_actuelle = ActiveSheet.Name
    Sheets(nom_feuille).Select
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=chemin_pdf _
        , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=True
    Sheets(feuille_actuelle).Select
    Application.ScreenUpdating = True
End Sub

	```
