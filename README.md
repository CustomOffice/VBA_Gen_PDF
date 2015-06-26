# VBA_Gen_PDF
Macro VBA pour générer un pdf à partir d'un onglet

## Infos
TITRE : Génère un pdf à partir d'un onglet                                  		    
DATE : 24.04.2015                              											

DESCRIPTION :Création d'un PDF à partir d'un onglet excel										

REGLES :																						
- utilise le nom de l'onglet et génère un pdf de cet onglet en appelé le nom_pdf   				
- le chemin par défaut pour l'enregistrement du pdf est l'emplacement du fichier excel			
- si un chemin est spécifié, par défaut il est en absolu, c'est à dire, le chemin complet, si   
vous voulez utiliser le chemin en relatif, il faut forcé l'argument chemin_realtif à true      

##Code
```bash
Sub gen_pdf(nom_feuille As String, nom_pdf As String, Optional chemin As String = "", Optional chemin_relatif As Boolean = False)
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'!!!TITRE :         Génère un pdf à partir d'un onglet                                            	!!!
'!!!DATE :         24.04.2015                              											!!!
'!!!                                                                          						!!!
'!!!DESCRIPTION :Création d'un PDF à partir d'un onglet excel										!!!
'!!!                                                                               					!!!
'!!!REGLES :																						!!!
'!!!- utilise le nom de l'onglet et génère un pdf de cet onglet en appelé le nom_pdf   				!!!
'!!!- le chemin par défaut pour l'enregistrement du pdf est l'emplacement du fichier excel			!!!
'!!!- si un chemin est spécifié, par défaut il est en absolu, c'est à dire, le chemin complet, si   !!!
'!!!vous voulez utiliser le chemin en relatif, il faut forcé l'argument chemin_realtif à true       !!!
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
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
