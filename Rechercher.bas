Attribute VB_Name = "modRecherche"
'
' DDT PDF Excel - bouton « Rechercher »
'
' Lit le numéro (Recherche!B2) et le dossier des PDF (Recherche!B3, facultatif),
' lance main.py avec le Python du .venv, puis recopie les feuilles « Résultats »
' et « Non analysés » du fichier produit dans ce classeur.
'
' Le classeur doit être enregistré à la racine du projet, à côté de main.py et de .venv.
' Le module s'appelle modRecherche (et non Rechercher) : un module et une macro
' du même nom empêchent Excel de retrouver la macro associée au bouton.
'
Option Explicit

Private Const FEUILLE_RECHERCHE As String = "Recherche"
Private Const FEUILLE_RESULTATS As String = "Résultats"
Private Const FEUILLE_NON_ANALYSES As String = "Non analysés"
Private Const TITRE As String = "Rechercher"


Public Sub Rechercher()
    Dim racine As String, numero As String, dossierPdf As String
    Dim python As String, script As String, journal As String, fichierTemp As String
    Dim moteur As String, programme As String
    Dim commande As String, codeSortie As Long, nbTrouves As Long
    Dim temp As Workbook, message As String, valeurB2 As Variant

    On Error GoTo Erreur

    racine = DossierClasseur()
    With ThisWorkbook.Worksheets(FEUILLE_RECHERCHE)
        valeurB2 = .Range("B2").Value
        numero = Trim$(CStr(valeurB2))
        dossierPdf = Trim$(CStr(.Range("B3").Value))
    End With

    If numero = "" Then
        MsgBox "Saisis le numéro à rechercher dans la cellule B2 de la feuille « Recherche ».", _
               vbExclamation, TITRE
        Exit Sub
    End If
    ' Saisi sans le format Texte, 00481 devient le nombre 481 : les zéros en tête sont perdus.
    ' On teste le type de la valeur, pas seulement le format : passer B2 en Texte après la
    ' saisie ne rend pas les zéros, la cellule contient toujours le nombre 481.
    If VarType(valeurB2) <> vbString And IsNumeric(valeurB2) Then
        MsgBox "Mets la cellule B2 au format Texte puis ressaisis le numéro.", vbExclamation, TITRE
        Exit Sub
    End If
    ' " couperait la ligne de commande, % y serait interprété par cmd
    If InStr(numero, """") > 0 Or InStr(numero, "%") > 0 Then
        MsgBox "Le numéro ne doit contenir ni guillemet ("") ni signe %.", vbExclamation, TITRE
        Exit Sub
    End If

    If dossierPdf = "" Then dossierPdf = racine & "\PDF_Colis"
    ' Un \ juste avant le guillemet fermant (C:\Dossier\") serait lu par Python comme \"
    If Right$(dossierPdf, 1) = "\" Then dossierPdf = dossierPdf & "."

    moteur = racine & "\moteur\Recherche_PDF.exe"
    If Len(Dir$(moteur)) > 0 Then
        ' Version livrée : moteur empaqueté, sans Python (UTF-8 activé dans l'exe)
        programme = Guillemets(moteur)
    Else
        ' Mode développement : main.py avec le Python du .venv
        python = racine & "\.venv\Scripts\python.exe"
        script = racine & "\main.py"
        If Len(Dir$(python)) = 0 Then
            MsgBox "Moteur introuvable :" & vbLf & moteur & vbLf & vbLf & _
                   "Le classeur doit rester dans le dossier Recherche_PDF, à côté du dossier moteur.", _
                   vbCritical, TITRE
            Exit Sub
        End If
        If Len(Dir$(script)) = 0 Then
            MsgBox "main.py introuvable :" & vbLf & script & vbLf & vbLf & _
                   "Le classeur doit être à côté de main.py.", vbCritical, TITRE
            Exit Sub
        End If
        ' -X utf8 : journal.txt écrit en UTF-8 (relu tel quel par LireJournal)
        programme = Guillemets(python) & " -X utf8 " & Guillemets(script)
    End If

    journal = racine & "\journal.txt"
    ' Dans %TEMP% et non à côté du classeur : OneDrive synchroniserait (et verrouillerait) le fichier
    fichierTemp = Environ$("TEMP") & "\ddt_resultat_" & Format$(Now, "yyyymmdd_hhnnss") & ".xlsx"

    ' -- : le numéro n'est jamais pris pour une option, même s'il commence par "-"
    commande = programme & _
               " --dossier " & Guillemets(dossierPdf) & _
               " --sortie " & Guillemets(fichierTemp) & _
               " -- " & Guillemets(numero) & _
               " > " & Guillemets(journal) & " 2>&1"
    ' cmd /c "..." : la paire de guillemets extérieure protège ceux de la commande
    commande = Guillemets(Environ$("ComSpec")) & " /c """ & commande & """"

    Application.Cursor = xlWait
    Application.StatusBar = "Recherche en cours…"
    DoEvents

    ' 0 = fenêtre cachée, True = attendre la fin ; renvoie le code de sortie de main.py
    codeSortie = CreateObject("WScript.Shell").Run(commande, 0, True)

    If codeSortie <> 0 Then
        RetablirExcel
        MsgBox LireJournal(journal), vbCritical, TITRE & " - erreur (code " & codeSortie & ")"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Set temp = Workbooks.Open(Filename:=fichierTemp, ReadOnly:=True, UpdateLinks:=0, AddToMru:=False)
    CopierFeuille temp.Worksheets(FEUILLE_RESULTATS), FeuilleCible(FEUILLE_RESULTATS)
    CopierFeuille temp.Worksheets(FEUILLE_NON_ANALYSES), FeuilleCible(FEUILLE_NON_ANALYSES)
    temp.Close SaveChanges:=False
    Set temp = Nothing
    Kill fichierTemp

    nbTrouves = NombreTrouves(ThisWorkbook.Worksheets(FEUILLE_RESULTATS))
    ThisWorkbook.Worksheets(FEUILLE_RESULTATS).Activate
    RetablirExcel

    If nbTrouves = 0 Then
        MsgBox "Aucun PDF ne contient ce numéro.", vbInformation, TITRE
    Else
        MsgBox nbTrouves & " PDF trouvé(s)", vbInformation, TITRE
    End If
    Exit Sub

Erreur:
    message = Err.Description
    On Error Resume Next
    If Not temp Is Nothing Then temp.Close SaveChanges:=False
    If fichierTemp <> "" Then
        If Len(Dir$(fichierTemp)) > 0 Then Kill fichierTemp
    End If
    RetablirExcel
    On Error GoTo 0
    MsgBox "Erreur : " & message, vbCritical, TITRE
End Sub


Private Sub RetablirExcel()
    Application.Cursor = xlDefault
    Application.StatusBar = False
    Application.ScreenUpdating = True
End Sub


Private Function Guillemets(ByVal texte As String) As String
    Guillemets = """" & texte & """"
End Function


' Dossier local du classeur. Sur OneDrive, ThisWorkbook.Path renvoie une adresse web
' (https://d.docs.live.net/...) qu'il faut convertir en chemin C:\Users\...\OneDrive\...
Private Function DossierClasseur() As String
    Dim chemin As String, position As Long, i As Long, racineOneDrive As String

    chemin = ThisWorkbook.Path
    If chemin = "" Then Err.Raise vbObjectError + 1, , "Enregistre d'abord le classeur à côté de main.py."
    If LCase$(Left$(chemin, 4)) <> "http" Then
        DossierClasseur = chemin
        Exit Function
    End If

    If InStr(1, chemin, ".sharepoint.com", vbTextCompare) > 0 Then
        ' OneDrive professionnel : https://xxx-my.sharepoint.com/personal/compte/Documents/...
        racineOneDrive = Environ$("OneDriveCommercial")
        position = InStr(1, chemin, "/Documents", vbTextCompare) + Len("/Documents")
    Else
        ' OneDrive personnel : https://d.docs.live.net/<identifiant>/... (4e "/" = début du chemin)
        racineOneDrive = Environ$("OneDriveConsumer")
        If racineOneDrive = "" Then racineOneDrive = Environ$("OneDrive")
        For i = 1 To 4
            position = InStr(position + 1, chemin, "/")
        Next i
    End If

    chemin = racineOneDrive & Replace(Replace(Mid$(chemin, position), "/", "\"), "%20", " ")
    If racineOneDrive = "" Or Len(Dir$(chemin, vbDirectory)) = 0 Then
        Err.Raise vbObjectError + 2, , "Dossier local du classeur introuvable (OneDrive) : " & _
                  ThisWorkbook.Path
    End If
    DossierClasseur = chemin
End Function


Private Function FeuilleCible(ByVal nom As String) As Worksheet
    On Error Resume Next
    Set FeuilleCible = ThisWorkbook.Worksheets(nom)
    On Error GoTo 0
    If FeuilleCible Is Nothing Then
        Set FeuilleCible = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        FeuilleCible.Name = nom
    End If
End Function


' Vide la feuille cible puis y recopie la feuille source : valeurs, formats,
' liens hypertextes (Range.Copy les conserve) et largeurs de colonnes.
Private Sub CopierFeuille(ByVal source As Worksheet, ByVal cible As Worksheet)
    Dim plage As Range, colonne As Long

    cible.Hyperlinks.Delete
    cible.Cells.Clear

    Set plage = source.UsedRange
    plage.Copy Destination:=cible.Range(plage.Address)
    For colonne = 1 To plage.Column + plage.Columns.Count - 1
        cible.Columns(colonne).ColumnWidth = source.Columns(colonne).ColumnWidth
    Next colonne
End Sub


' Lignes de résultat = lignes à partir de la 4e dont la colonne « Page(s) » est remplie
' (la ligne « Aucun PDF ne contient ce numéro. » n'a rien en colonne B).
Private Function NombreTrouves(ByVal feuille As Worksheet) As Long
    Dim derniere As Long
    derniere = feuille.Cells(feuille.Rows.Count, 1).End(xlUp).Row
    If derniere < 4 Then Exit Function
    NombreTrouves = Application.WorksheetFunction.CountA(feuille.Range("B4:B" & derniere))
End Function


' Contenu de journal.txt (UTF-8), limité à la fin pour tenir dans un MsgBox.
Private Function LireJournal(ByVal chemin As String) As String
    Dim flux As Object, texte As String

    On Error GoTo Illisible
    Set flux = CreateObject("ADODB.Stream")
    flux.Type = 2          ' texte
    flux.Charset = "utf-8"
    flux.Open
    flux.LoadFromFile chemin
    texte = Trim$(flux.ReadText)
    flux.Close

    If texte = "" Then texte = "(journal.txt est vide)"
    If Len(texte) > 900 Then texte = "…" & Right$(texte, 900)
    LireJournal = texte
    Exit Function

Illisible:
    LireJournal = "La recherche a échoué et journal.txt est illisible :" & vbLf & chemin
End Function
