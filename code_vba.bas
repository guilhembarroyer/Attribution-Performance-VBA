Attribute VB_Name = "recuperation_donnees"
Option Explicit
Option Base 0
Sub procRecupData()

'procédure permettant de choisir une base de donnée et y récupérant les informations sur les dates, les champs disponibles

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'% Déclaration des variables %
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

Dim wbD As Workbook 'fichier des données
Dim wsD As Worksheet 'feuille utilisée dans le fichier des données
Dim adresse As String 'fullname du fichier des données
Dim nFields As Integer 'nbre de champs de la base utilisée
Dim nRecords As Long 'nbre d'enregistrements de la base utilisée
Dim colDates As Integer 'colonne des dates
Dim d As Long 'date recherchée (en entier)

Dim ws As Worksheet 'feuille de calcul (quelconque)
Dim message As String 'message
Dim i As Long, n As Long 'compteurs de ligne
Dim j As Integer 'compteurs de colonnes

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'% I. Ouverture du fichier des données %
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


'sélection du fichier (par l'utilisateur)
adresse = Application.GetOpenFilename

'affectation de wb au fichier choisi par l'utilisateur
Set wbD = Workbooks.Open(adresse)


'si plusieurs feuilles, choix de la feuille à utiliser par l'utilisateur via une inputbox après la récupération des noms des feui
message = wbD.Worksheets(1).Name
If wbD.Worksheets.Count > 1 Then

    'récupération des noms des feuilles (tables)
    message = "Le fichier sélectionné contient les feuilles suivantes :" & vbCrLf
    For Each ws In wbD.Worksheets
        message = ws.Name & vbCrLf
    Next ws
    message = message & vbCrLf & vbCrLf & "Quelle feuille doit être utulisée?"

    message = InputBox(message, "Choix de la feuille.", wbD.Worksheets(1).Name)

End If
Set wsD = wbD.Worksheets(message)

'dénombrement des champs et des enregistrements
nFields = wsD.Cells(1, wsD.Columns.Count).End(xlToLeft).Column
nRecords = wsD.Cells(wsD.Rows.Count, 1).End(xlUp).Row - 1

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'% II. Récupération des champs %
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

'affectation de ws à l'unique feuille de ce fichier
Set ws = ThisWorkbook.Worksheets("intro")

'report sur l'unique feuille de ce fichier du nom du fichier de données et de la feuille
ws.Range("fichier").Value = wbD.Name
ws.Range("feuille").Value = message

'effacement des champs sur la feuille du fichier et report des valeurs du fichier de données
With ws.Range("champs")
    .Resize(ws.Rows.Count - .Row, 1).ClearContents
    .Offset(1, 0).Resize(nFields, 1).Value = WorksheetFunction.Transpose(wsD.Cells(1, 1).Resize(1, nFields))
End With

'effacement des dates contenues par ce fichier
With ws.Range("dates")
    .Resize(ws.Rows.Count - .Row, 1).ClearContents
End With

'recherche d'un champ date
For j = 1 To nFields
    If LCase(Left(wsD.Cells(1, j).Value, 4)) = "date" Then colDates = j: Exit For
Next j

'cas où un champ de date a été trouvé
If colDates < nFields + 1 Then
    
    'mise des dates en format numérique
    wsD.Columns(colDates).Cells.NumberFormat = "General"
    'tri en fonction des dates
    wsD.Cells(2, 1).Resize(nRecords, nFields).Sort Key1:=wsD.Cells(1, colDates), Order1:=xlAscending  'xlDescending
    'boucle sur les lignes et report des dates différentes
    j = 1
    i = 2
    Do While i <= nRecords + 1
        'récupération de la date
        d = wsD.Cells(i, colDates).Value
        'calcul de dates identiques à celle de la ligne i
        n = WorksheetFunction.CountIf(wsD.Cells(2, colDates).Resize(nRecords, 1), "=" & d)
        'report de la date
        j = j + 1
        ws.Range("dates").Cells(j, 1).Value = wsD.Cells(i, colDates).Value
        'saut à la ligne de la date suivante
        i = i + n
    Loop
    wsD.Columns(colDates).NumberFormat = "dd/mm/yyyy"
End If
ws.Range("dates").Resize(j, 1).NumberFormat = "dd/mm/yyyy"

End Sub
Sub imp_data_attribution()


'procédure récupérant les données pour l'attribution de performance à la Brinson en s'appuyant sur les données de la feuille "intro" de _
ce ficher sur les données à utiliser (dans les cellules "fichier" et "feuille") pour la variable définie (dans la cellule "catégorie") _
et pour la date souhaité (celules "date").


'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'% Déclaration des variables %
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

Dim wbD As Workbook 'le fichier de données
Dim wsD As Worksheet 'la feuille des données
Dim nFields As Integer 'nbre de champs de la base utilisée
Dim nRecords As Long 'nbre d'enregistrements de la base utilisée

Dim colDates As Integer 'colonne des dates
Dim colSect As Integer 'colonne des secteurs
Dim colRend As Integer 'colonne des rendements
Dim colPort As Integer 'colonne du portefeuille
Dim colBench As Integer 'colonne du benchmark
Dim colId As Integer 'colonne des identifiants

Dim d As Long 'date recherchée (en entier)
Dim ligne_debut As Long 'première ligne contenant la date recherchée
Dim ligne_fin As Long 'dernière ligne contenant la date recherchée



Dim ws As Worksheet 'feuille de calcul (quelconque)
Dim message As String 'message
Dim i As Long, n As Long 'compteurs de ligne
Dim j As Integer 'compteurs de colonnes


'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'% I. Prise en main des données %
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


'affectatationde ws à l'unique feuille de ce fichier
Set ws = ThisWorkbook.Worksheets("intro")

'test sur l'ouverture u fichier de données
On Error Resume Next
Workbooks(ws.Range("fichier").Value).Activate
If Err.Number > 0 Then MsgBox "Le fichier de données doit être déjà ouvert." & vbCrLf & "La procédure va s'arrêter." & vbCrLf & _
"Ouver le fichier et relancer après la procédure.": Exit Sub
On Error GoTo 0

'affectation de wsD à la feuille de données définies sur l'unique feuille de ce fichier (cellules "fichier" et feuille")
Set wsD = Workbooks(ws.Range("fichier").Value).Worksheets(ws.Range("feuille").Value)

'dénombrement des champs et des enregistrements

nFields = wsD.Cells(1, wsD.Columns.Count).End(xlToLeft).Column
nRecords = wsD.Cells(Rows.Count, 1).End(xlUp).Row - 1

'recherche des positions des champs (variables colRend, colSect, colPort, colBench, colDates, colId à définir avec méthode find appliquée à la ligne 1)
'<ligne 1 de la feuille>.Find(what:=<chaine de caractère>, LookAt:=xlwhole).column

colRend = Rows(1).Find(what:="return", LookAt:=xlWhole).Column
colSect = Rows(1).Find(what:="sector", LookAt:=xlWhole).Column
colPort = Rows(1).Find(what:="portfolio", LookAt:=xlWhole).Column
colBench = Rows(1).Find(what:="benchmark", LookAt:=xlWhole).Column
colDates = Rows(1).Find(what:="date", LookAt:=xlWhole).Column
colId = Rows(1).Find(what:="id", LookAt:=xlWhole).Column



'récupération de la date (en entier)
d = ws.Range("date").Value

'tri des données par rapport aux dates (après les avoir mis au format "General")
'<plage>.Sort Key1:=<cellule ligne 1 contenant l'intitulé>, Order1:=xlAscending
wsD.Cells(2, 1).Resize(nRecords, nFields).Sort Key1:=wsD.Cells(1, colDates), Order1:=xlAscending


'recherche de la première et de la dernière ligne des enregistements dont la date est d

j = 0

For i = 1 To nRecords
    If wsD.Cells(i, colDates).Value = d And j = 0 Then
        ligne_debut = i
        j = 1
    ElseIf wsD.Cells(i, colDates) = d And j = 1 Then
        ligne_fin = i
    End If
Next i





'calcul du nombre d'enregistrement (n) dont la date est d
n = ligne_fin - ligne_debut + 1

'affectation de ws à la feuille "calcul"
Set ws = ThisWorkbook.Worksheets("calcul")

'effacement des données
i = ws.Cells(Rows.Count, 1).End(xlUp).Row
ws.Cells(1, 1).Resize(i, 5).ClearContents

'report des intitulés des variables souhaitées pour le tableau à construire sur la feuille "calcul"
ws.Cells(1, 1).Resize(1, 5).Value = Array("id", "secteur", "rend", "port", "bench")

'récupération des données (et report à partir de la ligne 2) du tableau de la feuille "calcul" (et dont la date est d)
ws.Cells(2, 1).Resize(n, 1).Value = wsD.Range(wsD.Cells(ligne_debut, colId), wsD.Cells(ligne_fin, colId)).Value
ws.Cells(2, 2).Resize(n, 1).Value = wsD.Range(wsD.Cells(ligne_debut, colSect), wsD.Cells(ligne_fin, colSect)).Value
ws.Cells(2, 3).Resize(n, 1).Value = wsD.Range(wsD.Cells(ligne_debut, colRend), wsD.Cells(ligne_fin, colRend)).Value
ws.Cells(2, 4).Resize(n, 1).Value = wsD.Range(wsD.Cells(ligne_debut, colPort), wsD.Cells(ligne_fin, colPort)).Value
ws.Cells(2, 5).Resize(n, 1).Value = wsD.Range(wsD.Cells(ligne_debut, colBench), wsD.Cells(ligne_fin, colBench)).Value

End Sub

