Attribute VB_Name = "recuperation_donnees"
Option Explicit
Option Base 0
Sub procRecupData()

'procŽdure permettant de choisir une base de donnŽe et y rŽcupŽrant les informations sur les dates, les champs disponibles

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'% DŽclaration des variables %
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

Dim wbD As Workbook 'fichier des donnŽes
Dim wsD As Worksheet 'feuille utilisŽe dans le fichier des donnŽes
Dim adresse As String 'fullname du fichier des donnŽes
Dim nFields As Integer 'nbre de champs de la base utilisŽe
Dim nRecords As Long 'nbre d'enregistrements de la base utilisŽe
Dim colDates As Integer 'colonne des dates
Dim d As Long 'date recherchŽe (en entier)

Dim ws As Worksheet 'feuille de calcul (quelconque)
Dim message As String 'message
Dim i As Long, n As Long 'compteurs de ligne
Dim j As Integer 'compteurs de colonnes

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'% I. Ouverture du fichier des donnŽes %
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


'sŽlection du fichier (par l'utilisateur)
adresse = Application.GetOpenFilename

'affectation de wb au fichier choisi par l'utilisateur
Set wbD = Workbooks.Open(adresse)


'si plusieurs feuilles, choix de la feuille ˆ utiliser par l'utilisateur via une inputbox aprs la rŽcupŽration des noms des feui
message = wbD.Worksheets(1).Name
If wbD.Worksheets.Count > 1 Then

    'rŽcupŽration des noms des feuilles (tables)
    message = "Le fichier sŽlectionnŽ contient les feuilles suivantes :" & vbCrLf
    For Each ws In wbD.Worksheets
        message = ws.Name & vbCrLf
    Next ws
    message = message & vbCrLf & vbCrLf & "Quelle feuille doit tre utulisŽe?"

    message = InputBox(message, "Choix de la feuille.", wbD.Worksheets(1).Name)

End If
Set wsD = wbD.Worksheets(message)

'dŽnombrement des champs et des enregistrements
nFields = wsD.Cells(1, wsD.Columns.Count).End(xlToLeft).Column
nRecords = wsD.Cells(wsD.Rows.Count, 1).End(xlUp).Row - 1

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'% II. RŽcupŽration des champs %
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

'affectation de ws ˆ l'unique feuille de ce fichier
Set ws = ThisWorkbook.Worksheets("intro")

'report sur l'unique feuille de ce fichier du nom du fichier de donnŽes et de la feuille
ws.Range("fichier").Value = wbD.Name
ws.Range("feuille").Value = message

'effacement des champs sur la feuille du fichier et report des valeurs du fichier de donnŽes
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

'cas o un champ de date a ŽtŽ trouvŽ
If colDates < nFields + 1 Then
    
    'mise des dates en format numŽrique
    wsD.Columns(colDates).Cells.NumberFormat = "General"
    'tri en fonction des dates
    wsD.Cells(2, 1).Resize(nRecords, nFields).Sort Key1:=wsD.Cells(1, colDates), Order1:=xlAscending  'xlDescending
    'boucle sur les lignes et report des dates diffŽrentes
    j = 1
    i = 2
    Do While i <= nRecords + 1
        'rŽcupŽration de la date
        d = wsD.Cells(i, colDates).Value
        'calcul de dates identiques ˆ celle de la ligne i
        n = WorksheetFunction.CountIf(wsD.Cells(2, colDates).Resize(nRecords, 1), "=" & d)
        'report de la date
        j = j + 1
        ws.Range("dates").Cells(j, 1).Value = wsD.Cells(i, colDates).Value
        'saut ˆ la ligne de la date suivante
        i = i + n
    Loop
    wsD.Columns(colDates).NumberFormat = "dd/mm/yyyy"
End If
ws.Range("dates").Resize(j, 1).NumberFormat = "dd/mm/yyyy"

End Sub
Sub imp_data_attribution()


'procŽdure rŽcupŽrant les donnŽes pour l'attribution de performance ˆ la Brinson en s'appuyant sur les donnŽes de la feuille "intro" de _
ce ficher sur les donnŽes ˆ utiliser (dans les cellules "fichier" et "feuille") pour la variable dŽfinie (dans la cellule "catŽgorie") _
et pour la date souhaitŽ (celules "date").


'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'% DŽclaration des variables %
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

Dim wbD As Workbook 'le fichier de donnŽes
Dim wsD As Worksheet 'la feuille des donnŽes
Dim nFields As Integer 'nbre de champs de la base utilisŽe
Dim nRecords As Long 'nbre d'enregistrements de la base utilisŽe

Dim colDates As Integer 'colonne des dates
Dim colSect As Integer 'colonne des secteurs
Dim colRend As Integer 'colonne des rendements
Dim colPort As Integer 'colonne du portefeuille
Dim colBench As Integer 'colonne du benchmark
Dim colId As Integer 'colonne des identifiants

Dim d As Long 'date recherchŽe (en entier)
Dim ligne_debut As Long 'premire ligne contenant la date recherchŽe
Dim ligne_fin As Long 'dernire ligne contenant la date recherchŽe



Dim ws As Worksheet 'feuille de calcul (quelconque)
Dim message As String 'message
Dim i As Long, n As Long 'compteurs de ligne
Dim j As Integer 'compteurs de colonnes


'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'% I. Prise en main des donnŽes %
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


'affectatationde ws ˆ l'unique feuille de ce fichier
Set ws = ThisWorkbook.Worksheets("intro")

'test sur l'ouverture u fichier de donnŽes
On Error Resume Next
Workbooks(ws.Range("fichier").Value).Activate
If Err.Number > 0 Then MsgBox "Le fichier de donnŽes doit tre dŽjˆ ouvert." & vbCrLf & "La procŽdure va s'arrter." & vbCrLf & _
"Ouver le fichier et relancer aprs la procŽdure.": Exit Sub
On Error GoTo 0

'affectation de wsD ˆ la feuille de donnŽes dŽfinies sur l'unique feuille de ce fichier (cellules "fichier" et feuille")
Set wsD = Workbooks(ws.Range("fichier").Value).Worksheets(ws.Range("feuille").Value)

'dŽnombrement des champs et des enregistrements

nFields = wsD.Cells(1, wsD.Columns.Count).End(xlToLeft).Column
nRecords = wsD.Cells(Rows.Count, 1).End(xlUp).Row - 1

'recherche des positions des champs (variables colRend, colSect, colPort, colBench, colDates, colId ˆ dŽfinir avec mŽthode find appliquŽe ˆ la ligne 1)
'<ligne 1 de la feuille>.Find(what:=<chaine de caractre>, LookAt:=xlwhole).column

colRend = Rows(1).Find(what:="return", LookAt:=xlWhole).Column
colSect = Rows(1).Find(what:="sector", LookAt:=xlWhole).Column
colPort = Rows(1).Find(what:="portfolio", LookAt:=xlWhole).Column
colBench = Rows(1).Find(what:="benchmark", LookAt:=xlWhole).Column
colDates = Rows(1).Find(what:="date", LookAt:=xlWhole).Column
colId = Rows(1).Find(what:="id", LookAt:=xlWhole).Column



'rŽcupŽration de la date (en entier)
d = ws.Range("date").Value

'tri des donnŽes par rapport aux dates (aprs les avoir mis au format "General")
'<plage>.Sort Key1:=<cellule ligne 1 contenant l'intitulŽ>, Order1:=xlAscending
wsD.Cells(2, 1).Resize(nRecords, nFields).Sort Key1:=wsD.Cells(1, colDates), Order1:=xlAscending


'recherche de la premire et de la dernire ligne des enregistements dont la date est d

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

'affectation de ws ˆ la feuille "calcul"
Set ws = ThisWorkbook.Worksheets("calcul")

'effacement des donnŽes
i = ws.Cells(Rows.Count, 1).End(xlUp).Row
ws.Cells(1, 1).Resize(i, 5).ClearContents

'report des intitulŽs des variables souhaitŽes pour le tableau ˆ construire sur la feuille "calcul"
ws.Cells(1, 1).Resize(1, 5).Value = Array("id", "secteur", "rend", "port", "bench")

'rŽcupŽration des donnŽes (et report ˆ partir de la ligne 2) du tableau de la feuille "calcul" (et dont la date est d)
ws.Cells(2, 1).Resize(n, 1).Value = wsD.Range(wsD.Cells(ligne_debut, colId), wsD.Cells(ligne_fin, colId)).Value
ws.Cells(2, 2).Resize(n, 1).Value = wsD.Range(wsD.Cells(ligne_debut, colSect), wsD.Cells(ligne_fin, colSect)).Value
ws.Cells(2, 3).Resize(n, 1).Value = wsD.Range(wsD.Cells(ligne_debut, colRend), wsD.Cells(ligne_fin, colRend)).Value
ws.Cells(2, 4).Resize(n, 1).Value = wsD.Range(wsD.Cells(ligne_debut, colPort), wsD.Cells(ligne_fin, colPort)).Value
ws.Cells(2, 5).Resize(n, 1).Value = wsD.Range(wsD.Cells(ligne_debut, colBench), wsD.Cells(ligne_fin, colBench)).Value

End Sub

