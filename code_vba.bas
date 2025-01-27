Attribute VB_Name = "recuperation_donnees"
Option Explicit
Option Base 0
Sub procRecupData()

'proc�dure permettant de choisir une base de donn�e et y r�cup�rant les informations sur les dates, les champs disponibles

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'% D�claration des variables %
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

Dim wbD As Workbook 'fichier des donn�es
Dim wsD As Worksheet 'feuille utilis�e dans le fichier des donn�es
Dim adresse As String 'fullname du fichier des donn�es
Dim nFields As Integer 'nbre de champs de la base utilis�e
Dim nRecords As Long 'nbre d'enregistrements de la base utilis�e
Dim colDates As Integer 'colonne des dates
Dim d As Long 'date recherch�e (en entier)

Dim ws As Worksheet 'feuille de calcul (quelconque)
Dim message As String 'message
Dim i As Long, n As Long 'compteurs de ligne
Dim j As Integer 'compteurs de colonnes

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'% I. Ouverture du fichier des donn�es %
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


's�lection du fichier (par l'utilisateur)
adresse = Application.GetOpenFilename

'affectation de wb au fichier choisi par l'utilisateur
Set wbD = Workbooks.Open(adresse)


'si plusieurs feuilles, choix de la feuille � utiliser par l'utilisateur via une inputbox apr�s la r�cup�ration des noms des feui
message = wbD.Worksheets(1).Name
If wbD.Worksheets.Count > 1 Then

    'r�cup�ration des noms des feuilles (tables)
    message = "Le fichier s�lectionn� contient les feuilles suivantes :" & vbCrLf
    For Each ws In wbD.Worksheets
        message = ws.Name & vbCrLf
    Next ws
    message = message & vbCrLf & vbCrLf & "Quelle feuille doit �tre utulis�e?"

    message = InputBox(message, "Choix de la feuille.", wbD.Worksheets(1).Name)

End If
Set wsD = wbD.Worksheets(message)

'd�nombrement des champs et des enregistrements
nFields = wsD.Cells(1, wsD.Columns.Count).End(xlToLeft).Column
nRecords = wsD.Cells(wsD.Rows.Count, 1).End(xlUp).Row - 1

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'% II. R�cup�ration des champs %
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

'affectation de ws � l'unique feuille de ce fichier
Set ws = ThisWorkbook.Worksheets("intro")

'report sur l'unique feuille de ce fichier du nom du fichier de donn�es et de la feuille
ws.Range("fichier").Value = wbD.Name
ws.Range("feuille").Value = message

'effacement des champs sur la feuille du fichier et report des valeurs du fichier de donn�es
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

'cas o� un champ de date a �t� trouv�
If colDates < nFields + 1 Then
    
    'mise des dates en format num�rique
    wsD.Columns(colDates).Cells.NumberFormat = "General"
    'tri en fonction des dates
    wsD.Cells(2, 1).Resize(nRecords, nFields).Sort Key1:=wsD.Cells(1, colDates), Order1:=xlAscending  'xlDescending
    'boucle sur les lignes et report des dates diff�rentes
    j = 1
    i = 2
    Do While i <= nRecords + 1
        'r�cup�ration de la date
        d = wsD.Cells(i, colDates).Value
        'calcul de dates identiques � celle de la ligne i
        n = WorksheetFunction.CountIf(wsD.Cells(2, colDates).Resize(nRecords, 1), "=" & d)
        'report de la date
        j = j + 1
        ws.Range("dates").Cells(j, 1).Value = wsD.Cells(i, colDates).Value
        'saut � la ligne de la date suivante
        i = i + n
    Loop
    wsD.Columns(colDates).NumberFormat = "dd/mm/yyyy"
End If
ws.Range("dates").Resize(j, 1).NumberFormat = "dd/mm/yyyy"

End Sub
Sub imp_data_attribution()


'proc�dure r�cup�rant les donn�es pour l'attribution de performance � la Brinson en s'appuyant sur les donn�es de la feuille "intro" de _
ce ficher sur les donn�es � utiliser (dans les cellules "fichier" et "feuille") pour la variable d�finie (dans la cellule "cat�gorie") _
et pour la date souhait� (celules "date").


'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'% D�claration des variables %
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

Dim wbD As Workbook 'le fichier de donn�es
Dim wsD As Worksheet 'la feuille des donn�es
Dim nFields As Integer 'nbre de champs de la base utilis�e
Dim nRecords As Long 'nbre d'enregistrements de la base utilis�e

Dim colDates As Integer 'colonne des dates
Dim colSect As Integer 'colonne des secteurs
Dim colRend As Integer 'colonne des rendements
Dim colPort As Integer 'colonne du portefeuille
Dim colBench As Integer 'colonne du benchmark
Dim colId As Integer 'colonne des identifiants

Dim d As Long 'date recherch�e (en entier)
Dim ligne_debut As Long 'premi�re ligne contenant la date recherch�e
Dim ligne_fin As Long 'derni�re ligne contenant la date recherch�e



Dim ws As Worksheet 'feuille de calcul (quelconque)
Dim message As String 'message
Dim i As Long, n As Long 'compteurs de ligne
Dim j As Integer 'compteurs de colonnes


'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'% I. Prise en main des donn�es %
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


'affectatationde ws � l'unique feuille de ce fichier
Set ws = ThisWorkbook.Worksheets("intro")

'test sur l'ouverture u fichier de donn�es
On Error Resume Next
Workbooks(ws.Range("fichier").Value).Activate
If Err.Number > 0 Then MsgBox "Le fichier de donn�es doit �tre d�j� ouvert." & vbCrLf & "La proc�dure va s'arr�ter." & vbCrLf & _
"Ouver le fichier et relancer apr�s la proc�dure.": Exit Sub
On Error GoTo 0

'affectation de wsD � la feuille de donn�es d�finies sur l'unique feuille de ce fichier (cellules "fichier" et feuille")
Set wsD = Workbooks(ws.Range("fichier").Value).Worksheets(ws.Range("feuille").Value)

'd�nombrement des champs et des enregistrements

nFields = wsD.Cells(1, wsD.Columns.Count).End(xlToLeft).Column
nRecords = wsD.Cells(Rows.Count, 1).End(xlUp).Row - 1

'recherche des positions des champs (variables colRend, colSect, colPort, colBench, colDates, colId � d�finir avec m�thode find appliqu�e � la ligne 1)
'<ligne 1 de la feuille>.Find(what:=<chaine de caract�re>, LookAt:=xlwhole).column

colRend = Rows(1).Find(what:="return", LookAt:=xlWhole).Column
colSect = Rows(1).Find(what:="sector", LookAt:=xlWhole).Column
colPort = Rows(1).Find(what:="portfolio", LookAt:=xlWhole).Column
colBench = Rows(1).Find(what:="benchmark", LookAt:=xlWhole).Column
colDates = Rows(1).Find(what:="date", LookAt:=xlWhole).Column
colId = Rows(1).Find(what:="id", LookAt:=xlWhole).Column



'r�cup�ration de la date (en entier)
d = ws.Range("date").Value

'tri des donn�es par rapport aux dates (apr�s les avoir mis au format "General")
'<plage>.Sort Key1:=<cellule ligne 1 contenant l'intitul�>, Order1:=xlAscending
wsD.Cells(2, 1).Resize(nRecords, nFields).Sort Key1:=wsD.Cells(1, colDates), Order1:=xlAscending


'recherche de la premi�re et de la derni�re ligne des enregistements dont la date est d

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

'affectation de ws � la feuille "calcul"
Set ws = ThisWorkbook.Worksheets("calcul")

'effacement des donn�es
i = ws.Cells(Rows.Count, 1).End(xlUp).Row
ws.Cells(1, 1).Resize(i, 5).ClearContents

'report des intitul�s des variables souhait�es pour le tableau � construire sur la feuille "calcul"
ws.Cells(1, 1).Resize(1, 5).Value = Array("id", "secteur", "rend", "port", "bench")

'r�cup�ration des donn�es (et report � partir de la ligne 2) du tableau de la feuille "calcul" (et dont la date est d)
ws.Cells(2, 1).Resize(n, 1).Value = wsD.Range(wsD.Cells(ligne_debut, colId), wsD.Cells(ligne_fin, colId)).Value
ws.Cells(2, 2).Resize(n, 1).Value = wsD.Range(wsD.Cells(ligne_debut, colSect), wsD.Cells(ligne_fin, colSect)).Value
ws.Cells(2, 3).Resize(n, 1).Value = wsD.Range(wsD.Cells(ligne_debut, colRend), wsD.Cells(ligne_fin, colRend)).Value
ws.Cells(2, 4).Resize(n, 1).Value = wsD.Range(wsD.Cells(ligne_debut, colPort), wsD.Cells(ligne_fin, colPort)).Value
ws.Cells(2, 5).Resize(n, 1).Value = wsD.Range(wsD.Cells(ligne_debut, colBench), wsD.Cells(ligne_fin, colBench)).Value

End Sub

