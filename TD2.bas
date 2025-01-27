Attribute VB_Name = "attribution_performances"
Option Explicit
Option Base 1
Sub calc_brinson()

'Les donn�es sont suppos�es avoir �t� d�pos�es sur la feuille "calcul" de ce fichier et �tre compos�es de 4 champs (Id, secteur, rendements, _
poids dans le portefeuille, poids dans le benchmark). Les r�sultats sont report�s progressivement sur la feuille "attribution".

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'% D�claration des variables %
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

Dim nFields As Integer 'nbre de champs de la base utilis�e
Dim nRecords As Long 'nbre d'enregistrements de la base utilis�e

Dim debut As Long 'premier enregistrement (apr�s tri) contenant le secteur recherch�
Dim fin As Long 'dernier enregistrement (apr�s tri) contenant le secteur recherch�


Dim wsCalc As Worksheet 'feuille "calcul"
Dim wsAttr As Worksheet 'feuille "attribution"
Dim rg As Range 'place des rendements

Dim i As Long, k As Long, n As Long 'compteurs de ligne
Dim j As Integer 'compteurs de colonnes


'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'% I. Prise en main des donn�es %
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


'affectatation de wsCalc et wsAttr
Set wsCalc = ThisWorkbook.Worksheets("calcul")
Set wsAttr = ThisWorkbook.Worksheets("attribution")
'd�nombrement d'enregistrements
nRecords = wsCalc.Cells(Rows.Count, 1).End(xlUp).Row - 1


'tri des donn�es par rapport aux secteurs
wsCalc.Cells(2, 1).Resize(nRecords, 5).Sort Key1:=wsCalc.Cells(1, 2), Order1:=xlAscending
 
'mise en forme de la feuille "attribution" (effacement des donn�es pr�c�dentes, report en ligne 1 des intitul�s)
With wsAttr
    .Cells.ClearContents
    .Cells(1, 1).Resize(1, 8).Value = Array("secteur", "xp", "xb", "rp", "rb", "selection", "timing", "interaction")
    .Rows(1).Font.Bold = True
End With

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'% II. Attributon de performance %
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

'position (ligne) initiale sur la feuille "calcul"
i = 2
'position initiale sur la feuille "attribution"
k = 1
'd�but de la boucle sur les lignes de la feuille "calcul"
Do While wsCalc.Cells(i + 1, 1).Value <> ""

    'report du secteur � la ligne k(apr�s incr�mentation du compteur k)
    k = k + 1
    wsAttr.Cells(k, 1).Value = wsCalc.Cells(i, 2).Value
    'calcul du nombre (n) de lignes occup�es par le secteur (avec countif)
    n = Application.WorksheetFunction.CountIf(wsCalc.Cells(1, 2).Resize(nRecords, 1), "=" & wsCalc.Cells(i, 2).Value)
    'affectation � rg de la plage du secteur
    Set rg = wsCalc.Cells(i, 1).Resize(n, 5)
    'calcul et report des poids et des rendements
    wsAttr.Cells(k, 2).Value = Application.WorksheetFunction.Sum(rg.Columns(4))
    wsAttr.Cells(k, 3).Value = Application.WorksheetFunction.Sum(rg.Columns(5))
    wsAttr.Cells(k, 4).Value = Application.WorksheetFunction.SumProduct(rg.Columns(4), rg.Columns(3))
    wsAttr.Cells(k, 5).Value = Application.WorksheetFunction.SumProduct(rg.Columns(5), rg.Columns(3))
    
    'calcul de la contribution de la s�lection, du market timing, du terme d'interaction
    wsAttr.Cells(k, 7).Value = (wsAttr.Cells(k, 2).Value - wsAttr.Cells(k, 3).Value) * wsAttr.Cells(k, 5).Value
    wsAttr.Cells(k, 6).Value = (wsAttr.Cells(k, 4).Value - wsAttr.Cells(k, 5).Value) * wsAttr.Cells(k, 3).Value
    wsAttr.Cells(k, 8).Value = (wsAttr.Cells(k, 4).Value - wsAttr.Cells(k, 5).Value) * (wsAttr.Cells(k, 2).Value - wsAttr.Cells(k, 3).Value)
    'ajustement de i (saut de n lignes)

    i = i + n
    
Loop

'calcul de l'attribution de performance au niveau du portefeuille (et mise en forme (ligne 1 et du portefeuille en gras, centrage horizontal, _
donn�es num�riques en % avec 1 d�cimale)

End Sub
