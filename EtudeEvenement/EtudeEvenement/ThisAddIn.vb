﻿Imports System.Windows.Forms.DataVisualization.Charting
Imports System.Diagnostics

Public Class ThisAddIn

    'Calcule les AR avec le modèle considéré
    Public Function calculAR(fenetreDebut As Integer) As Double(,)
        'appelle une fonction pour chaque modèle
        Select Case Globals.Ribbons.Ruban.choixSeuilFenetre.modele
            Case 0
                calculAR = modeleMoyenne(fenetreDebut)
            Case 1
                calculAR = modeleMarcheSimple()
            Case 2
                calculAR = modeleMarche(fenetreDebut)
            Case Else
                MsgBox("Erreur interne : numero de modèle incorrect dans ChoixSeuilFenetre", 16)
                calculAR = Nothing
        End Select
    End Function

    'Calcule les CAR "normalisés" pour le test statistique
    Public Function calculCAR(tabAR As Double(,), fenetreDebut As Integer, fenetreFin As Integer) As Double()
        Dim normCar(tabAR.GetUpperBound(1)) As Double   'Variable aléatoire correspondant aux CAR "normalisés"
        Dim currentSheet As Excel.Worksheet = CType(Globals.ThisAddIn.Application.Worksheets("Rt"), Excel.Worksheet)
        Dim indDebFenetre As Integer = 2 + fenetreDebut - currentSheet.Cells(2, 1).Value
        Dim indFinFenetre As Integer = 2 + fenetreFin - currentSheet.Cells(2, 1).Value
        Dim tailleFenetre As Integer = fenetreFin - fenetreDebut + 1

        'Calcul de la statistique pour chaque entreprise
        For colonne = 0 To tabAR.GetUpperBound(1)
            'Calcul de CAR sur la fenetre d'événement paramétrée
            Dim CAR As Double = 0
            For i = indDebFenetre - 2 To indFinFenetre - 2
                CAR = CAR + tabAR(i, colonne)
            Next i
            Dim moyenne As Double = 0
            For i = 0 To indDebFenetre - 1 - 2
                moyenne = moyenne + tabAR(i, colonne)
            Next i
            moyenne = moyenne / (indDebFenetre - 2)
            'Calcul de la variance des AR sur la période d'estimation
            Dim variance As Double = 0
            For i = 0 To indDebFenetre - 1 - 2
                Dim tmp As Double = tabAR(i, colonne) - moyenne
                variance = variance + tmp * tmp
            Next i
            variance = variance / (indDebFenetre - 3)
            normCar(colonne) = CAR / Math.Sqrt(tailleFenetre * variance)
            'Debug.WriteLine(normCar(colonne))
        Next colonne
        'retourne le tableau des CAR normalisés
        calculCAR = normCar
    End Function

    ''Calcul la statistique de test et effectue le test
    ''Renvoie true si l'hypothèse est rejetée
    'Public Function ThisAddIn_MethodeTabCAR(tabAR As Double(,), fenetreDebut As Integer, fenetreFin As Integer) As Double
    '    Dim varCAR(tabAR.GetUpperBound(1)) As Double                     'Variable aléatoire correspondant aux CAR
    '    Dim currentSheet As Excel.Worksheet = CType(Globals.ThisAddIn.Application.Worksheets("Rt"), Excel.Worksheet)
    '    Dim indDebFenetre As Integer = 2 + fenetreDebut - currentSheet.Cells(2, 1).Value
    '    Dim indFinFenetre As Integer = 2 + fenetreFin - currentSheet.Cells(2, 1).Value
    '    Dim tailleFenetre As Integer = fenetreFin - fenetreDebut + 1

    '    'Calcul de la statistique pour chaque entreprise
    '    For colonne = 0 To tabAR.GetUpperBound(1)
    '        'Calcul de CAR sur la fenetre d'événement
    '        Dim CAR As Double = 0
    '        For i = indDebFenetre - 2 To indFinFenetre - 2
    '            CAR = CAR + tabAR(i, colonne)
    '        Next i
    '        Dim moyenne As Double = 0
    '        For i = 0 To indDebFenetre - 1 - 2
    '            moyenne = moyenne + tabAR(i, colonne)
    '        Next i
    '        moyenne = moyenne / (indDebFenetre - 2)
    '        'Calcul de la variance des AR sur la période d'estimation
    '        Dim variance As Double = 0
    '        For i = 0 To indDebFenetre - 1 - 2
    '            Dim tmp As Double = tabAR(i, colonne) - moyenne
    '            variance = variance + tmp * tmp
    '        Next i
    '        variance = variance / (indDebFenetre - 3)
    '        varCAR(colonne) = CAR / Math.Sqrt(tailleFenetre * variance)
    '    Next colonne

    '    'Test statistique
    '    ThisAddIn_MethodeTabCAR = calculStatistique(varCAR, tabAR.GetLength(1))
    'End Function

    'Estimation des AR à partir du modèle de marché : K = alpha + beta*Rm
    Public Function modeleMarche(fenetreDebut As Integer) As Double(,)
        'on se positionne sur la feuille des Rt
        Dim currentSheet As Excel.Worksheet = CType(Application.Worksheets("Rt"), Excel.Worksheet)
        'compte le nombre de lignes et de colonnes
        Dim nbLignes As Integer = currentSheet.UsedRange.Rows.Count
        Dim nbColonnes As Integer = currentSheet.UsedRange.Columns.Count
        'tableau stockant les AR calculés grâce à la régression
        Dim tabAR(nbLignes - 2, nbColonnes - 2) As Double
        'indice de ligne de la dernière ligne de l'ensemble d'apprentissage
        Dim dernLigne As Integer = 1 + fenetreDebut - currentSheet.Cells(2, 1).Value

        For i = 0 To nbColonnes - 2
            Dim plageY As Excel.Range
            Dim plageX As Excel.Range
            plageY = Application.Range(currentSheet.Cells(2, i + 2), currentSheet.Cells(dernLigne, i + 2))
            'on se positionne sur la feuille des Rm pour récupérer plageX
            currentSheet = CType(Application.Worksheets("Rm"), Excel.Worksheet)
            plageX = Application.Range(currentSheet.Cells(2, i + 2), currentSheet.Cells(dernLigne, i + 2))
            'calcul des paramètres de la régression linéaire
            Dim beta As Double = Application.WorksheetFunction.Index(Application.WorksheetFunction.LinEst(plageY, plageX), 1)
            Dim alpha As Double = Application.WorksheetFunction.Index(Application.WorksheetFunction.LinEst(plageY, plageX), 2)

            'remplissage du tableau
            For t = 0 To nbLignes - 2
                Dim k As Double = alpha + beta * currentSheet.Cells(t + 2, i + 2).Value
                currentSheet = CType(Application.Worksheets("Rt"), Excel.Worksheet)
                tabAR(t, i) = currentSheet.Cells(t + 2, i + 2).Value - k
                currentSheet = CType(Application.Worksheets("Rm"), Excel.Worksheet)
            Next
            'on retourne sur la feuille des Rt
            currentSheet = CType(Application.Worksheets("Rt"), Excel.Worksheet)
        Next
        modeleMarche = tabAR
    End Function

    'Estimation des AR à partir du modèle de marché simplifié : K = moyenne des rentabilités
    Public Function modeleMarcheSimple() As Double(,)
        Dim currentSheet As Excel.Worksheet = CType(Application.Worksheets("Rt"), Excel.Worksheet)
        'compte le nombre de lignes et de colonnes
        Dim nbLignes As Integer = currentSheet.UsedRange.Rows.Count
        Dim nbColonnes As Integer = currentSheet.UsedRange.Columns.Count
        'tableau stockant les AR calculés grâce à la régression
        Dim tabAR(nbLignes - 2, nbColonnes - 2) As Double

        For i = 0 To nbColonnes - 2
            'remplissage du tableau
            For t = 0 To nbLignes - 2
                currentSheet = CType(Application.Worksheets("Rm"), Excel.Worksheet)
                Dim k As Double = currentSheet.Cells(t + 2, i + 2).Value
                currentSheet = CType(Application.Worksheets("Rt"), Excel.Worksheet)
                tabAR(t, i) = currentSheet.Cells(t + 2, i + 2).Value - k
            Next
        Next
        modeleMarcheSimple = tabAR
    End Function

    'Estimation des AR à partir du modèle de la moyenne : K = R
    Public Function modeleMoyenne(fenetreDebut As Integer) As Double(,)
        Dim currentSheet As Excel.Worksheet = CType(Application.Worksheets("Rt"), Excel.Worksheet)
        'compte le nombre de lignes et de colonnes
        Dim nbLignes As Integer = currentSheet.UsedRange.Rows.Count
        Dim nbColonnes As Integer = currentSheet.UsedRange.Columns.Count
        'Tableau des moyennes de chaque titre
        Dim tabMoy(nbColonnes - 2) As Double
        'indice de début de la fenêtre
        Dim indDebFenetre As Integer = 2 + fenetreDebut - currentSheet.Cells(2, 1).Value

        'Calcul des moyennes
        For colonne = 2 To nbColonnes
            Dim plage As Excel.Range = Application.Range(currentSheet.Cells(2, colonne), currentSheet.Cells(indDebFenetre - 1, colonne))
            tabMoy(colonne - 2) = Application.WorksheetFunction.Average(plage)
        Next colonne

        'Calcul des AR sur la fenêtre
        Dim tabAR(nbLignes - 2, nbColonnes - 2) As Double                          'Tableau des AR sur la fenêtre de l'événement
        For colonne = 2 To nbColonnes
            For indDate = 2 To nbLignes
                tabAR(indDate - 2, colonne - 2) = currentSheet.Cells(indDate, colonne).Value - tabMoy(colonne - 2)
            Next indDate
        Next colonne
        modeleMoyenne = tabAR
    End Function

    Public Function calculPValeur(tailleEchant As Integer, testHyp As Double) As Double
        Return Application.WorksheetFunction.T_Dist_2T(testHyp, tailleEchant - 1)
    End Function

    Public Sub tracerPValeur(tailleEchant As Integer, maxFenetre As Integer)
        For i = 0 To maxFenetre
            Dim tabAR As Double(,)
            tabAR = Globals.ThisAddIn.calculAR(-i, i)
            Dim tabCAR As Double()
            tabCAR = Globals.ThisAddIn.calculCAR(tabAR, -i, i)
            Dim testHyp As Double = Globals.ThisAddIn.calculStatistique(tabCAR)
            Dim pValeur As Double = Globals.ThisAddIn.calculPValeur(tailleEchant, testHyp) * 100

            Dim p As New DataPoint
            p.XValue = i
            p.YValues = {pValeur.ToString("0.00000")}

            Globals.Ribbons.Ruban.graphPVal.GraphiqueChart.Series("Series1").Points.Add(p)
        Next i
    End Sub

    'Renvoie true si l'hypothèse H0 est rejetée
    Public Function calculStatistique(tabCAR() As Double) As Double
        Dim tailleTabCAR As Integer = tabCAR.GetLength(0)
        Dim moyenneTab As Double = calcul_moyenne(tabCAR)
        Dim varianceTab As Double = calcul_variance(tabCAR, moyenneTab)
        calculStatistique = Math.Abs(Math.Sqrt(tailleTabCAR) * moyenneTab / Math.Sqrt(varianceTab))
    End Function

    Private Function calcul_moyenne(tab() As Double) As Double
        calcul_moyenne = 0
        For i = 0 To tab.GetUpperBound(0)
            calcul_moyenne = calcul_moyenne + tab(i)
        Next i
        calcul_moyenne = calcul_moyenne / (tab.GetLength(0))
    End Function

    Private Function calcul_variance(tab() As Double, moyenne As Double) As Double
        calcul_variance = 0
        For i = 0 To tab.GetUpperBound(0)
            Dim tmp As Double = tab(i) - moyenne
            calcul_variance = calcul_variance + tmp * tmp
        Next i
        calcul_variance = calcul_variance / (tab.GetLength(0) - 1)
    End Function

End Class
