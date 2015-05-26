Public Class ThisAddIn

    Private Sub ThisAddIn_Startup() Handles Me.Startup

    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

    'Calcul la statistique de test et effectue le test
    'Renvoie true si l'hypothèse est rejetée
    Public Function ThisAddIn_MethodeTabCAR(tabAR As Double(,), seuil As Double, fenetreDebut As Integer, fenetreFin As Integer) As Boolean
        Dim varCAR(tabAR.GetUpperBound(1)) As Double                     'Variable aléatoire correspondant aux CAR
        Dim currentSheet As Excel.Worksheet = CType(Globals.ThisAddIn.Application.Worksheets("Rt"), Excel.Worksheet)
        Dim indDebFenetre As Integer = 2 + fenetreDebut - currentSheet.Cells(2, 1).Value
        Dim indFinFenetre As Integer = 2 + fenetreFin - currentSheet.Cells(2, 1).Value
        Dim tailleFenetre As Integer = fenetreFin - fenetreDebut + 1

        'Calcul de la statistique pour chaque entreprise
        For colonne = 0 To tabAR.GetUpperBound(1)
            'Calcul de CAR sur la fenetre d'événement
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
            varCAR(colonne) = CAR / Math.Sqrt(tailleFenetre * variance)
        Next colonne

        'Test statistique
        ThisAddIn_MethodeTabCAR = test_student(varCAR, tabAR.GetLength(1), seuil)
    End Function

    'Estimation des AR à partir du modèle de marché : K = alpha + beta*Rm
    Public Function ThisAddIn_ModeleMarche(fenetreDebut As Integer, fenetreFin As Integer, seuil As Double) As Boolean
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
        ThisAddIn_ModeleMarche = ThisAddIn_MethodeTabCAR(tabAR, seuil, fenetreDebut, fenetreFin)
    End Function

    'Calcule les AR pour chaque titre puis appelle les calculs de statistique
    Public Function ThisAddIn_ModeleRentaMarche(fenetreDebut As Integer, fenetreFin As Integer, seuil As Double) As Boolean
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

        ThisAddIn_ModeleRentaMarche = ThisAddIn_MethodeTabCAR(tabAR, seuil, fenetreDebut, fenetreFin)
    End Function

    'Calcul les Ki, puis effectue les tests statistiques sur (Ri - Ki)
    'Renvoie true si l'hypothèse est rejetée
    Public Function ThisAddIn_CalcNormMoy(fenetreDebut As Integer, fenetreFin As Integer, seuil As Double) As Boolean
        Dim currentSheet As Excel.Worksheet = CType(Application.Worksheets("Rt"), Excel.Worksheet)
        Dim nbLignes As Integer = currentSheet.UsedRange.Rows.Count                'Nombre de lignes
        Dim nbColonnes As Integer = currentSheet.UsedRange.Columns.Count           'Nombre de colonnes
        Dim tabMoy(nbColonnes - 2) As Double                                       'Tableau des moyennes de chaque titre
        Dim indDebFenetre As Integer = 2 + fenetreDebut - currentSheet.Cells(2, 1).Value
        Dim indFinFenetre As Integer = 2 + fenetreFin - currentSheet.Cells(2, 1).Value
        Dim tailleFenetre As Integer = fenetreFin - fenetreDebut + 1

        'Calcul des moyennes
        For colonne = 2 To nbColonnes
            Dim plage As Excel.Range = Application.Range(currentSheet.Cells(2, colonne), currentSheet.Cells(indDebFenetre - 1, colonne))
            tabMoy(colonne - 2) = Application.WorksheetFunction.Average(plage)
            'On fait également le calcul sur la période après l'événement
            'If indFinFenetre < nbLignes Then
            '    plage = Application.Range(currentSheet.Cells(indFinFenetre + 1, colonne), currentSheet.Cells(nbLignes, colonne))
            '    tabMoy(colonne - 2) = tabMoy(colonne - 2) + Application.WorksheetFunction.Average(plage)
            'End If
        Next colonne

        'Calcul des AR sur la fenêtre
        Dim tabAR(nbLignes - 2, nbColonnes - 2) As Double                          'Tableau des AR sur la fenêtre de l'événement
        For colonne = 2 To nbColonnes
            For indDate = 2 To nbLignes
                tabAR(indDate - 2, colonne - 2) = currentSheet.Cells(indDate, colonne).Value - tabMoy(colonne - 2)
            Next indDate
        Next colonne
        ThisAddIn_CalcNormMoy = ThisAddIn_MethodeTabCAR(tabAR, seuil, fenetreDebut, fenetreFin)
    End Function

    'Renvoie true si l'hypothèse est rejetée
    Public Function ThisAddIn_MethodeCAR(seuil As Double) As Boolean
        Dim activeSheet As Excel.Worksheet = CType(Application.ActiveSheet, Excel.Worksheet)
        Dim nbLignes As Integer = activeSheet.UsedRange.Rows.Count                'Nombre de lignes
        Dim nbColonnes As Integer = activeSheet.UsedRange.Columns.Count           'Nombre de colonnes
        Dim varCAR(nbColonnes - 2) As Double                                      'Variable aléatoire correspondant aux CAR

        'Calcul de la statistique pour chaque entreprise
        For colonne = 2 To nbColonnes
            Dim plage As Excel.Range = Application.Range(Application.Cells(2, colonne), Application.Cells(nbLignes, colonne))
            Dim CAR As Double = Application.WorksheetFunction.Sum(plage)
            Dim variance As Double = Application.WorksheetFunction.Var(plage)
            varCAR(colonne - 2) = CAR / Math.Sqrt(variance * (nbLignes - 1))
        Next colonne

        'Test statistique
        ThisAddIn_MethodeCAR = test_student(varCAR, nbColonnes - 1, seuil)
    End Function

    Public Function ThisAddIn_PValeur(modele As Integer, Optional fenetreDebut As Integer = 0, Optional fenetreFin As Integer = 0) As Double
        Dim borneInf As Double = 0
        Dim borneSup As Double = 1
        Dim alpha As Double = (borneInf + borneSup) / 2
        While borneSup - borneInf > 0.0001
            Dim rejet As Boolean
            Select Case modele
                Case 0
                    rejet = ThisAddIn_CalcNormMoy(fenetreDebut, fenetreFin, alpha)
                Case 1
                    rejet = ThisAddIn_ModeleRentaMarche(fenetreDebut, fenetreFin, alpha)
                Case 2
                    rejet = ThisAddIn_ModeleMarche(fenetreDebut, fenetreFin, alpha)
                Case 3
                    rejet = ThisAddIn_MethodeCAR(alpha)
                Case Else
                    MsgBox("Erreur interne : Provient de ThisAddIn_ThinAddIn_PValeur", 16)
            End Select
            If rejet Then
                borneSup = alpha
            Else
                borneInf = alpha
            End If
            alpha = (borneInf + borneSup) / 2
        End While
        Return alpha
    End Function

    'Renvoie true si l'hypothèse H0 est rejetée
    Private Function test_student(tabCAR() As Double, tailleTabCAR As Integer, seuil As Double) As Boolean
        test_student = False
        Dim testHypothese As Double
        Dim moyenneTab As Double = calcul_moyenne(tabCAR)
        Dim varianceTab As Double = calcul_variance(tabCAR, moyenneTab)
        testHypothese = Math.Abs(Math.Sqrt(tailleTabCAR) * moyenneTab / varianceTab)
        If testHypothese > Application.WorksheetFunction.TInv(seuil, tailleTabCAR - 1) Then
            test_student = True
        End If
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
