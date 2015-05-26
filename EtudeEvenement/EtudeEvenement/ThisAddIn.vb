Public Class ThisAddIn

    Private Sub ThisAddIn_Startup() Handles Me.Startup

    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

    'Calcul la statistique de test et effectue le test
    'Renvoie true si l'hypothèse est rejetée
    Public Function ThisAddIn_MethodeTabCAR(tabAR As Double(,), seuil As Double) As Boolean
        Dim varCAR(tabAR.GetLength(1)) As Double                     'Variable aléatoire correspondant aux CAR

        'Calcul de la statistique pour chaque entreprise
        For colonne = 0 To tabAR.GetUpperBound(1)
            'Calcul de CAR
            Dim CAR As Double = 0
            For i = 0 To tabAR.GetUpperBound(0)
                CAR = CAR + tabAR(i, colonne)
            Next i
            Dim moyenne As Double = CAR / tabAR.GetLength(0)
            'Calcul de la variance des AR
            Dim variance As Double = 0
            For i = 0 To tabAR.GetUpperBound(0)
                Dim tmp As Double = tabAR(i, colonne) - moyenne
                variance = variance + tmp * tmp
            Next i
            variance = variance / tabAR.GetLength(0)
            varCAR(colonne) = CAR / Math.Sqrt((tabAR.GetLength(0) - 1) * variance)
        Next colonne

        'Test statistique
        ThisAddIn_MethodeTabCAR = test_student(varCAR, tabAR.GetLength(1), seuil)
    End Function

    Public Function ThisAddIn_ModeleMarche(fenetre As Integer, seuil As Double) As Boolean
        'on se positionne sur la feuille des Rt
        Application.Sheets("Rt").activate()
        Dim currentSheet As Excel.Worksheet = CType(Application.ActiveSheet, Excel.Worksheet)
        'compte le nombre de lignes et de colonnes
        Dim nbLignes As Integer = currentSheet.UsedRange.Rows.Count
        Dim nbColonnes As Integer = currentSheet.UsedRange.Columns.Count
        'tableau stockant les AR calculés grâce à la régression
        Dim tabAR(fenetre - 1, nbColonnes - 2) As Double
        'indice de ligne de la dernière ligne de l'ensemble d'apprentissage
        Dim dernLigne As Integer = nbLignes - fenetre

        For i = 0 To nbColonnes - 2
            Dim plageY As Excel.Range
            Dim plageX As Excel.Range
            plageY = Application.Range(Application.Cells(2, i + 2), Application.Cells(dernLigne, i + 2))
            'on se positionne sur la feuille des Rm pour récupérer plageX
            Application.Sheets("Rm").activate()
            plageX = Application.Range(Application.Cells(2, i + 2), Application.Cells(dernLigne, i + 2))
            'calcul des paramètres de la régression linéaire
            Dim beta As Double = Application.WorksheetFunction.Index(Application.WorksheetFunction.LinEst(plageY, plageX), 1)
            Dim alpha As Double = Application.WorksheetFunction.Index(Application.WorksheetFunction.LinEst(plageY, plageX), 2)

            'remplissage du tableau
            For t = 0 To fenetre - 1
                Dim k As Double = alpha + beta * Application.Cells(dernLigne + t + 1, i + 2).Value
                Application.Worksheets("Rt").activate()
                tabAR(t, i) = Application.Cells(dernLigne + t + 1, i + 2).Value - k
                Application.Worksheets("Rm").activate()
            Next
            'on retourne sur la feuille des Rt
            Application.Worksheets("Rt").activate()
        Next
        ThisAddIn_ModeleMarche = ThisAddIn_MethodeTabCAR(tabAR, seuil)
    End Function

    'Récupère les Rt et les Rm, puis appelle les calculs de statistique
    Public Function ThisAddIn_ModeleRentaMarche(fenetre As Integer, seuil As Double) As Boolean
        Application.Sheets("Rt").activate()
        Dim currentSheet As Excel.Worksheet = CType(Application.ActiveSheet, Excel.Worksheet)
        Dim nbLignes As Integer = currentSheet.UsedRange.Rows.Count 'Nombre de lignes
        Dim nbColonnes As Integer = currentSheet.UsedRange.Columns.Count 'Nombre de colonnes

        Dim Rt(nbLignes - 2, nbColonnes - 2) As Double
        Dim Rm(nbLignes - 2, nbColonnes - 2) As Double
        Dim AR(nbLignes - 2, nbColonnes - 2) As Double 'Le tableau des rentabilités anormales

        'La construction du vecteur Rt
        For ligne = 2 To nbLignes
            For colonne = 2 To nbColonnes
                Rt(ligne - 2, colonne - 2) = currentSheet.Application.Cells(ligne, colonne).Value
            Next colonne
        Next ligne

        'La construction du vecteur Rm
        Application.Sheets("Rm").activate()
        currentSheet = CType(Application.ActiveSheet, Excel.Worksheet)
        For ligne = 2 To nbLignes
            For colonne = 2 To nbColonnes
                Rm(ligne - 2, colonne - 2) = currentSheet.Application.Cells(ligne, colonne).Value
            Next colonne
        Next ligne

        'La construction du tableau des Rentabilités Anormales
        For ligne = 0 To nbLignes - 2
            For colonne = 0 To nbColonnes - 2
                AR(ligne, colonne) = Rt(ligne, colonne) - Rm(ligne, colonne)
            Next colonne
        Next ligne

        ThisAddIn_ModeleRentaMarche = ThisAddIn_MethodeTabCAR(AR, seuil)

    End Function

    'Calcul les Ki, puis effectue les tests statistiques sur (Ri - Ki)
    'Renvoie true si l'hypothèse est rejetée
    Public Function ThisAddIn_CalcNormMoy(fenetre As Integer, seuil As Double) As Boolean
        Application.Sheets("Rt").activate()
        Dim currentSheet As Excel.Worksheet = CType(Application.ActiveSheet, Excel.Worksheet)
        MsgBox(currentSheet.Name)
        Dim nbLignes As Integer = currentSheet.UsedRange.Rows.Count                'Nombre de lignes
        Dim nbColonnes As Integer = currentSheet.UsedRange.Columns.Count           'Nombre de colonnes
        Dim tabMoy(nbColonnes - 2) As Double                                      'Tableau des moyennes de chaque titre

        'Calcul des moyennes
        For colonne = 2 To nbColonnes
            Dim plage As Excel.Range = Application.Range(currentSheet.Application.Cells(2, colonne), currentSheet.Application.Cells(nbLignes - fenetre, colonne))
            tabMoy(colonne - 2) = Application.WorksheetFunction.Average(plage)
        Next colonne

        'Calcul des AR sur la fenêtre
        Dim tabAR(fenetre - 1, nbColonnes - 2) As Double                          'Tableau des AR sur la fenêtre de l'événement
        Dim debFenetre As Integer = nbLignes - fenetre + 1
        For colonne = 2 To nbColonnes
            For indDate = debFenetre To nbLignes
                tabAR(indDate - debFenetre, colonne - 2) = currentSheet.Cells(indDate, colonne).Value - tabMoy(colonne - 2)
            Next indDate
        Next colonne
        ThisAddIn_CalcNormMoy = ThisAddIn_MethodeTabCAR(tabAR, seuil)
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
            varCAR(colonne - 2) = CAR / Math.Sqrt((nbLignes - 1) * variance)
        Next colonne

        'Test statistique
        ThisAddIn_MethodeCAR = test_student(varCAR, nbColonnes, seuil)
    End Function

    'Renvoie true si l'hypothèse H0 est rejetée
    Private Function test_student(tabCAR() As Double, tailleTabCAR As Integer, seuil As Double) As Boolean
        test_student = False
        Dim testHypothese As Double
        Dim moyenneTab As Double = calcul_moyenne(tabCAR)
        Dim varianceTab As Double = calcul_variance(tabCAR, moyenneTab)
        testHypothese = Math.Abs(Math.Sqrt(tailleTabCAR) * moyenneTab / varianceTab)
        If testHypothese > Application.WorksheetFunction.TInv(seuil, tailleTabCAR) Then
            test_student = True
        End If
    End Function

    Private Function calcul_moyenne(tab() As Double) As Double
        For i = 0 To tab.GetUpperBound(0)
            calcul_moyenne = calcul_moyenne + tab(i)
        Next i
        calcul_moyenne = calcul_moyenne / (tab.GetLength(0))
    End Function

    Private Function calcul_variance(tab() As Double, moyenne As Double) As Double
        For i = 0 To tab.GetUpperBound(0)
            Dim tmp As Double = tab(i) - moyenne
            calcul_variance = calcul_variance + tmp * tmp
        Next i
        calcul_variance = calcul_variance / (tab.GetLength(0))
    End Function

End Class
