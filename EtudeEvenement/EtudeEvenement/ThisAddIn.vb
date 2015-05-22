Public Class ThisAddIn

    Private Sub ThisAddIn_Startup() Handles Me.Startup

    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

    'Renvoie true si l'hypothèse est rejetée
    Public Function ThisAddIn_MethodeCAR(seuil) As Boolean
        Dim activeSheet As Excel.Worksheet = CType(Application.ActiveSheet, Excel.Worksheet)
        Dim nbLignes As Integer = activeSheet.UsedRange.Rows.Count                'Nombre de lignes
        Dim nbColonnes As Integer = activeSheet.UsedRange.Columns.Count           'Nombre de colonnes
        Dim varCAR(nbColonnes - 1) As Double                                      'Variable aléatoire correspondant aux CAR

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
        For i = 0 To UBound(tab)
            calcul_moyenne = calcul_moyenne + tab(i)
        Next i
        calcul_moyenne = calcul_moyenne / (UBound(tab) - 1)
    End Function

    Private Function calcul_variance(tab() As Double, moyenne As Double) As Double
        For i = 0 To UBound(tab)
            Dim tmp As Double = tab(i) - moyenne
            calcul_variance = calcul_variance + tmp * tmp
        Next i
        calcul_variance = calcul_variance / (UBound(tab) - 1)
    End Function

End Class
