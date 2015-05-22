Public Class ThisAddIn

    Private Sub ThisAddIn_Startup() Handles Me.Startup

    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

    Public Sub ThisAddIn_Methode_CAR()
        Dim activeSheet As Excel.Worksheet = CType(Application.ActiveSheet, Excel.Worksheet)
        Dim nbLignes As Integer = activeSheet.UsedRange.Rows.Count                'Nombre de lignes
        Dim nbColonnes As Integer = activeSheet.UsedRange.Columns.Count           'Nombre de colonnes
        Dim VarCAR(nbColonnes - 1) As Double                                      'Variable aléatoire correspondant aux CAR

        'Calcul de la statistique pour chaque entreprise
        For colonne = 2 To nbColonnes
            Dim plage As Excel.Range = Application.Range(Application.Cells(2, colonne), Application.Cells(nbLignes, colonne))
            Dim CAR As Double = Application.WorksheetFunction.Sum(plage)
            Dim variance As Double = Application.WorksheetFunction.Var(plage)
            VarCAR(colonne - 2) = CAR / Math.Sqrt((nbLignes - 1) * variance)
        Next colonne

        'Test statistique
        Dim testHypothese As Double
        Dim seuil As Double = 0.05
        Dim moyenneTab As Double = calcul_moyenne(VarCAR)
        Dim varianceTab As Double = calcul_variance(VarCAR, moyenneTab)
        testHypothese = Math.Abs(Math.Sqrt(nbColonnes) * moyenneTab / varianceTab)
        If testHypothese > Application.WorksheetFunction.TInv(seuil, nbColonnes) Then
            MsgBox("Rejet de l'hypothèse")
        Else
            MsgBox("Non rejet de l'hypothèse")
        End If
    End Sub

    Private Function calcul_moyenne(tab) As Double
        For i = 0 To UBound(tab)
            calcul_moyenne = calcul_moyenne + tab(i)
        Next i
        calcul_moyenne = calcul_moyenne / (UBound(tab) - 1)
    End Function

    Private Function calcul_variance(tab, moyenne) As Double
        For i = 0 To UBound(tab)
            Dim tmp As Double = tab(i) - moyenne
            calcul_variance = calcul_variance + tmp * tmp
        Next i
        calcul_variance = calcul_variance / (UBound(tab) - 1)
    End Function

End Class
