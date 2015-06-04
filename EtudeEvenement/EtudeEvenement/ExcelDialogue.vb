Imports System.Windows.Forms.DataVisualization.Charting
Imports System.Diagnostics

Module ExcelDialogue

    '***************************** graphique p-Valeur *****************************

    Public Sub tracerPValeur(tailleEchant As Integer, maxFenetre As Integer)
        'Sélection de la feuille contenant les Rt
        Dim currentSheet As Excel.Worksheet = CType(Globals.ThisAddIn.Application.Worksheets("Rt"), Excel.Worksheet)

        'Tant que la fenêtre contient au moins un élément
        For i = 0 To maxFenetre
            Dim tabAR As Double(,)
            'On apprend sur toutes les données disponibles


            ''''Commenté pour test
            'tabAR = calculAR(currentSheet.Cells(2, 1).Value, -i - 1)

            Dim pValeur As Double

            Select Case Globals.Ribbons.Ruban.choixSeuilFenetre.test
                Case 0
                    'test simple
                    Dim tabCAR As Double()

                    ''''Commenté pour test
                    'tabCAR = Globals.ThisAddIn.calculCAR(tabAR, currentSheet.Cells(2, 1).Value, -i - 1, -i, i)
                    Dim testHyp As Double = TestsStatistiques.calculStatStudent(tabCAR)
                    pValeur = TestsStatistiques.calculPValeur(tailleEchant, testHyp) * 100
                Case 1
                    'test de Patell
                    Dim testHyp As Double = TestsStatistiques.patellTest(tabAR, currentSheet.Cells(2, 1).Value, -i - 1, -i, i)
                    pValeur = 2 * (1 - Globals.ThisAddIn.Application.WorksheetFunction.Norm_S_Dist(Math.Abs(testHyp), True)) * 100
                Case 2
                    'test de signe
                    Dim testHyp As Double = TestsStatistiques.statTestSigne(tabAR, currentSheet.Cells(2, 1).Value, -i - 1, -i, i)
                    pValeur = 2 * (1 - Globals.ThisAddIn.Application.WorksheetFunction.Norm_S_Dist(Math.Abs(testHyp), True)) * 100
            End Select

            Dim p As New DataPoint
            p.XValue = i
            p.YValues = {pValeur.ToString("0.00000")}

            Globals.Ribbons.Ruban.graphPVal.GraphiqueChart.Series("Series1").Points.Add(p)
        Next i
    End Sub

    '***************************** traitement fichier AR *****************************

    Public Sub traitementAR(plageEst As String, plageEv As String)
        'Sélection de la feuille contenant les Rt
        Dim currentSheet As Excel.Worksheet = CType(Globals.ThisAddIn.Application.Worksheets("AR"), Excel.Worksheet)
        'remplissage des tableaux
        Dim tabEstAR(,) As Object = currentSheet.Range(plageEst).Value
        Dim tabEvAR(,) As Object = currentSheet.Range(plageEv).Value
        'taille fenêtre  d'événement
        Dim tailleFenetreEv As Integer = tabEvAR.GetLength(0)
        Dim N As Integer = tabEvAR.GetLength(1)

        'tableau des AR moyen normalisés
        Dim tabMoyAR() As Double = RentaAnormales.moyNormAR(tabEstAR, tabEvAR)
        'tableau des écart-types des AR normalisés
        Dim tabEcartAR() As Double = RentaAnormales.ecartNormAR(tabEstAR, tabEvAR, tabMoyAR)


        'A FAIRE : affichage résultats
        Dim nom As String
        nom = InputBox("Entrer Le nom de la feuille des résultats de l'étude d'événements: ")
        'Si l'utilisateur n'entre pas un nom
        If nom Is "" Then nom = "Resultat"
        Globals.ThisAddIn.Application.Sheets.Add(After:=Globals.ThisAddIn.Application.Worksheets(Globals.ThisAddIn.Application.Worksheets.Count))
        Globals.ThisAddIn.Application.ActiveSheet.Name = nom

        nomColonne(Globals.ThisAddIn.Application.Worksheets(nom).Range("B1"), "Moyenne")
        nomColonne(Globals.ThisAddIn.Application.Worksheets(nom).Range("C1"), "Ecart-type")
        nomColonne(Globals.ThisAddIn.Application.Worksheets(nom).Range("D1"), "T-test")
        nomColonne(Globals.ThisAddIn.Application.Worksheets(nom).Range("E1"), "P-valeur (%)")


        'Nombre de les colonnes de la feuille des résultats
        Dim nbColonnes As Integer = tabMoyAR.GetUpperBound(0) + 1

        Dim j As Integer
        j = 2
        For Each var_Rge In currentSheet.Range(plageEv)
            nomColonne(Globals.ThisAddIn.Application.Worksheets(nom).Range("A" & j), "AR(" & var_Rge.value & ")")
            j = j + 1
        Next var_Rge


        For i = 0 To nbColonnes - 1
            j = i + 2
            'La colonne des moyennes
            Globals.ThisAddIn.Application.Worksheets(nom).Range("B" & j).Value = tabMoyAR(i)
            Globals.ThisAddIn.Application.Worksheets(nom).Range("B" & j).Borders.Value = 1

            'La colonne des écart-types
            Globals.ThisAddIn.Application.Worksheets(nom).Range("C" & j).Value = tabEcartAR(i)
            Globals.ThisAddIn.Application.Worksheets(nom).Range("C" & j).Borders.Value = 1

            'La statistique du test
            Dim stat As Double = Math.Abs(Math.Sqrt(N) * tabMoyAR(i) / tabEcartAR(i))
            Globals.ThisAddIn.Application.Worksheets(nom).Range("D" & j).Value = stat
            Globals.ThisAddIn.Application.Worksheets(nom).Range("D" & j).Borders.Value = 1

            'La colonne des p-valeurs
            'A Décommenter après
            'Dim pValeur As Double = Globals.ThisAddIn.Application.WorksheetFunction.T_Dist_2T(stat, N - 1) * 100
            'Globals.ThisAddIn.Application.Worksheets(nom).Range("E" & j).Value = pValeur
            Globals.ThisAddIn.Application.Worksheets(nom).Range("E" & j).Borders.Value = 1
        Next i


    End Sub


    Private Sub nomColonne(r As Excel.Range, Valeur As String)
        r.Value = Valeur
        r.Font.Bold = True
        r.Borders.Value = 1
    End Sub
End Module
