﻿Imports System.Windows.Forms.DataVisualization.Charting
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

        'tableau des AR moyen normalisés
        Dim tabMoyAR() As Double = RentaAnormales.moyNormAR(tabEstAR, tabEvAR)
        'tableau des écart-types des AR normalisés
        Dim tabEcartAR() As Double = RentaAnormales.ecartNormAR(tabEstAR, tabEvAR, tabMoyAR)

        'A FAIRE : affichage résultats
        MsgBox("ok")
    End Sub

End Module
