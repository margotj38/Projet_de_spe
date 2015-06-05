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

            Select Case Globals.Ribbons.Ruban.selFenetres.test
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

    'Traitement des ARs à partir des deux tableaux d'estimation et de 
    Public Sub traitementTabAR(tabEvAR(,) As Object, tabEstAR(,) As Object)

        'tableau des AR moyen normalisés
        Dim tabMoyAR() As Double = RentaAnormales.moyNormAR(tabEstAR, tabEvAR)
        'tableau des écart-types des AR normalisés
        Dim tabEcartAR() As Double = RentaAnormales.ecartNormAR(tabEstAR, tabEvAR, tabMoyAR)

        Dim tailleFenetreEv As Integer = tabEvAR.GetLength(0)
        Dim N As Integer = tabEvAR.GetLength(1)

        'tableau des CAR
        Dim tabCAR(tailleFenetreEv - 1, N - 1) As Double

        For e = 1 To N
            Dim somme As Double = 0
            For i = 1 To tailleFenetreEv
                somme = somme + tabEvAR(i, e)
                tabCAR(i - 1, e - 1) = somme
            Next
        Next
    End Sub

    '***************************** traitement fichier AR *****************************
    Public Sub traitementAR(plageEst As String, plageEv As String)
        'Sélection de la feuille contenant les Rt
        Dim currentSheet As Excel.Worksheet = CType(Globals.ThisAddIn.Application.Worksheets("AR"), Excel.Worksheet)

        Dim tmpRange As Excel.Range
        tmpRange = currentSheet.Range(plageEst)
        'tableau des données pour l'estimation
        Dim tabEstAR(,) As Object = currentSheet.Range(tmpRange.Cells(1, 2), tmpRange.Cells(tmpRange.Rows.Count, tmpRange.Columns.Count)).Value
        'extraction de la première colonne correspondant aux dates
        tmpRange = currentSheet.Range(plageEv)
        Dim dates As Excel.Range = currentSheet.Range(tmpRange.Cells(1, 1), tmpRange.Cells(tmpRange.Rows.Count, 1))
        'tableau des données pour l'estimation
        Dim tabEvAR(,) As Object = currentSheet.Range(tmpRange.Cells(1, 2), tmpRange.Cells(tmpRange.Rows.Count, tmpRange.Columns.Count)).Value
        'taille fenêtre  d'événement
        Dim tailleFenetreEv As Integer = tabEvAR.GetLength(0)
        Dim N As Integer = tabEvAR.GetLength(1)

        'tableau des AR moyen normalisés
        Dim tabMoyAR() As Double = RentaAnormales.moyNormAR(tabEstAR, tabEvAR)
        'tableau des écart-types des AR normalisés
        Dim tabEcartAR() As Double = RentaAnormales.ecartNormAR(tabEstAR, tabEvAR, tabMoyAR)


        '-----------------------------------A FAIRE : affichage résultats

        'La création d'une nouvelle feuille
        Dim nom As String
        nom = InputBox("Entrer Le nom de la feuille des résultats de l'étude d'événements: ")
        'Si l'utilisateur n'entre pas un nom
        If nom Is "" Then nom = "Resultat"
        Globals.ThisAddIn.Application.Sheets.Add(After:=Globals.ThisAddIn.Application.Worksheets(Globals.ThisAddIn.Application.Worksheets.Count))
        Globals.ThisAddIn.Application.ActiveSheet.Name = nom

        'Le nom de chaque colonne
        nomCellule(Globals.ThisAddIn.Application.Worksheets(nom).Range("B1"), "Moyenne")
        nomCellule(Globals.ThisAddIn.Application.Worksheets(nom).Range("C1"), "Ecart-type")
        nomCellule(Globals.ThisAddIn.Application.Worksheets(nom).Range("D1"), "T-test")
        nomCellule(Globals.ThisAddIn.Application.Worksheets(nom).Range("E1"), "P-valeur (%)")

        Dim j As Integer
        j = 2
        For Each var_Rge In dates
            nomCellule(Globals.ThisAddIn.Application.Worksheets(nom).Range("A" & j), "AR(" & var_Rge.value & ")")
            j = j + 1
        Next var_Rge


        For i = 0 To tailleFenetreEv - 1
            j = i + 2
            'La colonne des moyennes
            valeurCellule(Globals.ThisAddIn.Application.Worksheets(nom).Range("B" & j), tabMoyAR(i))

            'La colonne des écart-types
            valeurCellule(Globals.ThisAddIn.Application.Worksheets(nom).Range("C" & j), tabEcartAR(i))

            'La statistique du test
            Dim stat As Double = Math.Abs(Math.Sqrt(N) * tabMoyAR(i) / tabEcartAR(i))
            valeurCellule(Globals.ThisAddIn.Application.Worksheets(nom).Range("D" & j), stat)

            'La colonne des p-valeurs
            'A Décommenter après
            Dim pValeur As Double = Globals.ThisAddIn.Application.WorksheetFunction.T_Dist_2T(stat, N - 1)
            valeurCellule(Globals.ThisAddIn.Application.Worksheets(nom).Range("E" & j), pValeur * 100)
            'La signification du test
            Globals.ThisAddIn.Application.Worksheets(nom).Range("F" & j).Value = signification(pValeur)
        Next i

        '************************ Tableau de résultats des CAR
        nomCellule(Globals.ThisAddIn.Application.Worksheets(nom).Range("B" & tailleFenetreEv + 4), "Moyenne")
        nomCellule(Globals.ThisAddIn.Application.Worksheets(nom).Range("C" & tailleFenetreEv + 4), "Ecart-type")
        nomCellule(Globals.ThisAddIn.Application.Worksheets(nom).Range("D" & tailleFenetreEv + 4), "T-test")
        nomCellule(Globals.ThisAddIn.Application.Worksheets(nom).Range("E" & tailleFenetreEv + 4), "P-valeur (%)")

        j = tailleFenetreEv + 5
        For Each var_Rge In dates
            nomCellule(Globals.ThisAddIn.Application.Worksheets(nom).Range("A" & j), "CAR(" & var_Rge.value & ")")
            j = j + 1
        Next var_Rge


        'Remplissage des tableaux :moyenne, variance
        Dim tabCAR(,) As Double = RentaAnormales.CalculCar(tabEvAR)
        Dim tabMoyCar() As Double = RentaAnormales.calculMoyenneCar(tabCAR)
        Dim tabVarCar() As Double = RentaAnormales.calculVarianceCar(tabCAR, tabMoyCar)

        For i = 0 To tailleFenetreEv - 1
            j = i + tailleFenetreEv + 5
            'La colonne des moyennes
            valeurCellule(Globals.ThisAddIn.Application.Worksheets(nom).Range("B" & j), tabMoyCar(i))

            'La colonne des écart-types
            valeurCellule(Globals.ThisAddIn.Application.Worksheets(nom).Range("C" & j), Math.Sqrt(tabVarCar(i)))

            'La statistique du test
            Dim stat As Double = Math.Abs(Math.Sqrt(N) * tabMoyCar(i) / Math.Sqrt(tabVarCar(i)))
            valeurCellule(Globals.ThisAddIn.Application.Worksheets(nom).Range("D" & j), stat)

            'La colonne des p-valeurs
            Dim pValeur As Double = Globals.ThisAddIn.Application.WorksheetFunction.T_Dist_2T(stat, N - 1)
            valeurCellule(Globals.ThisAddIn.Application.Worksheets(nom).Range("E" & j), pValeur * 100)
            'La signification du test
            Globals.ThisAddIn.Application.Worksheets(nom).Range("F" & j).Value = signification(pValeur)
        Next i
    End Sub

    'Associe un nom à une cellule avec une mise en forme
    Private Sub nomCellule(r As Excel.Range, Valeur As String)
        r.Value = Valeur
        r.Font.Bold = True
        r.Borders.Value = 1
        r.Interior.ColorIndex = 27
    End Sub

    'Associe une valeur à une cellule avec des bordures sur le tableau
    Private Sub valeurCellule(r As Excel.Range, valeur As Double)
        r.Value = valeur
        r.Borders.Value = 1
    End Sub

    'Renvoie une chaine de caractère qui indique la signification d'un test
    Function signification(seuil As Double) As String
        Dim signifi As String
        Select Case seuil
            Case Is < 0.001
                signifi = "***"
            Case Is < 0.01
                signifi = "**"
            Case Is < 0.05
                signifi = "*"
            Case Is < 0.1
                signifi = "."
            Case Else
                signifi = ""
        End Select

        signification = signifi
    End Function

    Public Sub affichageRentaCentrees(tabrenta(,) As Double)
        'Perte du numéro des entreprises : pb ?

        'Création d'une nouvelle feuille
        Dim nom As String
        nom = InputBox("Entrer le nom de la feuille des rentabilités centrées : ")
        'Si l'utilisateur n'entre pas un nom
        If nom Is "" Then nom = "Rentabilités centrées"
        Globals.ThisAddIn.Application.Sheets.Add(After:=Globals.ThisAddIn.Application.Worksheets(Globals.ThisAddIn.Application.Worksheets.Count))
        Globals.ThisAddIn.Application.ActiveSheet.Name = nom

        'Affichage des dates
        Globals.ThisAddIn.Application.Worksheets(nom).Range("A1").Value = "Dates"
        For i = 0 To tabrenta.GetUpperBound(0)
            Globals.ThisAddIn.Application.Worksheets(nom).Range("A" & i + 2).Value = tabrenta(i, 0)
            Globals.ThisAddIn.Application.Worksheets(nom).Range("A" & i + 2).Borders.Value = 1
        Next i

        'On écrit la première ligne
        For colonne = 1 To tabrenta.GetUpperBound(1)
            Globals.ThisAddIn.Application.Worksheets(nom).Cells(1, colonne + 1).Value = "R" & colonne
        Next colonne

        'Affichage des rentabilités
        For colonne = 1 To tabrenta.GetUpperBound(1)
            For i = 0 To tabrenta.GetUpperBound(0)
                Globals.ThisAddIn.Application.Worksheets(nom).Cells(i + 2, colonne + 1).Value = tabrenta(i, colonne)
            Next i
        Next colonne
    End Sub

    Public Sub affichageAR(ByRef tabAREst(,) As Double, ByRef tabAREv(,) As Double, _
                           ByRef tabDateEst() As Integer, ByRef tabDateEv() As Integer)
        'Création d'une nouvelle feuille
        Dim nom As String
        nom = InputBox("Entrer le nom de la feuille des rentabilités anormales : ")
        'Si l'utilisateur n'entre pas un nom
        If nom Is "" Then nom = "Rentabilités anormales"
        Globals.ThisAddIn.Application.Sheets.Add(After:=Globals.ThisAddIn.Application.Worksheets(Globals.ThisAddIn.Application.Worksheets.Count))
        Globals.ThisAddIn.Application.ActiveSheet.Name = nom

        Dim currentSheet As Excel.Worksheet = CType(Globals.ThisAddIn.Application.Worksheets(nom), Excel.Worksheet)

        'Affichage de la première ligne
        For i = 0 To tabAREst.GetUpperBound(1)
            nomCellule(currentSheet.Cells(1, i + 2), "AR" & i + 1)
        Next i

        'Affichage des dates pour la période d'estimation
        nomCellule(currentSheet.Range("A1"), "Dates")
        For i = 0 To tabDateEst.GetUpperBound(0)
            nomCellule(currentSheet.Range("A" & i + 2), tabDateEst(i).ToString)
        Next i

        'Affichage des données pour la période d'estimation
        For colonne = 0 To tabAREst.GetUpperBound(1)
            For i = 0 To tabAREst.GetUpperBound(0)
                valeurCellule(currentSheet.Cells(i + 2, colonne + 2), tabAREst(i, colonne))
            Next i
        Next colonne

        'Affichage des dates pour la période d'événement
        For i = 0 To tabDateEv.GetUpperBound(0)
            nomCellule(currentSheet.Range("A" & 4 + tabDateEst.GetLength(0) + i), tabDateEv(i).ToString)
        Next i

        'Affichage des données pour la période d'estimation
        For colonne = 0 To tabAREv.GetUpperBound(1)
            For i = 0 To tabAREv.GetUpperBound(0)
                valeurCellule(currentSheet.Cells(4 + tabDateEst.GetLength(0) + i, colonne + 2), tabAREst(i, colonne))
            Next i
        Next colonne

    End Sub

End Module
