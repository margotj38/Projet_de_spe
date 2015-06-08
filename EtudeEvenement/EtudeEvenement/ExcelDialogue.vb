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
                    'Dim testHyp As Double = TestsStatistiques.calculStatStudent(tabCAR)
                    'pValeur = TestsStatistiques.calculPValeur(tailleEchant, testHyp) * 100
                Case 1
                    'test de Patell
                    'Dim testHyp As Double = TestsStatistiques.patellTest(tabAR, currentSheet.Cells(2, 1).Value, -i - 1, -i, i)
                    'pValeur = 2 * (1 - Globals.ThisAddIn.Application.WorksheetFunction.Norm_S_Dist(Math.Abs(testHyp), True)) * 100
                Case 2
                    'test de signe
                    'Dim testHyp As Double = TestsStatistiques.statTestSigne(tabAR, currentSheet.Cells(2, 1).Value, -i - 1, -i, i)
                    'pValeur = 2 * (1 - Globals.ThisAddIn.Application.WorksheetFunction.Norm_S_Dist(Math.Abs(testHyp), True)) * 100
            End Select

            Dim p As New DataPoint
            p.XValue = i
            p.YValues = {pValeur.ToString("0.00000")}

            Globals.Ribbons.Ruban.graphPVal.GraphiqueChart.Series("Series1").Points.Add(p)
        Next i
    End Sub

    'Traitement des ARs à partir des deux tableaux des AR
    Public Sub traitementTabAR(tabEvAR(,) As Double, tabEstAR(,) As Double, datesEvAR() As Integer)

        'La création d'une nouvelle feuille
        Dim nom As String
        nom = InputBox("Entrer Le nom de la feuille des résultats de l'étude d'événements: ")
        'Si l'utilisateur n'entre pas un nom
        If nom Is "" Then nom = "Resultat"
        Globals.ThisAddIn.Application.Sheets.Add(After:=Globals.ThisAddIn.Application.Worksheets(Globals.ThisAddIn.Application.Worksheets.Count))
        Globals.ThisAddIn.Application.ActiveSheet.Name = nom

        '----------------- AR -----------------
        'tableau des AR moyen normalisés
        Dim tabMoyAR() As Double = RentaAnormales.moyNormAR(tabEstAR, tabEvAR)
        'tableau des écart-types des AR normalisés
        Dim tabEcartAR() As Double = RentaAnormales.ecartNormAR(tabEstAR, tabEvAR, tabMoyAR)

        Dim N As Integer = tabEvAR.GetLength(1)

        'affichage des résultats des AR
        afficheResAR(tabMoyAR, tabEcartAR, datesEvAR, N, nom)

        '----------------- CAR -----------------
        Dim tailleFenetreEv As Integer = tabEvAR.GetLength(0)
        'Remplissage des tableaux : CAR, moyenne, variance
        Dim tabCAR(,) As Double = RentaAnormales.CalculCar(tabEvAR)
        Dim tabMoyCar() As Double = RentaAnormales.moyNormCar(tabEstAR, tabCAR)
        Dim tabVarCar() As Double = RentaAnormales.ecartNormCar(tabEstAR, tabCAR, tabMoyCar)

        'affichage des résultats des CAR
        afficheResCAR(tabMoyCar, tabVarCar, datesEvAR, N, nom)

    End Sub

    '***************************** traitement fichier AR *****************************
    Public Sub traitementPlageAR(plageEst As String, plageEv As String, feuille As String)
        Dim currentSheet As Excel.Worksheet = CType(Globals.ThisAddIn.Application.Worksheets(feuille), Excel.Worksheet)

        'La création d'une nouvelle feuille
        Dim nom As String
        nom = InputBox("Entrer Le nom de la feuille des résultats de l'étude d'événements: ")
        'Si l'utilisateur n'entre pas un nom
        If nom Is "" Then nom = "Resultat"
        Globals.ThisAddIn.Application.Sheets.Add(After:=Globals.ThisAddIn.Application.Worksheets(Globals.ThisAddIn.Application.Worksheets.Count))
        Globals.ThisAddIn.Application.ActiveSheet.Name = nom

        '----------------- AR -----------------
        Dim tmpRange As Excel.Range
        tmpRange = currentSheet.Range(plageEst)
        'tableau des données pour l'estimation
        Dim tabEstAR(0 To tmpRange.Rows.Count - 1, 0 To tmpRange.Columns.Count - 2) As Double
        For ligne = 0 To tabEstAR.GetUpperBound(0)
            For colonne = 0 To tabEstAR.GetUpperBound(1)
                tabEstAR(ligne, colonne) = tmpRange.Cells(ligne + 1, colonne + 2).Value
            Next
        Next
        'tabEstAR = currentSheet.Range(tmpRange.Cells(1, 2), tmpRange.Cells(tmpRange.Rows.Count, tmpRange.Columns.Count)).Value
        'extraction de la première colonne correspondant aux dates
        tmpRange = currentSheet.Range(plageEv)
        Dim dates(0 To tmpRange.Rows.Count - 1) As Integer
        For ligne = 0 To dates.GetUpperBound(0)
            dates(ligne) = tmpRange.Cells(ligne + 1, 1).Value
        Next
        'dates = currentSheet.Range(tmpRange.Cells(1, 1), tmpRange.Cells(tmpRange.Rows.Count, 1)).Value
        'tableau des données pour l'estimation
        Dim tabEvAR(0 To tmpRange.Rows.Count - 1, 0 To tmpRange.Columns.Count - 2) As Double
        For ligne = 0 To tabEvAR.GetUpperBound(0)
            For colonne = 0 To tabEstAR.GetUpperBound(1)
                tabEvAR(ligne, colonne) = tmpRange.Cells(ligne + 1, colonne + 2).Value
            Next
        Next
        'tabEvAR = currentSheet.Range(tmpRange.Cells(1, 2), tmpRange.Cells(tmpRange.Rows.Count, tmpRange.Columns.Count)).Value

        'nombre d'entreprises de l'échantillon
        Dim N As Integer = tabEvAR.GetLength(1)

        'tableau des AR moyen normalisés
        Dim tabMoyAR() As Double = RentaAnormales.moyNormAR(tabEstAR, tabEvAR)
        'tableau des écart-types des AR normalisés
        Dim tabEcartAR() As Double = RentaAnormales.ecartNormAR(tabEstAR, tabEvAR, tabMoyAR)

        'affichage des résultats
        afficheResAR(tabMoyAR, tabEcartAR, dates, N, nom)

        '----------------- CAR -----------------
        Dim tailleFenetreEv As Integer = tabEvAR.GetLength(0)
        'Remplissage des tableaux :moyenne, variance
        Dim tabCAR(,) As Double = RentaAnormales.CalculCar(tabEvAR)
        Dim tabMoyCar() As Double = RentaAnormales.moyNormCar(tabEstAR, tabCAR)
        Dim tabVarCar() As Double = RentaAnormales.ecartNormCar(tabEstAR, tabCAR, tabMoyCar)

        'affichage des résultats des CAR
        afficheResCAR(tabMoyCar, tabVarCar, dates, N, nom)

    End Sub


    '----------------------------------- affichage résultats AR
    Public Sub afficheResAR(tabMoyAR() As Double, tabEcartAR() As Double, datesEvAR() As Integer, tailleEch As Integer, nomFeuille As String)

        'Le nom de chaque colonne
        nomCellule(Globals.ThisAddIn.Application.Worksheets(nomFeuille).Range("B1"), "Moyenne")
        nomCellule(Globals.ThisAddIn.Application.Worksheets(nomFeuille).Range("C1"), "Ecart-type")
        nomCellule(Globals.ThisAddIn.Application.Worksheets(nomFeuille).Range("D1"), "T-test")
        nomCellule(Globals.ThisAddIn.Application.Worksheets(nomFeuille).Range("E1"), "P-valeur (%)")

        'indice pour l'écriture dans les cellules
        Dim j As Integer
        For i = 0 To datesEvAR.GetUpperBound(0)

            j = i + 2

            'entête des lignes
            nomCellule(Globals.ThisAddIn.Application.Worksheets(nomFeuille).Range("A" & j), "AR(" & datesEvAR(i) & ")")

            'La colonne des moyennes
            valeurCellule(Globals.ThisAddIn.Application.Worksheets(nomFeuille).Range("B" & j), tabMoyAR(i))

            'La colonne des écart-types
            valeurCellule(Globals.ThisAddIn.Application.Worksheets(nomFeuille).Range("C" & j), tabEcartAR(i))

            'La statistique du test
            Dim stat As Double = Math.Abs(Math.Sqrt(tailleEch) * tabMoyAR(i) / tabEcartAR(i))
            valeurCellule(Globals.ThisAddIn.Application.Worksheets(nomFeuille).Range("D" & j), stat)

            'La colonne des p-valeurs
            'A Décommenter après
            Dim pValeur As Double = Globals.ThisAddIn.Application.WorksheetFunction.T_Dist_2T(stat, tailleEch - 1)
            valeurCellule(Globals.ThisAddIn.Application.Worksheets(nomFeuille).Range("E" & j), pValeur * 100)
            'La signification du test
            Globals.ThisAddIn.Application.Worksheets(nomFeuille).Range("F" & j).Value = signification(pValeur)
        Next i
    End Sub

    '----------------------------------- affichage résultats CAR
    Public Sub afficheResCAR(tabMoyCAR() As Double, tabVarCAR() As Double, datesEvAR() As Integer, tailleEch As Integer, nomFeuille As String)

        Dim tailleFenetreEv As Integer = datesEvAR.GetLength(0)

        nomCellule(Globals.ThisAddIn.Application.Worksheets(nomFeuille).Range("B" & tailleFenetreEv + 4), "Moyenne")
        nomCellule(Globals.ThisAddIn.Application.Worksheets(nomFeuille).Range("C" & tailleFenetreEv + 4), "Ecart-type")
        nomCellule(Globals.ThisAddIn.Application.Worksheets(nomFeuille).Range("D" & tailleFenetreEv + 4), "T-test")
        nomCellule(Globals.ThisAddIn.Application.Worksheets(nomFeuille).Range("E" & tailleFenetreEv + 4), "P-valeur (%)")

        'indice pour l'écriture dans les cellules
        Dim j As Integer
        For i = 0 To tailleFenetreEv - 1
            j = i + tailleFenetreEv + 5

            nomCellule(Globals.ThisAddIn.Application.Worksheets(nomFeuille).Range("A" & j), "CAR(" & datesEvAR(i) & ")")

            'La colonne des moyennes
            valeurCellule(Globals.ThisAddIn.Application.Worksheets(nomFeuille).Range("B" & j), tabMoyCAR(i))

            'La colonne des écart-types
            valeurCellule(Globals.ThisAddIn.Application.Worksheets(nomFeuille).Range("C" & j), Math.Sqrt(tabVarCAR(i)))

            'La statistique du test
            Dim stat As Double = Math.Abs(Math.Sqrt(tailleEch) * tabMoyCAR(i) / Math.Sqrt(tabVarCAR(i)))
            valeurCellule(Globals.ThisAddIn.Application.Worksheets(nomFeuille).Range("D" & j), stat)

            'La colonne des p-valeurs
            Dim pValeur As Double = Globals.ThisAddIn.Application.WorksheetFunction.T_Dist_2T(stat, tailleEch - 1)
            valeurCellule(Globals.ThisAddIn.Application.Worksheets(nomFeuille).Range("E" & j), pValeur * 100)
            'La signification du test
            Globals.ThisAddIn.Application.Worksheets(nomFeuille).Range("F" & j).Value = signification(pValeur)
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

        'Affichage en-tête estimation
        currentSheet.Cells(1, 1).Value = "AR sur la période d'estimation"
        currentSheet.Cells(1, 1).Font.Bold = True
        currentSheet.Cells(1, 1).Interior.ColorIndex = 50
        currentSheet.Cells(1, 2).Interior.ColorIndex = 50
        currentSheet.Cells(1, 3).Interior.ColorIndex = 50

        'Affichage de la première ligne
        For i = 0 To tabAREst.GetUpperBound(1)
            nomCellule(currentSheet.Cells(3, i + 2), "AR" & i + 1)
        Next i

        'Affichage des dates pour la période d'estimation
        nomCellule(currentSheet.Cells(3, 1), "Dates")
        For i = 0 To tabDateEst.GetUpperBound(0)
            nomCellule(currentSheet.Cells(i + 4, 1), tabDateEst(i).ToString)
        Next i

        'Affichage des données pour la période d'estimation
        For colonne = 0 To tabAREst.GetUpperBound(1)
            For i = 0 To tabAREst.GetUpperBound(0)
                If tabAREst(i, colonne) = -2146826246 Then
                    currentSheet.Cells(i + 4, colonne + 2).Value = -2146826246
                    currentSheet.Cells(i + 4, colonne + 2).Borders.Value = 1
                Else
                    valeurCellule(currentSheet.Cells(i + 4, colonne + 2), tabAREst(i, colonne))
                End If
            Next i
        Next colonne

        'Affichage en-tête événement
        currentSheet.Cells(tabAREst.GetUpperBound(0) + 7, 1).Value = "AR sur la période d'événement"
        currentSheet.Cells(tabAREst.GetUpperBound(0) + 7, 1).Font.Bold = True
        currentSheet.Cells(tabAREst.GetUpperBound(0) + 7, 1).Interior.ColorIndex = 50
        currentSheet.Cells(tabAREst.GetUpperBound(0) + 7, 2).Interior.ColorIndex = 50
        currentSheet.Cells(tabAREst.GetUpperBound(0) + 7, 3).Interior.ColorIndex = 50

        'Affichage de la première ligne
        For i = 0 To tabAREst.GetUpperBound(1)
            nomCellule(currentSheet.Cells(8 + tabDateEst.GetLength(0), i + 2), "AR" & i + 1)
        Next i

        'Affichage des dates pour la période d'événement
        nomCellule(currentSheet.Cells(8 + tabDateEst.GetLength(0), 1), "Dates")
        For i = 0 To tabDateEv.GetUpperBound(0)
            nomCellule(currentSheet.Cells(9 + tabDateEst.GetLength(0) + i, 1), tabDateEv(i).ToString)
        Next i

        'Affichage des données pour la période d'événement
        For colonne = 0 To tabAREv.GetUpperBound(1)
            For i = 0 To tabAREv.GetUpperBound(0)
                If tabAREv(i, colonne) = -2146826246 Then
                    currentSheet.Cells(9 + tabDateEst.GetLength(0) + i, colonne + 2).Value = -2146826246
                    currentSheet.Cells(9 + tabDateEst.GetLength(0) + i, colonne + 2).Borders.Value = 1
                Else
                    valeurCellule(currentSheet.Cells(9 + tabDateEst.GetLength(0) + i, colonne + 2), tabAREv(i, colonne))
                End If
            Next i
        Next colonne

    End Sub


    Public Sub affichagePatell(ByRef tabDateEv() As Integer, ByRef testHypAAR() As Double, ByRef testHypCAAR() As Double)
        'Création d'une nouvelle feuille
        Dim nom As String
        nom = InputBox("Entrer le nom de la feuille des résultats : ")
        'Si l'utilisateur n'entre pas un nom
        If nom Is "" Then nom = "Résulatats Patell"
        Globals.ThisAddIn.Application.Sheets.Add(After:=Globals.ThisAddIn.Application.Worksheets(Globals.ThisAddIn.Application.Worksheets.Count))
        Globals.ThisAddIn.Application.ActiveSheet.Name = nom

        '*** Test AAR = 0 ***

        'Le nom de chaque colonne
        nomCellule(Globals.ThisAddIn.Application.Worksheets(nom).Range("A1"), "AAR")
        nomCellule(Globals.ThisAddIn.Application.Worksheets(nom).Range("B1"), "Test Patell")
        nomCellule(Globals.ThisAddIn.Application.Worksheets(nom).Range("C1"), "P-valeur (%)")

        'Affichage des dates et des statistiques du test de Patell et de la P-Valeur
        For i = 0 To tabDateEv.GetUpperBound(0)
            nomCellule(Globals.ThisAddIn.Application.Worksheets(nom).cells(i + 2, 1), tabDateEv(i))
            valeurCellule(Globals.ThisAddIn.Application.Worksheets(nom).cells(i + 2, 2), testHypAAR(i))
            Dim pValeur As Double
            pValeur = 2 * (1 - Globals.ThisAddIn.Application.WorksheetFunction.Norm_S_Dist(Math.Abs(testHypAAR(i)), True))
            valeurCellule(Globals.ThisAddIn.Application.Worksheets(nom).cells(i + 2, 3), pValeur * 100)
            'La signification du test
            Globals.ThisAddIn.Application.Worksheets(nom).Cells(i + 2, 4).Value = signification(pValeur)
        Next i

        '*** Test CAAR = 0 ***

        Dim debutAffichage As Integer = tabDateEv.GetLength(0) + 4

        'Le nom de chaque colonne
        nomCellule(Globals.ThisAddIn.Application.Worksheets(nom).Cells(debutAffichage, 1), "CAAR")
        nomCellule(Globals.ThisAddIn.Application.Worksheets(nom).Cells(debutAffichage, 2), "Test Patell")
        nomCellule(Globals.ThisAddIn.Application.Worksheets(nom).Cells(debutAffichage, 3), "P-valeur (%)")

        'Affichage des dates et des statistiques du test de Patell et de la P-Valeur
        For i = 0 To tabDateEv.GetUpperBound(0)
            nomCellule(Globals.ThisAddIn.Application.Worksheets(nom).cells(i + debutAffichage + 1, 1), "[" & tabDateEv(0) & "; " & tabDateEv(i) & "]")
            valeurCellule(Globals.ThisAddIn.Application.Worksheets(nom).cells(i + debutAffichage + 1, 2), testHypCAAR(i))
            Dim pValeur As Double
            pValeur = 2 * (1 - Globals.ThisAddIn.Application.WorksheetFunction.Norm_S_Dist(Math.Abs(testHypCAAR(i)), True))
            valeurCellule(Globals.ThisAddIn.Application.Worksheets(nom).cells(i + debutAffichage + 1, 3), pValeur * 100)
            'La signification du test
            Globals.ThisAddIn.Application.Worksheets(nom).Cells(i + debutAffichage + 1, 4).Value = signification(pValeur)
        Next i

    End Sub

    Public Sub affichageSigne(ByRef tabDateEv() As Integer, ByRef tabEstAR(,) As Double, ByRef tabEvAR(,) As Double)
        Dim tailleFenetreEv As Integer = tabEvAR.GetLength(0)

        'Création d'une nouvelle feuille
        Dim nom As String
        nom = InputBox("Entrer le nom de la feuille des résultats : ")
        'Si l'utilisateur n'entre pas un nom
        If nom Is "" Then nom = "Résultats Signe"
        Globals.ThisAddIn.Application.Sheets.Add(After:=Globals.ThisAddIn.Application.Worksheets(Globals.ThisAddIn.Application.Worksheets.Count))
        Globals.ThisAddIn.Application.ActiveSheet.Name = nom


        'Le nom de chaque colonne
        nomCellule(Globals.ThisAddIn.Application.Worksheets(nom).Range("B1"), "Test signe")
        nomCellule(Globals.ThisAddIn.Application.Worksheets(nom).Range("C1"), "P-valeur (%)")

        'Appel de la fonction qui calcule la statistique du test de signe
        Dim stat() As Double = TestsStatistiques.statTestSigne(tabEstAR, tabEvAR)

        'Affichage des dates et des statistiques du test de Patell et de la P-Valeur
        For i = 0 To tailleFenetreEv - 1
            nomCellule(Globals.ThisAddIn.Application.Worksheets(nom).cells(i + 2, 1), tabDateEv(i))
            valeurCellule(Globals.ThisAddIn.Application.Worksheets(nom).cells(i + 2, 2), stat(i))
            Dim pValeur As Double
            'Calcul de la p-valeur
            pValeur = 2 * (1 - Globals.ThisAddIn.Application.WorksheetFunction.Norm_S_Dist(Math.Abs(stat(i)), True))
            valeurCellule(Globals.ThisAddIn.Application.Worksheets(nom).cells(i + 2, 3), pValeur * 100)
            'La signification du test
            Globals.ThisAddIn.Application.Worksheets(nom).Cells(i + 2, 4).Value = signification(pValeur)
        Next i

    End Sub

End Module
