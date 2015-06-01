Imports System.Windows.Forms.DataVisualization.Charting
Imports System.Diagnostics

Public Class ThisAddIn

    'Calcule les AR avec le modèle considéré
    Public Function calculAR(fenetreEstDebut As Integer, fenetreEstFin As Integer) As Double(,)
        'appelle une fonction pour chaque modèle
        Select Case Globals.Ribbons.Ruban.choixSeuilFenetre.modele
            Case 0
                calculAR = modeleMoyenne(fenetreEstDebut, fenetreEstFin)
            Case 1
                calculAR = modeleMarcheSimple()
            Case 2
                calculAR = modeleMarche(fenetreEstDebut, fenetreEstFin)
            Case Else
                MsgBox("Erreur interne : numero de modèle incorrect dans ChoixSeuilFenetre", 16)
                calculAR = Nothing
        End Select
    End Function

    'Calcule les CAR "normalisés" pour le test statistique
    Public Function calculCAR(tabAR As Double(,), fenetreEstDebut As Integer, fenetreEstFin As Integer, fenetreEvDebut As Integer, fenetreEvFin As Integer) As Double()
        Dim normCar(tabAR.GetUpperBound(1)) As Double   'Variable aléatoire correspondant aux CAR "normalisés"
        Dim currentSheet As Excel.Worksheet = CType(Globals.ThisAddIn.Application.Worksheets("Rt"), Excel.Worksheet)
        Dim indFenetreEstDeb As Integer = fenetreEstDebut - currentSheet.Cells(2, 1).Value
        Dim indFenetreEstFin As Integer = fenetreEstFin - currentSheet.Cells(2, 1).Value
        Dim indFenetreEvDeb As Integer = fenetreEvDebut - currentSheet.Cells(2, 1).Value
        Dim indFenetreEvFin As Integer = fenetreEvFin - currentSheet.Cells(2, 1).Value
        Dim tailleFenetreEst As Integer = fenetreEstFin - fenetreEstDebut + 1
        Dim tailleFenetreEv As Integer = fenetreEvFin - fenetreEvDebut + 1

        'Calcul de la statistique pour chaque entreprise
        For colonne = 0 To tabAR.GetUpperBound(1)
            'Calcul de CAR sur la fenetre d'événement paramétrée
            Dim CAR As Double = 0
            For i = indFenetreEvDeb To indFenetreEvFin
                CAR = CAR + tabAR(i, colonne)
            Next i
            Dim moyenne As Double = 0
            For i = indFenetreEstDeb To indFenetreEstFin
                moyenne = moyenne + tabAR(i, colonne)
            Next i
            moyenne = moyenne / tailleFenetreEst
            'Calcul de la variance des AR sur la période d'estimation
            Dim variance As Double = 0
            For i = indFenetreEstDeb To indFenetreEstFin
                Dim tmp As Double = tabAR(i, colonne) - moyenne
                variance = variance + tmp * tmp
            Next i
            variance = variance / (tailleFenetreEst - 1)
            normCar(colonne) = CAR / Math.Sqrt(tailleFenetreEv * variance)
            'Debug.WriteLine(normCar(colonne))
        Next colonne
        'retourne le tableau des CAR normalisés
        calculCAR = normCar
    End Function

    'Estimation des AR à partir du modèle de marché : K = alpha + beta*Rm
    Public Function modeleMarche(fenetreEstDebut As Integer, fenetreEstFin As Integer) As Double(,)
        'On se positionne sur la feuille des Rt
        Dim currentSheet As Excel.Worksheet = CType(Application.Worksheets("Rt"), Excel.Worksheet)
        'On compte le nombre de lignes et de colonnes
        Dim nbLignes As Integer = currentSheet.UsedRange.Rows.Count
        Dim nbColonnes As Integer = currentSheet.UsedRange.Columns.Count

        'Indices de la fenetre d'estimation
        Dim indFenetreEstDeb As Integer = 2 + fenetreEstDebut - currentSheet.Cells(2, 1).Value
        Dim indFenetreEstFin As Integer = 2 + fenetreEstFin - currentSheet.Cells(2, 1).Value

        'Tableau stockant les AR calculés grâce à la régression
        Dim tabAR(nbLignes - 2, nbColonnes - 2) As Double

        For i = 0 To nbColonnes - 2
            Dim plageY As Excel.Range
            Dim plageX As Excel.Range
            plageY = Application.Range(currentSheet.Cells(indFenetreEstDeb, i + 2), currentSheet.Cells(indFenetreEstFin, i + 2))
            'On se positionne sur la feuille des Rm pour récupérer plageX
            currentSheet = CType(Application.Worksheets("Rm"), Excel.Worksheet)
            plageX = Application.Range(currentSheet.Cells(indFenetreEstDeb, i + 2), currentSheet.Cells(indFenetreEstFin, i + 2))
            'Calcul des paramètres de la régression linéaire
            Dim beta As Double = Application.WorksheetFunction.Index(Application.WorksheetFunction.LinEst(plageY, plageX), 1)
            Dim alpha As Double = Application.WorksheetFunction.Index(Application.WorksheetFunction.LinEst(plageY, plageX), 2)

            'Remplissage du tableau
            For t = 0 To nbLignes - 2
                Dim k As Double = alpha + beta * currentSheet.Cells(t + 2, i + 2).Value
                currentSheet = CType(Application.Worksheets("Rt"), Excel.Worksheet)
                tabAR(t, i) = currentSheet.Cells(t + 2, i + 2).Value - k
                currentSheet = CType(Application.Worksheets("Rm"), Excel.Worksheet)
            Next
            'On retourne sur la feuille des Rt
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
    Public Function modeleMoyenne(fenetreEstDebut As Integer, fenetreEstFin As Integer) As Double(,)
        Dim currentSheet As Excel.Worksheet = CType(Application.Worksheets("Rt"), Excel.Worksheet)
        'On compte le nombre de lignes et de colonnes
        Dim nbLignes As Integer = currentSheet.UsedRange.Rows.Count
        Dim nbColonnes As Integer = currentSheet.UsedRange.Columns.Count

        'Indices de la fenetre d'estimation
        Dim indFenetreEstDeb As Integer = 2 + fenetreEstDebut - currentSheet.Cells(2, 1).Value
        Dim indFenetreEstFin As Integer = 2 + fenetreEstFin - currentSheet.Cells(2, 1).Value

        'Tableau des moyennes de chaque titre
        Dim tabMoy(nbColonnes - 2) As Double

        'Calcul des moyennes
        For colonne = 2 To nbColonnes
            Dim plage As Excel.Range = Application.Range(currentSheet.Cells(indFenetreEstDeb, colonne), currentSheet.Cells(indFenetreEstFin, colonne))
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

    Public Function patellTest(tabAR(,) As Double, fenetreEstDebut As Integer, fenetreEstFin As Integer, fenetreEvDebut As Integer, fenetreEvFin As Integer) As Double
        'La formule utilisée est donnée page 80 de "Eventus-Guide"
        'On se positionne sur la feuille des Rt
        Dim currentSheet As Excel.Worksheet = CType(Application.Worksheets("Rm"), Excel.Worksheet)

        'Indices de la fenêtre d'estimation
        Dim indFenetreEstDeb As Integer = fenetreEstDebut - currentSheet.Cells(2, 1).Value
        Dim indFenetreEstFin As Integer = fenetreEstFin - currentSheet.Cells(2, 1).Value
        Dim indFenetreEvDeb As Integer = fenetreEvDebut - currentSheet.Cells(2, 1).Value
        Dim indFenetreEvFin As Integer = fenetreEvFin - currentSheet.Cells(2, 1).Value
        Dim M As Integer = indFenetreEstFin - indFenetreEstDeb + 1

        Dim sAtjCarre(tabAR.GetUpperBound(0), tabAR.GetUpperBound(1)) As Double

        'Calcul des (s_Aj)²
        Dim sAjCarre(tabAR.GetUpperBound(1)) As Double
        sAjCarre = patellCalcSAj(tabAR, indFenetreEstDeb, indFenetreEstFin)

        'Attention, modification de la formule : on l'étend à plusieurs Rm => toujours ok ?
        'Calcul des Rm_Est
        Dim rmEst(tabAR.GetUpperBound(1)) As Double
        rmEst = patellCalcRmEst(tabAR.GetLength(1), indFenetreEstDeb, indFenetreEstFin)

        'Calcul somme au dénominateur
        Dim sommeDenom(tabAR.GetUpperBound(1)) As Double
        sommeDenom = patellCalcSommeDenom(rmEst, indFenetreEstDeb, indFenetreEstFin)

        'Calcul de la formule complète
        For i = 0 To sAtjCarre.GetUpperBound(0)
            For j = 0 To sAtjCarre.GetUpperBound(1)
                Dim tmp = currentSheet.Cells(i + 2, j + 2).Value - rmEst(j)
                sAtjCarre(i, j) = sAjCarre(j) * (1 + (1 / M) + (tmp * tmp / sommeDenom(j)))
            Next j
        Next i

        'Tableau des SAR
        Dim SAR(tabAR.GetUpperBound(0), tabAR.GetUpperBound(1)) As Double
        For i = 0 To tabAR.GetUpperBound(0)
            For j = 0 To tabAR.GetUpperBound(1)
                SAR(i, j) = tabAR(i, j) / Math.Sqrt(sAtjCarre(i, j))
            Next j
        Next i

        'Calcul de Z-t1,t2
        Dim testHyp As Double
        testHyp = patellCalcZ(SAR, indFenetreEvDeb, indFenetreEvFin, M)
        Return testHyp
    End Function

    Private Function patellCalcSAj(tabAR(,) As Double, indFenetreEstDeb As Integer, indFenetreEstFin As Integer) As Double()
        Dim sAjCarre(tabAR.GetUpperBound(1)) As Double
        For colonne = 0 To tabAR.GetUpperBound(1)
            sAjCarre(colonne) = 0
            For k = indFenetreEstDeb To indFenetreEstFin
                sAjCarre(colonne) = sAjCarre(colonne) + tabAR(k, colonne) * tabAR(k, colonne)
            Next k
            sAjCarre(colonne) = sAjCarre(colonne) / (indFenetreEstFin - indFenetreEstDeb + 1)
        Next colonne
        Return sAjCarre
    End Function

    Private Function patellCalcRmEst(nbTitres As Integer, indFenetreEstDeb As Integer, indFenetreEstFin As Integer) As Double()
        'On se positionne sur la feuille des Rm
        Dim currentSheet As Excel.Worksheet = CType(Application.Worksheets("Rm"), Excel.Worksheet)
        Dim rmEst(nbTitres - 1) As Double
        For colonne = 0 To nbTitres - 1
            rmEst(colonne) = 0
            For i = indFenetreEstDeb To indFenetreEstFin
                rmEst(colonne) = rmEst(colonne) + currentSheet.Cells(i + 2, colonne + 2).Value
            Next i
            rmEst(colonne) = rmEst(colonne) / (indFenetreEstFin - indFenetreEstDeb + 1)
        Next colonne
        Return rmEst
    End Function

    Private Function patellCalcSommeDenom(rmEst As Double(), indFenetreEstDeb As Integer, indFenetreEstFin As Integer) As Double()
        'on se positionne sur la feuille des Rm
        Dim currentSheet As Excel.Worksheet = CType(Application.Worksheets("Rm"), Excel.Worksheet)
        Dim sommeDenom(rmEst.GetUpperBound(0)) As Double

        For colonne = 0 To rmEst.GetUpperBound(0)
            sommeDenom(colonne) = 0
            For i = indFenetreEstDeb To indFenetreEstFin
                Dim tmp As Double = currentSheet.Cells(i + 2, colonne + 2).Value - rmEst(colonne)
                sommeDenom(colonne) = sommeDenom(colonne) + tmp * tmp
            Next i
        Next colonne
        Return sommeDenom
    End Function

    Private Function patellCalcZ(SAR As Double(,), indFenetreEvDeb As Integer, indFenetreEvFin As Integer, M As Integer) As Double
        Dim currentSheet As Excel.Worksheet = CType(Application.Worksheets("Rt"), Excel.Worksheet)

        Dim q As Double = (indFenetreEvFin - indFenetreEvDeb + 1) * (M - 2) / (M - 4)

        Dim z As Double = 0
        For colonne = 0 To SAR.GetUpperBound(1)
            Dim zj As Double = 0
            For t = indFenetreEvDeb To indFenetreEvFin
                zj = zj + SAR(t, colonne)
            Next t
            z = z + zj / Math.Sqrt(q)
        Next colonne
        z = z / Math.Sqrt(SAR.GetLength(1))
        Return z
    End Function

    Public Function calculPValeur(tailleEchant As Integer, testHyp As Double) As Double
        Return Application.WorksheetFunction.T_Dist_2T(testHyp, tailleEchant - 1)
    End Function

    Public Sub tracerPValeur(tailleEchant As Integer, maxFenetre As Integer)
        'Sélection de la feuille contenant les Rt
        Dim currentSheet As Excel.Worksheet = CType(Application.Worksheets("Rt"), Excel.Worksheet)

        'Tant que la fenêtre contient au moins un élément
        For i = 0 To maxFenetre
            Dim tabAR As Double(,)
            'On apprend sur toutes les données disponibles
            tabAR = calculAR(currentSheet.Cells(2, 1).Value, -i - 1)

            Dim pValeur As Double

            Select Case Globals.Ribbons.Ruban.choixSeuilFenetre.test
                Case 0
                    'test simple
                    Dim tabCAR As Double()
                    tabCAR = Globals.ThisAddIn.calculCAR(tabAR, currentSheet.Cells(2, 1).Value, -i - 1, -i, i)
                    Dim testHyp As Double = Globals.ThisAddIn.calculStatistique(tabCAR)
                    pValeur = Globals.ThisAddIn.calculPValeur(tailleEchant, testHyp) * 100
                Case 1
                    'test de Patell
                    Dim testHyp As Double = Globals.ThisAddIn.patellTest(tabAR, currentSheet.Cells(2, 1).Value, -i - 1, -i, i)
                    pValeur = 2 * (1 - Globals.ThisAddIn.Application.WorksheetFunction.Norm_S_Dist(Math.Abs(testHyp), True)) * 100
                Case 2
                    'test de signe
                    Dim testHyp As Double = Globals.ThisAddIn.statTestSigne(tabAR, currentSheet.Cells(2, 1).Value, -i - 1, -i, i)
                    pValeur = 2 * (1 - Globals.ThisAddIn.Application.WorksheetFunction.Norm_S_Dist(Math.Abs(testHyp), True)) * 100
            End Select

            Dim p As New DataPoint
            p.XValue = i
            p.YValues = {pValeur.ToString("0.00000")}

            Globals.Ribbons.Ruban.graphPVal.GraphiqueChart.Series("Series1").Points.Add(p)
        Next i
    End Sub

    'A revoir !!
    Function statTestSigne(tabAR(,) As Double, fenetreEstDebut As Integer, fenetreEstFin As Integer, fenetreEvDebut As Integer, fenetreEvFin As Integer) As Double
        Dim currentSheet As Excel.Worksheet = CType(Globals.ThisAddIn.Application.Worksheets("Rt"), Excel.Worksheet)
        Dim indFenetreEstDeb As Integer = fenetreEstDebut - currentSheet.Cells(2, 1).Value
        Dim indFenetreEstFin As Integer = fenetreEstFin - currentSheet.Cells(2, 1).Value
        Dim indFenetreEvDeb As Integer = fenetreEvDebut - currentSheet.Cells(2, 1).Value
        Dim indFenetreEvFin As Integer = fenetreEvFin - currentSheet.Cells(2, 1).Value
        Dim tailleFenetreEst As Integer = fenetreEstFin - fenetreEstDebut + 1
        Dim tailleFenetreEv As Integer = fenetreEvFin - fenetreEvDebut + 1

        Dim nbPosAR As Double
        nbPosAR = 0
        'On prend les AR > 0 sur la fenêtre d'événement
        For colonne = 0 To tabAR.GetUpperBound(1)
            For i = indFenetreEvDeb To indFenetreEstFin
                If (tabAR(i, colonne) > 0) Then
                    nbPosAR = nbPosAR + 1
                End If
            Next i
        Next colonne

        'Estimation de p sur la fenêtre d'estimation
        Dim p As Double
        p = 0
        For colonne = 0 To tabAR.GetUpperBound(1)
            For i = indFenetreEstDeb To indFenetreEstFin
                If (tabAR(i, colonne) > 0) Then
                    p = p + 1
                End If
            Next i
            p = p / tailleFenetreEst
        Next colonne
        p = p / tailleFenetreEv

        'Calcul de la statistique du test
        Dim stat As Double
        stat = (nbPosAR - tailleFenetreEv * p) / (Math.Sqrt(tailleFenetreEv * p * (1 - p)))

        statTestSigne = stat
    End Function

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

    Public Sub calculRentabilite()
        'Création de la feuille contenant les rentabilités
        Application.Sheets.Add()
        Application.ActiveSheet.Name = "Rt"
        'On sélectionne la feuille contenant les cours
        Dim currentSheet As Excel.Worksheet = CType(Globals.ThisAddIn.Application.Worksheets("Prix"), Excel.Worksheet)
        Dim nbLignes As Integer = currentSheet.UsedRange.Rows.Count
        Dim nbColonnes As Integer = currentSheet.UsedRange.Columns.Count

        'On commence par recopier les dates
        Dim tmp(nbLignes - 3) As Date
        For ligne = 3 To nbLignes
            tmp(ligne - 3) = currentSheet.Cells(ligne, 1).Value
        Next ligne
        'On se place sur la feuille des rentabilités
        currentSheet = CType(Application.Worksheets("Rt"), Excel.Worksheet)
        For ligne = 2 To nbLignes - 1
            currentSheet.Cells(ligne, 1).Value = tmp(ligne - 2)
        Next ligne

        'On écrit la première ligne de la feuille des renatbilités
        'Premiere case
        currentSheet = CType(Application.Worksheets("Prix"), Excel.Worksheet)
        Dim nom As String = currentSheet.Cells(1, 1).Value
        currentSheet = CType(Application.Worksheets("Rt"), Excel.Worksheet)
        currentSheet.Cells(1, 1).Value = nom
        'Deuxième case
        currentSheet = CType(Application.Worksheets("Prix"), Excel.Worksheet)
        Dim marche As String = currentSheet.Cells(1, 2).Value
        currentSheet = CType(Application.Worksheets("Rt"), Excel.Worksheet)
        currentSheet.Cells(1, 2).Value = marche
        'Cases suivantes
        For colonne = 3 To nbColonnes
            currentSheet.Cells(1, colonne).Value = "R" & colonne - 2
        Next colonne

        currentSheet = CType(Application.Worksheets("Prix"), Excel.Worksheet)

        'On s'occupe des titres des entreprises et de la rentabilité de marché
        'On calcule les rentabilités
        Dim tabRenta(nbLignes - 3, nbColonnes - 2)
        For titre = 2 To nbColonnes
            For indDate = 3 To nbLignes
                'On calcule la rentabilité
                tabRenta(indDate - 3, titre - 2) = (currentSheet.Cells(indDate, titre).Value - currentSheet.Cells(indDate - 1, titre).Value) / currentSheet.Cells(indDate - 1, titre).Value
            Next indDate
        Next titre
        'On affiche les rentabilités
        currentSheet = CType(Application.Worksheets("Rt"), Excel.Worksheet)
        For titre = 2 To nbColonnes
            For indDate = 3 To nbLignes
                currentSheet.Cells(indDate - 1, titre).Value = tabRenta(indDate - 3, titre - 2)
            Next indDate
        Next titre
    End Sub

    Public Sub calculRentabiliteAvecNA()
        'Création de la feuille contenant les rentabilités
        Application.Sheets.Add()
        Application.ActiveSheet.Name = "Rt"
        'Et de celle contenant les rentabilités de marché associées
        Application.Sheets.Add()
        Application.ActiveSheet.Name = "Rm"
        'On sélectionne la feuille contenant les cours
        Dim currentSheet As Excel.Worksheet = CType(Globals.ThisAddIn.Application.Worksheets("prixCentres"), Excel.Worksheet)
        Dim nbLignes As Integer = currentSheet.UsedRange.Rows.Count
        Dim nbColonnes As Integer = currentSheet.UsedRange.Columns.Count

        'On commence par recopier les dates
        Dim tmp(nbLignes - 3) As Integer
        For ligne = 3 To nbLignes
            tmp(ligne - 3) = currentSheet.Cells(ligne, 1).Value
        Next ligne
        'On se place sur la feuille des rentabilités
        currentSheet = CType(Application.Worksheets("Rt"), Excel.Worksheet)
        For ligne = 2 To nbLignes - 1
            currentSheet.Cells(ligne, 1).Value = tmp(ligne - 2)
        Next ligne
        'Puis sur celle des rentabilités de marché
        currentSheet = CType(Application.Worksheets("Rm"), Excel.Worksheet)
        For ligne = 2 To nbLignes - 1
            currentSheet.Cells(ligne, 1).Value = tmp(ligne - 2)
        Next ligne

        'On écrit la première ligne des deux feuilles (Rt et Rm)
        'Premiere case
        currentSheet = CType(Application.Worksheets("prixCentres"), Excel.Worksheet)
        Dim nom As String = currentSheet.Cells(1, 1).Value
        currentSheet = CType(Application.Worksheets("Rt"), Excel.Worksheet)
        currentSheet.Cells(1, 1).Value = nom
        currentSheet = CType(Application.Worksheets("Rm"), Excel.Worksheet)
        currentSheet.Cells(1, 1).Value = nom
        'Cases suivantes (sauf la deuxième qui correspond au marché)
        For colonne = 2 To nbColonnes - 1
            currentSheet.Cells(1, colonne).Value = "R" & colonne - 1
        Next colonne
        'Cases suivantes pour les Rt
        currentSheet = CType(Application.Worksheets("Rt"), Excel.Worksheet)
        For colonne = 2 To nbColonnes - 1
            currentSheet.Cells(1, colonne).Value = "Rm titre " & colonne - 1
        Next colonne

        currentSheet = CType(Application.Worksheets("prixCentres"), Excel.Worksheet)

        'On s'occupe des titres des entreprises
        'Variable permettant de savoir à quelle date il faut remonter (une avant, deux avant, ...)
        Dim prixPresent As Integer = 0
        'Pour savoir combien de tableaux stockant les Rt et Rm on va déclaré
        Dim maxPrixAbsent As Integer = 0
        'On calcule les rentabilités et les rentabilités de marché associées
        Dim tabRenta(nbLignes - 3, nbColonnes - 3)
        Dim tabRentaMarche(nbLignes - 3, nbColonnes - 3)

        For titre = 3 To nbColonnes
            For indDate = 2 To nbLignes
                If prixPresent = 0 Then
                    'Si on est sur le premier prix
                    If Not Application.WorksheetFunction.IsNA(currentSheet.Cells(indDate, titre)) Then
                        prixPresent = prixPresent + 1
                        If prixPresent > maxPrixAbsent Then
                            maxPrixAbsent = prixPresent
                        End If
                    End If
                ElseIf Application.WorksheetFunction.IsNA(currentSheet.Cells(indDate, titre)) Then
                    'Si il n'y a pas de prix à cette date
                    'On met un équivalent de #N/A dans les tableaux
                    tabRenta(indDate - 3, titre - 3) = Nothing
                    tabRentaMarche(indDate - 3, titre - 3) = Nothing
                    prixPresent = prixPresent + 1
                    If prixPresent > maxPrixAbsent Then
                        maxPrixAbsent = prixPresent
                    End If
                Else
                    'Sinon on fait le calcul en remontant au dernier prix disponible
                    tabRenta(indDate - 3, titre - 3) = (currentSheet.Cells(indDate, titre).Value - currentSheet.Cells(indDate - prixPresent, titre).Value) / currentSheet.Cells(indDate - prixPresent, titre).Value
                    'On fait de même pour les rentabilités de marché
                    currentSheet = CType(Application.Worksheets("marcheCentre"), Excel.Worksheet)
                    tabRentaMarche(indDate - 3, titre - 3) = (currentSheet.Cells(indDate, titre).Value - currentSheet.Cells(indDate - prixPresent, titre).Value) / currentSheet.Cells(indDate - prixPresent, titre).Value
                    'Puis on se replace sur la feuille des prix
                    currentSheet = CType(Application.Worksheets("prixCentres"), Excel.Worksheet)
                    'Et on indique qu'un prix était présent
                    prixPresent = 1
                End If
            Next indDate
            prixPresent = 0
        Next titre

        'On crée les tableaux de rentabilité (et rentabilité de marché) pour les périodes d'estimation et d'événement
        Dim rentaEst()() As Double = New Double(maxPrixAbsent - 1)() {}
        Dim rentaEv()() As Double = New Double(maxPrixAbsent - 1)() {}
        For i = 0 To maxPrixAbsent - 1
            rentaEst(i) = New Double() {}
        Next i

        'On affiche les rentabilités
        currentSheet = CType(Application.Worksheets("Rt"), Excel.Worksheet)
        For titre = 2 To nbColonnes - 1
            For indDate = 3 To nbLignes
                If Not IsNothing(tabRenta(indDate - 3, titre - 2)) Then
                    currentSheet.Cells(indDate - 1, titre).Value = tabRenta(indDate - 3, titre - 2)
                End If
            Next indDate
        Next titre
        'Et les rentabilités de marché
        currentSheet = CType(Application.Worksheets("Rm"), Excel.Worksheet)
        For titre = 2 To nbColonnes - 1
            For indDate = 3 To nbLignes
                If Not IsNothing(tabRentaMarche(indDate - 3, titre - 2)) Then
                    currentSheet.Cells(indDate - 1, titre).Value = tabRentaMarche(indDate - 3, titre - 2)
                End If
            Next indDate
        Next titre
    End Sub

End Class
