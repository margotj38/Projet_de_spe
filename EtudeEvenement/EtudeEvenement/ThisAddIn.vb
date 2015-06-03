Imports System.Windows.Forms.DataVisualization.Charting
Imports System.Diagnostics

Public Class ThisAddIn

    'Calcule les AR avec le modèle considéré
    Public Function calculAR(tailleComplete As Integer, maxPrixAbsent As Integer, fenetreEstDebut As Integer, _
                             fenetreEstFin As Integer, premiereDate As Integer, Optional tabRenta(,) As Double = Nothing, Optional tabRentaMarche(,) As Double = Nothing) As Double(,)
        'appelle une fonction pour chaque modèle
        Select Case Globals.Ribbons.Ruban.choixSeuilFenetre.modele
            Case 0
                calculAR = modeleMoyenne(tailleComplete, premiereDate, fenetreEstDebut, fenetreEstFin, tabRenta)
            Case 1
                calculAR = modeleMarcheSimple()
            Case 2
                'Création des tableaux pour pouvoir les X et Y de la régression
                Dim tabRentaReg(,,)() = constructionTableauxNA(maxPrixAbsent, fenetreEstDebut, fenetreEstFin, tabRenta, tabRentaMarche)
                calculAR = modeleMarche(tailleComplete, premiereDate, fenetreEstDebut, fenetreEstFin, tabRenta, tabRentaMarche, tabRentaReg)
            Case Else
                MsgBox("Erreur interne : numero de modèle incorrect dans ChoixSeuilFenetre", 16)
                calculAR = Nothing
        End Select
    End Function

    'Calcule les CAR "normalisés" pour le test statistique
    Public Function calculCAR(tabAR As Double(,), debutInd As Integer, fenetreEstDebut As Integer, fenetreEstFin As Integer, _
                              fenetreEvDebut As Integer, fenetreEvFin As Integer) As Double()
        Dim normCar(tabAR.GetUpperBound(1)) As Double   'Variable aléatoire correspondant aux CAR "normalisés"
        Dim indFenetreEstDeb As Integer = fenetreEstDebut - debutInd
        Dim indFenetreEstFin As Integer = fenetreEstFin - debutInd
        Dim indFenetreEvDeb As Integer = fenetreEvDebut - debutInd
        Dim indFenetreEvFin As Integer = fenetreEvFin - debutInd
        Dim tailleFenetreEst As Integer = fenetreEstFin - fenetreEstDebut + 1
        Dim tailleFenetreEv As Integer = fenetreEvFin - fenetreEvDebut + 1

        'Calcul de la statistique pour chaque entreprise
        For colonne = 0 To tabAR.GetUpperBound(1)
            'Calcul de CAR sur la fenetre d'événement paramétrée
            Dim CAR As Double = 0
            For i = indFenetreEvDeb To indFenetreEvFin
                '(-2146826246 est la valeur obtenue lorsqu'un ".Value" est fait sur une cellule #N/A)
                If Not tabAR(i, colonne) = -2146826246 Then
                    CAR = CAR + tabAR(i, colonne)
                End If
            Next i

            Dim moyenne As Double = 0
            For i = indFenetreEstDeb To indFenetreEstFin
                If Not tabAR(i, colonne) = -2146826246 Then
                    moyenne = moyenne + tabAR(i, colonne)
                End If
            Next i
            moyenne = moyenne / tailleFenetreEst

            'Calcul de la variance des AR sur la période d'estimation
            Dim variance As Double = 0
            'Variable pour savoir si des NA précédaient
            Dim precAR As Integer = 1
            For i = indFenetreEstDeb To indFenetreEstFin
                If tabAR(i, colonne) = -2146826246 Then
                    precAR = precAR + 1
                Else
                    'On divise l'AR par precAR pour pouvoir le comparer à la moyenne
                    Dim tmp As Double = tabAR(i, colonne) / precAR - moyenne
                    'On multiplie tmp² par precAR afin de simuler une variance sur le bon nombre de périodes
                    variance = variance + tmp * tmp * precAR
                    precAR = 1
                End If
            Next i
            variance = variance / (tailleFenetreEst - 1)
            normCar(colonne) = CAR / Math.Sqrt(tailleFenetreEv * variance)
        Next colonne
        'retourne le tableau des CAR normalisés
        calculCAR = normCar
    End Function

    'Estimation des AR à partir du modèle de marché : K = alpha + beta*Rm
    Public Function modeleMarche(tailleFenetre As Integer, premiereDate As Integer, fenetreEstDebut As Integer, fenetreEstFin As Integer, ByRef tabRenta(,) As Double, ByRef tabRentaMarche(,) As Double, ByRef tabRentaReg(,,)() As Double) As Double(,)
        'Indices de la fenêtre d'estimation dans le tableau tabRenta
        Dim indFenetreEstDeb As Integer = fenetreEstDebut - premiereDate
        Dim indFenetreEstFin As Integer = fenetreEstFin - premiereDate
        'nombre de différentes régressions
        Dim nbReg = tabRentaReg.GetLength(1)
        'déclaration des tableaux contenant les alpha et beta de la régression
        Dim a(nbReg) As Double
        Dim b(nbReg) As Double
        'moyenne pondérée pour obtenir les véritables alpha et beta
        Dim alpha As Double = 0
        Dim beta As Double = 0
        'tableau des AR
        Dim tabAR(tabRenta.GetUpperBound(0), tabRenta.GetUpperBound(1)) As Double

        'pour chaque entreprise...
        For colonne = 0 To tabRentaReg.GetUpperBound(0)
            'nombre de rentabilités totale (sans NA)
            Dim nbRent As Integer = 0
            'pour chaque tableau
            For reg = 0 To nbReg - 1
                If Not tabRentaReg(colonne, reg, 0).GetLength(0) = 0 Then
                    'extraction des Rt
                    Dim Y() As Double = tabRentaReg(colonne, reg, 0)
                    Dim X() As Double = tabRentaReg(colonne, reg, 1)
                    'Dim Y(rentaEst(reg, colonne).GetUpperBound(1)) As Double
                    'Dim X(rentaEst(reg, colonne).GetUpperBound(1)) As Double
                    'For t = 0 To rentaEst(reg, colonne).GetUpperBound(1)
                    '    Y(t) = rentaEst(reg, colonne)(1, t)
                    '    'extraction des Rm
                    '    X(t) = rentaEst(reg, colonne)(0, t)
                    'Next
                    'calcul des coefficients des différentes régressions
                    a(reg) = Application.WorksheetFunction.Index(Application.WorksheetFunction.LinEst(Y, X), 2) / (reg + 1)
                    b(reg) = Application.WorksheetFunction.Index(Application.WorksheetFunction.LinEst(Y, X), 1) / (reg + 1)
                    'somme pondérée
                    alpha = alpha + a(reg) * tabRentaReg(colonne, reg, 1).GetLength(0)
                    beta = beta + b(reg) * tabRentaReg(colonne, reg, 1).GetLength(0)
                    nbRent = nbRent + tabRentaReg(colonne, reg, 1).GetLength(0)
                End If
            Next
            'moyenne pondérée
            alpha = alpha / nbRent
            beta = beta / nbRent

            'remplissage des AR
            'Variable pour savoir si des AR précédents sont manquants
            Dim prixPresent As Integer = 1
            For i = 0 To tabRenta.GetUpperBound(0)
                If tabRenta(i, colonne) = -2146826246 Then
                    tabAR(i, colonne) = -2146826246
                    prixPresent = prixPresent + 1
                Else
                    tabAR(i, colonne) = (tabRenta(i, colonne) - (alpha + beta * tabRentaMarche(i, colonne))) * prixPresent
                    prixPresent = 1
                End If
            Next i
        Next

        ''On se positionne sur la feuille des Rt
        'Dim currentSheet As Excel.Worksheet = CType(Application.Worksheets("Rt"), Excel.Worksheet)
        ''On compte le nombre de lignes et de colonnes
        'Dim nbLignes As Integer = currentSheet.UsedRange.Rows.Count
        'Dim nbColonnes As Integer = currentSheet.UsedRange.Columns.Count

        ''Indices de la fenetre d'estimation
        'Dim indFenetreEstDeb As Integer = 2 + fenetreEstDebut - currentSheet.Cells(2, 1).Value
        'Dim indFenetreEstFin As Integer = 2 + fenetreEstFin - currentSheet.Cells(2, 1).Value

        ''Tableau stockant les AR calculés grâce à la régression
        'Dim tabAR(nbLignes - 2, nbColonnes - 2) As Double

        'For i = 0 To nbColonnes - 2
        '    Dim plageY As Excel.Range
        '    Dim plageX As Excel.Range
        '    plageY = Application.Range(currentSheet.Cells(indFenetreEstDeb, i + 2), currentSheet.Cells(indFenetreEstFin, i + 2))
        '    'On se positionne sur la feuille des Rm pour récupérer plageX
        '    currentSheet = CType(Application.Worksheets("Rm"), Excel.Worksheet)
        '    plageX = Application.Range(currentSheet.Cells(indFenetreEstDeb, i + 2), currentSheet.Cells(indFenetreEstFin, i + 2))
        '    'Calcul des paramètres de la régression linéaire
        '    Dim beta As Double = Application.WorksheetFunction.Index(Application.WorksheetFunction.LinEst(plageY, plageX), 1)
        '    Dim alpha As Double = Application.WorksheetFunction.Index(Application.WorksheetFunction.LinEst(plageY, plageX), 2)

        '    'Remplissage du tableau
        '    For t = 0 To nbLignes - 2
        '        Dim k As Double = alpha + beta * currentSheet.Cells(t + 2, i + 2).Value
        '        currentSheet = CType(Application.Worksheets("Rt"), Excel.Worksheet)
        '        tabAR(t, i) = currentSheet.Cells(t + 2, i + 2).Value - k
        '        currentSheet = CType(Application.Worksheets("Rm"), Excel.Worksheet)
        '    Next
        '    'On retourne sur la feuille des Rt
        '    currentSheet = CType(Application.Worksheets("Rt"), Excel.Worksheet)
        'Next
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
    Public Function modeleMoyenne(tailleFenetre As Integer, premiereDate As Integer, fenetreEstDebut As Integer, fenetreEstFin As Integer, ByRef tabRenta(,) As Double) As Double(,)
        'Indices de la fenêtre d'estimation dans le tableau tabRenta
        Dim indFenetreEstDeb As Integer = fenetreEstDebut - premiereDate
        Dim indFenetreEstFin As Integer = fenetreEstFin - premiereDate

        'Tableau des moyennes de chaque titre
        Dim tabMoy(tabRenta.GetUpperBound(1)) As Double

        'Calcul des moyennes sur la fenêtre d'estimation
        For colonne = 0 To tabRenta.GetUpperBound(1)
            For i = indFenetreEstDeb To indFenetreEstFin
                'S'il n'y avait pas de NA, on somme
                If Not tabRenta(i, colonne) = -2146826246 Then
                    tabMoy(colonne) = tabMoy(colonne) + tabRenta(i, colonne)
                End If
            Next i
            tabMoy(colonne) = tabMoy(colonne) / (indFenetreEstFin - indFenetreEstDeb + 1)
        Next colonne

        'Calcul des AR sur la fenêtre
        'Variable pour savoir si des AR précédents sont manquants
        Dim prixPresent As Integer = 1
        Dim tabAR(tabRenta.GetUpperBound(0), tabRenta.GetUpperBound(1)) As Double
        For colonne = 0 To tabRenta.GetUpperBound(1)
            For i = 0 To tabRenta.GetUpperBound(0)
                If tabRenta(i, colonne) = -2146826246 Then
                    tabAR(i, colonne) = -2146826246
                    prixPresent = prixPresent + 1
                Else
                    'On divise la rentabilité par prixPresent pour se ramenner à un équivalent sur une période
                    'Puis on multiplie par cette même valeur pour avoir un AR correspondant au bon nombre de périodes
                    tabAR(i, colonne) = (tabRenta(i, colonne) / prixPresent - tabMoy(colonne)) * prixPresent
                    prixPresent = 1
                End If
            Next i
        Next colonne
        Return tabAR
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


            ''''Commenté pour test
            'tabAR = calculAR(currentSheet.Cells(2, 1).Value, -i - 1)

            Dim pValeur As Double

            Select Case Globals.Ribbons.Ruban.choixSeuilFenetre.test
                Case 0
                    'test simple
                    Dim tabCAR As Double()

                    ''''Commenté pour test
                    'tabCAR = Globals.ThisAddIn.calculCAR(tabAR, currentSheet.Cells(2, 1).Value, -i - 1, -i, i)
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
        Dim N = currentSheet.UsedRange.Columns.Count - 1

        Dim nbPosAR As Double = 0
        'On prend les AR > 0 sur la fenêtre d'événement
        For colonne = 0 To tabAR.GetUpperBound(1)
            For i = indFenetreEvDeb To indFenetreEvFin
                If (tabAR(i, colonne) > 0) Then
                    nbPosAR = nbPosAR + 1
                End If
            Next i
        Next colonne
        'MsgBox(nbPosAR)

        'Estimation de p sur la fenêtre d'estimation
        Dim p As Double = 0
        Dim nb As Double = 0
        For colonne = 0 To tabAR.GetUpperBound(1)
            For i = indFenetreEstDeb To indFenetreEstFin
                If (tabAR(i, colonne) > 0) Then
                    nb = nb + 1
                End If
            Next i
            p = p + nb / tailleFenetreEst
        Next colonne
        MsgBox(p)
        p = p / tailleFenetreEv
        'MsgBox(p)

        'Calcul de la statistique du test
        Dim stat As Double
        stat = (nbPosAR - tailleFenetreEv * p) / (Math.Sqrt(tailleFenetreEv * p * (1 - p)))

        'MsgBox(stat)
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

    'génère le tableau des prix centrés autour de la date d'évènement
    '2 nouveaux onglets sont créés, un pour les prix, un pour le marché
    Sub prixCentres()
        'Création des deux nouvelles feuilles
        Application.Sheets.Add()
        Application.ActiveSheet.Name = "prixCentres"
        Application.Sheets.Add()
        Application.ActiveSheet.Name = "marcheCentre"

        'on se positionne sur la feuille des evenements
        Dim currentSheet As Excel.Worksheet = CType(Globals.ThisAddIn.Application.Worksheets("DateEvt"), Excel.Worksheet)
        'tableau des dates d'évènements
        Dim datesEv(,)
        datesEv = currentSheet.Range("A2:B101").Value
        'Tri du tableau selon les dates
        Tri(datesEv, 2, LBound(datesEv, 1), UBound(datesEv, 1))

        'on se positionne sur la feuille des prix
        currentSheet = CType(Globals.ThisAddIn.Application.Worksheets("Prix"), Excel.Worksheet)
        Dim nbLignes As Integer = currentSheet.UsedRange.Rows.Count
        Dim nbColonnes As Integer = currentSheet.UsedRange.Columns.Count

        'calul taille fenetre globale
        Dim minUp As Integer, minDown As Integer
        'indice premiere date evenement - indice premiere date
        minUp = currentSheet.Range("A:A").Find(Format(datesEv(1, 2), "Short date").ToString).Row - 2
        'indice derniere date - derniere date evenement
        minDown = nbLignes - currentSheet.Columns("A:A").Find(Format(datesEv(UBound(datesEv, 1), 2), "Short date").ToString).Row

        'écritures des entêtes de lignes et colonnes sur la nouvelle feuille prixCentres
        currentSheet = CType(Globals.ThisAddIn.Application.Worksheets("prixCentres"), Excel.Worksheet)
        currentSheet.Cells(1, 1).Value = "Date"
        For i = 2 To nbColonnes - 1
            currentSheet.Cells(1, i).Value = "P" & i - 1
        Next
        For i = -minUp To minDown
            currentSheet.Cells(i + minUp + 2, 1).Value = i
        Next
        'de même pour marcheCentre
        currentSheet = CType(Globals.ThisAddIn.Application.Worksheets("marcheCentre"), Excel.Worksheet)
        currentSheet.Cells(1, 1).Value = "Date"
        For i = 2 To nbColonnes - 1
            currentSheet.Cells(1, i).Value = "Pm pour P" & i - 1
        Next
        For i = -minUp To minDown
            currentSheet.Cells(i + minUp + 2, 1).Value = i
        Next

        For i = 1 To nbColonnes - 2
            'on se positionne sur la feuille contenant les prix
            currentSheet = CType(Globals.ThisAddIn.Application.Worksheets("Prix"), Excel.Worksheet)
            Dim fenetreInf As Integer, fenetreSup As Integer
            Dim dateCour As Excel.Range, firmeCour As Excel.Range
            Dim data As Excel.Range, marche As Excel.Range
            dateCour = currentSheet.Columns("A:A").Find(Format(datesEv(i, 2), "Short date").ToString)
            fenetreInf = dateCour.Row - minUp
            fenetreSup = dateCour.Row + minDown
            firmeCour = currentSheet.Rows("1:1").Find(datesEv(i, 1).ToString)
            'récupération des prix centrés autour de l'évènement
            data = currentSheet.Range(currentSheet.Cells(fenetreInf, firmeCour.Column), currentSheet.Cells(fenetreSup, firmeCour.Column))
            'récupération des indices de marché correspondants
            marche = currentSheet.Range(currentSheet.Cells(fenetreInf, 2), currentSheet.Cells(fenetreSup, 2))
            'on se positionne sur la feuille contenant les prix centrés pour écrire dedans
            currentSheet = CType(Globals.ThisAddIn.Application.Worksheets("prixCentres"), Excel.Worksheet)
            currentSheet.Range(currentSheet.Cells(2, i + 1), currentSheet.Cells(minUp + minDown + 2, i + 1)).Value = data.Value
            'on se positionne sur la feuille contenant les indices de marché pour écrire dedans
            currentSheet = CType(Globals.ThisAddIn.Application.Worksheets("marcheCentre"), Excel.Worksheet)
            currentSheet.Range(currentSheet.Cells(2, i + 1), currentSheet.Cells(minUp + minDown + 2, i + 1)).Value = marche.Value
        Next
    End Sub

    Sub Tri(a(,) As Object, ColTri As Integer, gauche As Integer, droite As Integer) ' Quick sort
        Dim ref As Date = a((gauche + droite) \ 2, ColTri)
        Dim g As Integer = gauche
        Dim d As Integer = droite
        Do
            Do While a(g, ColTri) < ref : g = g + 1 : Loop
            Do While ref < a(d, ColTri) : d = d - 1 : Loop
            If g <= d Then
                Dim tempDate As Date = a(g, 2)
                a(g, 2) = a(d, 2)
                a(d, 2) = tempDate
                Dim temp As String = a(g, 1)
                a(g, 1) = a(d, 1)
                a(d, 1) = temp
                g = g + 1
                d = d - 1
            End If
        Loop While g <= d
        If g < droite Then Tri(a, ColTri, g, droite)
        If gauche < d Then Tri(a, ColTri, gauche, d)
    End Sub

    Public Function calculARAvecNA(fenetreEstDebut As Integer, fenetreEstFin As Integer, _
                                       fenetreEvDebut As Integer, fenetreEvFin As Integer) As Double(,)
        'On sélectionne la feuille contenant les cours
        Dim currentSheet As Excel.Worksheet = CType(Globals.ThisAddIn.Application.Worksheets("prixCentres"), Excel.Worksheet)
        Dim nbLignes As Integer = currentSheet.UsedRange.Rows.Count
        Dim nbColonnes As Integer = currentSheet.UsedRange.Columns.Count

        'Premier passage : on range les rentabilités dans deux tableaux (tabRenta et tabRentaMarche)
        Dim tabRenta(nbLignes - 3, nbColonnes - 2) As Double
        Dim tabRentaMarche(nbLignes - 3, nbColonnes - 2) As Double
        Dim maxPrixAbsent As Integer
        constructionTableauRenta(nbLignes, nbColonnes, maxPrixAbsent, tabRenta, tabRentaMarche)

        'Deuxième passage : on remplit les tableaux nécessaires à la régression linéaire
        'On crée les tableaux de rentabilité (et rentabilité de marché) pour les périodes d'estimation et d'événement
        Dim rentaEst(,)(,) As Double = New Double(maxPrixAbsent - 1, nbColonnes - 2)(,) {}
        Dim rentaEv(,)(,) As Double = New Double(maxPrixAbsent - 1, nbColonnes - 2)(,) {}
        For i = 0 To maxPrixAbsent - 1
            For j = 0 To nbColonnes - 2
                'Tableau à 2 lignes (Rt et Rm) et à 50 colonnes (à redimensionner)
                rentaEst(i, j) = New Double(1, 49) {}
                rentaEv(i, j) = New Double(1, 49) {}
            Next j
        Next i
        'constructionTableauxNA(nbLignes, nbColonnes, maxPrixAbsent, fenetreEstDebut, fenetreEstFin, _
        'fenetreEvDebut, fenetreEvFin, tabRenta, tabRentaMarche, rentaEst, rentaEv)

        'On calcule maintenant les AR
        Dim tailleComplete As Integer = fenetreEstFin - fenetreEstDebut + 1 + fenetreEvFin - fenetreEvDebut + 1
        calculARAvecNA = calculAR(tailleComplete, maxPrixAbsent, fenetreEstDebut, fenetreEstFin, currentSheet.Cells(2, 1).Value + 1, tabRenta, tabRentaMarche)
    End Function

    Private Function constructionTableauxNA(maxPrixAbsent As Integer, fenetreEstDebut As Integer, fenetreEstFin As Integer, _
                                       ByRef tabRenta(,) As Double, ByRef tabRentaMarche(,) As Double) As Double(,,)()
        'Déclaration du tableau à retourner
        Dim tabRentaReg(tabRenta.GetUpperBound(1), maxPrixAbsent - 1, 1)() As Double
        For i = 0 To tabRenta.GetUpperBound(1)
            For j = 0 To maxPrixAbsent - 1
                For k = 0 To 1
                    tabRentaReg(i, j, k) = New Double(fenetreEstFin - fenetreEstDebut + 1) {}
                Next
            Next
        Next

        Dim currentSheet As Excel.Worksheet = CType(Application.Worksheets("prixCentres"), Excel.Worksheet)
        Dim nbLignes As Integer = currentSheet.UsedRange.Rows.Count
        Dim nbColonnes As Integer = currentSheet.UsedRange.Columns.Count
        'On récupère les indices correspondants aux différentes dates
        Dim indFenetreEstDeb As Integer = fenetreEstDebut - currentSheet.Cells(2, 1).Value
        Dim indFenetreEstFin As Integer = fenetreEstFin - currentSheet.Cells(2, 1).Value

        Dim prixPresent = 1
        For titre = 0 To nbColonnes - 2
            'Tableau permettant de savoir si un redimensionnement est nécessaire
            Dim tabRedimEst(maxPrixAbsent - 1) As Integer
            For indDate = indFenetreEstDeb To indFenetreEstFin
                If tabRenta(indDate, titre) = -2146826246 Then
                    'Si il n'y a pas de prix à cette date
                    prixPresent = prixPresent + 1
                Else
                    'Sinon, on range les rentabilités dans le tableau

                    'On ajoute Rt et Rm au tableau
                    'Les rentabilités sont ramenées en équivalent à une période (par division par prixPresent)
                    tabRentaReg(titre, prixPresent - 1, 0)(tabRedimEst(prixPresent - 1)) = tabRenta(indDate, titre)
                    tabRentaReg(titre, prixPresent - 1, 1)(tabRedimEst(prixPresent - 1)) = tabRentaMarche(indDate, titre)

                    'On indique qu'on a ajouté un nouvel élément
                    tabRedimEst(prixPresent - 1) = tabRedimEst(prixPresent - 1) + 1
                    'Et on indique qu'un prix était présent
                    prixPresent = 1
                End If
            Next indDate
            'A la fin, on redimensionne les tableaux pour qu'ils ne contiennent que des valeurs utiles
            For prixPres = 0 To maxPrixAbsent - 1
                'Si la taille du tableau et le nombre de valeurs qu'il contient sont différents
                If Not tabRentaReg(titre, prixPres, 0).GetLength(0) = tabRedimEst(prixPres) Then
                    'On redimensionne pour ne garder que ce qui est utile
                    ReDim Preserve tabRentaReg(titre, prixPres, 0)(tabRedimEst(prixPres) - 1)
                    ReDim Preserve tabRentaReg(titre, prixPres, 1)(tabRedimEst(prixPres) - 1)
                End If
            Next prixPres
            prixPresent = 1
        Next titre
        Return tabRentaReg
    End Function

    Private Sub constructionTableauRenta(nbLignes As Integer, nbColonnes As Integer, ByRef maxPrixAbsent As Integer, _
                                         ByRef tabRenta(,) As Double, ByRef tabRentaMarche(,) As Double)
        Dim currentSheet As Excel.Worksheet = CType(Application.Worksheets("prixCentres"), Excel.Worksheet)
        'Variable permettant de savoir à quelle date il faut remonter (une avant, deux avant, ...)
        Dim prixPresent As Integer = 0
        'Pour savoir combien de tableaux stockant les Rt et Rm on va déclaré
        maxPrixAbsent = 0

        'On calcule les rentabilités et les rentabilités de marché associées
        For titre = 2 To nbColonnes
            For indDate = 2 To nbLignes
                If prixPresent = 0 Then
                    'Si on est sur le premier prix
                    '(-2146826246 est la valeur obtenue lorsqu'un ".Value" est fait sur une cellule #N/A)
                    If Not (Application.WorksheetFunction.IsNA(currentSheet.Cells(indDate, titre)) Or _
                            currentSheet.Cells(indDate, titre).Value = -2146826246) Then
                        prixPresent = prixPresent + 1
                        If prixPresent > maxPrixAbsent Then
                            maxPrixAbsent = prixPresent
                        End If
                    End If
                ElseIf Application.WorksheetFunction.IsNA(currentSheet.Cells(indDate, titre)) Or _
                            currentSheet.Cells(indDate, titre).Value = -2146826246 Then
                    'Si il n'y a pas de prix à cette date
                    'On met un équivalent de #N/A dans les tableaux
                    tabRenta(indDate - 3, titre - 2) = -2146826246
                    tabRentaMarche(indDate - 3, titre - 2) = -2146826246
                    prixPresent = prixPresent + 1
                    If prixPresent > maxPrixAbsent Then
                        maxPrixAbsent = prixPresent
                    End If
                Else
                    'Sinon on fait le calcul en remontant au dernier prix disponible
                    tabRenta(indDate - 3, titre - 2) = (currentSheet.Cells(indDate, titre).Value - currentSheet.Cells(indDate - prixPresent, titre).Value) / currentSheet.Cells(indDate - prixPresent, titre).Value
                    'On fait de même pour les rentabilités de marché
                    currentSheet = CType(Application.Worksheets("marcheCentre"), Excel.Worksheet)
                    tabRentaMarche(indDate - 3, titre - 2) = (currentSheet.Cells(indDate, titre).Value - currentSheet.Cells(indDate - prixPresent, titre).Value) / currentSheet.Cells(indDate - prixPresent, titre).Value
                    'Puis on se replace sur la feuille des prix
                    currentSheet = CType(Application.Worksheets("prixCentres"), Excel.Worksheet)
                    'Et on indique qu'un prix était présent
                    prixPresent = 1
                End If
            Next indDate
            prixPresent = 0
        Next titre
    End Sub

End Class
