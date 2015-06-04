﻿Module TestsStatistiques

    '***************************** T-Test *****************************

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

    'Renvoie true si l'hypothèse H0 est rejetée
    Public Function calculStatStudent(tabCAR() As Double) As Double
        Dim tailleTabCAR As Integer = tabCAR.GetLength(0)
        Dim moyenneTab As Double = calcul_moyenne(tabCAR)
        Dim varianceTab As Double = calcul_variance(tabCAR, moyenneTab)
        calculStatStudent = Math.Abs(Math.Sqrt(tailleTabCAR) * moyenneTab / Math.Sqrt(varianceTab))
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

    Public Function calculPValeur(tailleEchant As Integer, testHyp As Double) As Double
        Return Globals.ThisAddIn.Application.WorksheetFunction.T_Dist_2T(testHyp, tailleEchant - 1)
    End Function

    '***************************** Test de Patell *****************************

    Public Function patellTest(tabAR(,) As Double, fenetreEstDebut As Integer, fenetreEstFin As Integer, fenetreEvDebut As Integer, fenetreEvFin As Integer) As Double
        'La formule utilisée est donnée page 80 de "Eventus-Guide"
        'On se positionne sur la feuille des Rt
        Dim currentSheet As Excel.Worksheet = CType(Globals.ThisAddIn.Application.Worksheets("Rm"), Excel.Worksheet)

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
        Dim currentSheet As Excel.Worksheet = CType(Globals.ThisAddIn.Application.Worksheets("Rm"), Excel.Worksheet)
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
        Dim currentSheet As Excel.Worksheet = CType(Globals.ThisAddIn.Application.Worksheets("Rm"), Excel.Worksheet)
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
        Dim currentSheet As Excel.Worksheet = CType(Globals.ThisAddIn.Application.Worksheets("Rt"), Excel.Worksheet)

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


    '***************************** Test de signe *****************************

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

End Module