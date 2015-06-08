''' <summary>
''' MODULE STATS
''' </summary>
''' <remarks></remarks>

Module TestsStatistiques

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

    Function calcul_moyenne(tab() As Double) As Double
        'tab() peut contenir des #N/A
        calcul_moyenne = 0

        'Variable pour savoir si un #N/A précédait
        Dim prixPresent As Integer = 1
        For i = 0 To tab.GetUpperBound(0)
            If tab(i) = -2146826246 Then
                prixPresent = prixPresent + 1
            Else
                'Sinon on somme en multipliant par le nombre de #N/A présents + 1 (ie prixPresent)
                calcul_moyenne = calcul_moyenne + tab(i) * prixPresent
                prixPresent = 1
            End If
        Next i
        'On divise par la taille de la fenêtre d'estimation moins le nombre de #N/A finaux (ie prixPresent - 1)
        calcul_moyenne = calcul_moyenne / (tab.GetLength(0) - (prixPresent - 1))
    End Function

    Function calcul_variance(tab() As Double, moyenne As Double) As Double
        'tab() peut contenir des #N/A
        calcul_variance = 0

        'Variable pour savoir si un #N/A précédait
        Dim prixPresent As Integer = 1
        For i = 0 To tab.GetUpperBound(0)
            If tab(i) = -2146826246 Then
                prixPresent = prixPresent + 1
            Else
                Dim tmp As Double = tab(i) - moyenne
                'On somme en multipliant par le nombre de #N/A présents + 1 (ie prixPresent)
                calcul_variance = calcul_variance + tmp * tmp * prixPresent
                prixPresent = 1
            End If
        Next i
        'On divise par la taille de la fenêtre d'estimation - 1, moins le nombre de #N/A finaux (ie prixPresent - 1)
        calcul_variance = calcul_variance / (tab.GetLength(0) - prixPresent)
    End Function

    Public Function calculPValeur(tailleEchant As Integer, testHyp As Double) As Double
        Return Globals.ThisAddIn.Application.WorksheetFunction.T_Dist_2T(testHyp, tailleEchant - 1)
    End Function


    '***************************** Test de Patell *****************************

    Public Function calculNbNonMissingReturn(ByRef tabAREst(,) As Double) As Integer()
        Dim nbNMR(tabAREst.GetUpperBound(1)) As Integer
        For colonne = 0 To tabAREst.GetUpperBound(1)
            Dim nbMR As Integer = 0
            For i = 0 To tabAREst.GetUpperBound(0)
                'On compte le nombre d'AR manquants
                If tabAREst(i, colonne) = -2146826246 Then
                    nbMR = nbMR + 1
                End If
            Next i
            nbNMR(colonne) = tabAREst.GetLength(0) - nbMR
        Next colonne
        Return nbNMR
    End Function

    ''' <summary>
    ''' PATELL
    ''' </summary>
    ''' <param name="tabAREst"></param>
    ''' <param name="tabAREv"></param>
    ''' <param name="tabDateEst"></param>
    ''' <param name="tabDateEv"></param>
    ''' <param name="tabRentaClassiquesMarcheEst"></param>
    ''' <param name="tabRentaClassiquesMarcheEv"></param>
    ''' <param name="Mj"></param>
    ''' <param name="testHypAAR"></param>
    ''' <param name="testHypCAAR"></param>
    ''' <remarks></remarks>
    Public Sub patellTest(ByRef tabAREst(,) As Double, ByRef tabAREv(,) As Double, _
                               ByRef tabDateEst() As Integer, ByRef tabDateEv() As Integer, _
                               ByRef tabRentaClassiquesMarcheEst(,) As Double, ByRef tabRentaClassiquesMarcheEv(,) As Double, _
                               ByRef Mj() As Integer, ByRef testHypAAR() As Double, ByRef testHypCAAR() As Double)
        'La formule utilisée est donnée page 80 de "Eventus-Guide"

        '(s_Atj)² uniquement pour la période d'événement
        Dim sAtjCarre(tabAREv.GetUpperBound(0), tabAREv.GetUpperBound(1)) As Double

        'Calcul des (s_Aj)²
        Dim sAjCarre() As Double = patellCalcSAj(tabAREst, Mj)

        'Calcul des Rm_Est
        Dim rmEst() As Double = patellCalcRmEst(tabRentaClassiquesMarcheEst)

        'Calcul somme au dénominateur
        Dim sommeDenom() As Double = patellCalcSommeDenom(tabRentaClassiquesMarcheEst, rmEst)

        'Calcul de la formule complète
        For i = 0 To sAtjCarre.GetUpperBound(0)
            For j = 0 To sAtjCarre.GetUpperBound(1)
                Dim tmp = tabRentaClassiquesMarcheEv(i, j) - rmEst(j)
                sAtjCarre(i, j) = sAjCarre(j) * (1 + (1 / Mj(j)) + (tmp * tmp / sommeDenom(j)))
            Next j
        Next i

        'Tableau des SARuniquement pour la période d'événement
        Dim SAR(tabAREv.GetUpperBound(0), tabAREv.GetUpperBound(1)) As Double
        For i = 0 To tabAREv.GetUpperBound(0)
            For j = 0 To tabAREv.GetUpperBound(1)
                If tabAREv(i, j) = -2146826246 Then
                    SAR(i, j) = -2146826246
                Else
                    SAR(i, j) = tabAREv(i, j) / Math.Sqrt(sAtjCarre(i, j))
                End If
            Next j
        Next i

        'Calcul des statistiques de l'hypothèse AAR = 0
        testHypAAR = patellCalcStatAAR(SAR, Mj)

        'Calcul de Z-t1,t2
        testHypCAAR = patellCalcZ(SAR, Mj)
    End Sub

    Private Function patellCalcSAj(ByRef tabAREst(,) As Double, ByRef Mj() As Integer) As Double()
        Dim sAjCarre(tabAREst.GetUpperBound(1)) As Double
        For colonne = 0 To tabAREst.GetUpperBound(1)
            sAjCarre(colonne) = 0
            For i = 0 To tabAREst.GetUpperBound(0)
                If Not tabAREst(i, colonne) = -2146826246 Then
                    sAjCarre(colonne) = sAjCarre(colonne) + tabAREst(i, colonne) * tabAREst(i, colonne)
                End If
            Next i
            sAjCarre(colonne) = sAjCarre(colonne) / (Mj(colonne) - 2)
        Next colonne
        Return sAjCarre
    End Function

    Private Function patellCalcRmEst(ByRef tabRentaClassiquesMarcheEst(,) As Double) As Double()
        Dim rmEst(tabRentaClassiquesMarcheEst.GetUpperBound(1) - 1) As Double
        'La première colonne contient les dates, on ne l'utilise donc pas
        For colonne = 1 To tabRentaClassiquesMarcheEst.GetUpperBound(1) - 1
            rmEst(colonne - 1) = 0
            For i = 0 To tabRentaClassiquesMarcheEst.GetUpperBound(0)
                rmEst(colonne - 1) = rmEst(colonne - 1) + tabRentaClassiquesMarcheEst(i, colonne)
            Next i
            rmEst(colonne - 1) = rmEst(colonne - 1) / (tabRentaClassiquesMarcheEst.GetLength(0))
        Next colonne
        Return rmEst
    End Function

    Private Function patellCalcSommeDenom(ByRef tabRentaClassiquesMarcheEst(,) As Double, ByRef rmEst As Double()) As Double()
        Dim sommeDenom(tabRentaClassiquesMarcheEst.GetUpperBound(1) - 1) As Double

        For colonne = 1 To tabRentaClassiquesMarcheEst.GetUpperBound(1) - 1
            sommeDenom(colonne - 1) = 0
            For i = 0 To tabRentaClassiquesMarcheEst.GetUpperBound(0)
                Dim tmp As Double = tabRentaClassiquesMarcheEst(i, colonne) - rmEst(colonne - 1)
                sommeDenom(colonne - 1) = sommeDenom(colonne - 1) + tmp * tmp
            Next i
        Next colonne
        Return sommeDenom
    End Function

    Private Function patellCalcZ(SAR As Double(,), Mj() As Integer) As Double()
        Dim z(SAR.GetUpperBound(0)) As Double
        For datesCum = 0 To SAR.GetUpperBound(0)
            For colonne = 0 To SAR.GetUpperBound(1)
                Dim zj As Double = 0
                Dim nbVal = datesCum + 1
                For t = 0 To datesCum
                    If SAR(t, colonne) = -2146826246 Then
                        nbVal = nbVal - 1
                    Else
                        zj = zj + SAR(t, colonne)
                    End If
                Next t
                Dim q As Double = nbVal * (Mj(colonne) - 2) / (Mj(colonne) - 4)
                If Not q = 0 Then
                    z(datesCum) = z(datesCum) + zj / Math.Sqrt(q)
                End If
            Next colonne
            z(datesCum) = z(datesCum) / Math.Sqrt(SAR.GetLength(1))
        Next datesCum
        Return z
    End Function

    Private Function patellCalcStatAAR(ByRef SAR(,) As Double, ByRef Mj() As Integer) As Double()
        'Calcul de ASAR (event study tools)
        Dim ASAR(SAR.GetUpperBound(0)) As Double
        For i = 0 To SAR.GetUpperBound(0)
            For colonne = 0 To SAR.GetUpperBound(1)
                If Not SAR(i, colonne) = -2146826246 Then
                    ASAR(i) = ASAR(i) + SAR(i, colonne)
                End If
            Next colonne
        Next i

        'Calcul (s_ASAR)²
        Dim sASARCarre As Double
        For colonne = 0 To SAR.GetUpperBound(1)
            sASARCarre = sASARCarre + (Mj(colonne) - 2) / (Mj(colonne) - 4)
        Next colonne

        'Calcul des statistiques de l'hypothèse AAR = 0
        Dim testHypAAR(SAR.GetUpperBound(0)) As Double
        For i = 0 To SAR.GetUpperBound(0)
            testHypAAR(i) = ASAR(i) / Math.Sqrt(sASARCarre)
        Next i
        Return testHypAAR
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
