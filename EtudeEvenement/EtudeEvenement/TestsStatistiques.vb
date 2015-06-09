''' <summary>
''' MODULE STATS
''' </summary>
''' <remarks></remarks>
''' 
Module TestsStatistiques

    '***************************** T-Test *****************************

    Function calcul_moyenne(tab() As Double) As Double
        'tab() peut contenir des NaN
        calcul_moyenne = 0

        Dim nbNaN As Integer = 0
        For i = 0 To tab.GetUpperBound(0)
            If Double.IsNaN(tab(i)) Then
                nbNaN = nbNaN + 1
            Else
                'Sinon on somme
                calcul_moyenne = calcul_moyenne + tab(i)
            End If
        Next i
        'On divise pour obtenir la moyenne en tenant compte des Nan
        If Not (tab.GetLength(0) - nbNaN) = 0 Then
            calcul_moyenne = calcul_moyenne / (tab.GetLength(0) - nbNaN)
        End If
    End Function

    Function calcul_variance(tab() As Double, moyenne As Double) As Double
        'tab() peut contenir des NaN
        calcul_variance = 0

        Dim nbNaN As Integer = 0
        For i = 0 To tab.GetUpperBound(0)
            If Double.IsNaN(tab(i)) Then
                nbNaN = nbNaN + 1
            Else
                Dim tmp As Double = tab(i) - moyenne
                'On somme les différences au carré
                calcul_variance = calcul_variance + tmp * tmp
            End If
        Next i
        'On divise pour obtenir la variance en tenant compte des Nan
        If Not (tab.GetLength(0) - 1 - nbNaN) = 0 Then
            calcul_variance = calcul_variance / (tab.GetLength(0) - 1 - nbNaN)
        End If
    End Function

    Public Function calculStatStudent(moy As Double, ecart As Double, tailleEchant As Integer)
        Return Math.Abs(Math.Sqrt(tailleEchant) * moy / ecart)
    End Function

    Public Function calculPValeurStudent(testHyp As Double, tailleEchant As Integer) As Double
        Return Globals.ThisAddIn.Application.WorksheetFunction.T_Dist_2T(testHyp, tailleEchant - 1)
    End Function


    '***************************** Test de Patell *****************************

    Public Function calculNbNonMissingReturn(ByRef tabAREst(,) As Double) As Integer()
        Dim nbNMR(tabAREst.GetUpperBound(1)) As Integer
        For colonne = 0 To tabAREst.GetUpperBound(1)
            Dim nbMR As Integer = 0
            For i = 0 To tabAREst.GetUpperBound(0)
                'On compte le nombre d'AR manquants
                If Double.IsNaN(tabAREst(i, colonne)) Then
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
                If Double.IsNaN(tabAREv(i, j)) Then
                    SAR(i, j) = Double.NaN
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
                If Not Double.IsNaN(tabAREst(i, colonne)) Then
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
                    If Double.IsNaN(SAR(t, colonne)) Then
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
                If Not Double.IsNaN(SAR(i, colonne)) Then
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

    Function statTestSigne(ByRef tabEstAR(,) As Double, ByRef tabEvAR(,) As Double) As Double()
        Dim tailleFenetreEv As Integer = tabEvAR.GetLength(0)
        Dim tailleFenetreEst As Integer = tabEstAR.GetLength(0)
        Dim N As Integer = tabEvAR.GetLength(1)  'Le nombre des entreprises

        'Tableau qui contient le nombre de AR>0 à une date donnée
        Dim nbARPos(tailleFenetreEv - 1) As Double

        'Compteur des AR > 0
        Dim cmp As Integer

        'On prend les AR > 0 sur la fenêtre d'événement à une date donnée
        For t = 0 To tabEvAR.GetUpperBound(0)
            cmp = 0
            For e = 0 To tabEvAR.GetUpperBound(1)
                If (tabEvAR(t, e) > 0) Then
                    cmp = cmp + 1
                End If
            Next e
            nbARPos(t) = cmp
        Next t

        'Estimation de p sur la fenêtre d'estimation
        Dim p As Double = 0
        Dim nb As Double
        For e = 0 To tabEstAR.GetUpperBound(1)
            nb = 0
            For t = 0 To tabEstAR.GetUpperBound(0)
                If (tabEstAR(t, e) > 0) Then
                    nb = nb + 1
                End If
            Next t
            p = p + (nb / tailleFenetreEst)
        Next e
        p = p / N


        'Calcul de la statistique du test
        Dim tabStatSigne(tailleFenetreEv - 1) As Double
        For t = 0 To tabEvAR.GetUpperBound(0)
            tabStatSigne(t) = (nbARPos(t) - N * p) / (Math.Sqrt(N * p * (1 - p)))
        Next t

        'Retourner la statistique du test
        statTestSigne = tabStatSigne

    End Function


    Public Function calculPValeurTestSigne(stat As Double) As Double
        Return 2 * (1 - Globals.ThisAddIn.Application.WorksheetFunction.Norm_S_Dist(Math.Abs(stat), True))
    End Function


End Module
