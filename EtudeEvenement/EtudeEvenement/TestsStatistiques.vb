''' <summary>
''' Modules contenant toutes les fonctions relatives aux différents tests statistiques (test de Student, test de Patell et 
''' test de signe.
''' </summary>
''' <remarks></remarks>
''' 
Module TestsStatistiques

    '***************************** T-Test *****************************


    ''' <summary>
    ''' Fonction qui calcule la statistique de test pour les AAR en chaque temps de la fenêtre d'événement
    ''' </summary>
    ''' <param name="tabEstAR"> AR calculés sur la fenêtre d'estimation. </param>
    ''' <param name="tabEvAR"> AR calculés sur la fenêtre d'événement. </param>
    ''' <returns> Tableau des statistiques de test en tout temps de la fenetre d'événement. </returns>
    ''' <remarks></remarks>
    Public Function calculStatStudentAAR(ByRef tabEstAR(,) As Double, ByRef tabEvAR As Double(,)) As Double()
        Dim stat(tabEvAR.GetLength(0)) As Double
        'Récupération des variances empiriques temporelles sur la fenêtre d'estimation
        Dim tabVar() As Double = RentaAnormales.varEstAR(tabEstAR, RentaAnormales.moyEstAR(tabEstAR))
        'Calcul de la somme des variances empiriques sur la période d'estimation
        Dim sommeVar As Double = 0
        For i = 0 To tabVar.GetUpperBound(0)
            sommeVar = sommeVar + tabVar(i)
        Next
        'Récupération des moyennes des AR sur les entreprises en chaque temps
        Dim tabAAR() As Double = RentaAnormales.moyAR(tabEvAR)
        'Calcul des statistiques de test
        For t = 0 To tabEvAR.GetUpperBound(0)
            stat(t) = tabAAR(t) / (Math.Sqrt(sommeVar) / tabEvAR.GetLength(1))
        Next
        Return stat
    End Function

    ''' <summary>
    ''' Fonction qui calcule la statistique de test pour les CAAR en chaque temps de la fenêtre d'événement
    ''' </summary>
    ''' <param name="tabEstAR"> AR calculés sur la fenêtre d'estimation. </param>
    ''' <param name="tabEvAR"> AR calculés sur la fenêtre d'événement. </param>
    ''' <returns> Tableau des statistiques de test en tout temps de la fenetre d'événement. </returns>
    ''' <remarks></remarks>
    Public Function calculStatStudentCAAR(ByRef tabEstAR(,) As Double, ByRef tabEvAR As Double(,)) As Double()
        Dim stat(tabEvAR.GetLength(0)) As Double
        'Récupération des variances empiriques temporelles sur la fenêtre d'estimation
        Dim tabVar() As Double = RentaAnormales.varEstAR(tabEstAR, RentaAnormales.moyEstAR(tabEstAR))
        'Calcul de la somme des variances empiriques sur la période d'estimation
        Dim sommeVar As Double = 0
        For i = 0 To tabVar.GetUpperBound(0)
            sommeVar = sommeVar + tabVar(i)
        Next
        'Récupération des moyennes des AR sur les entreprises en chaque temps
        Dim tabAAR() As Double = RentaAnormales.moyAR(tabEvAR)
        'Calcul des CAAR
        Dim tabCAAR() As Double = RentaAnormales.CalculCAAR(tabAAR)
        'Calcul des statistiques de test
        For t = 0 To tabCAAR.GetUpperBound(0)
            stat(t) = tabCAAR(t) / (Math.Sqrt(sommeVar) * Math.Sqrt(t + 1) / tabEvAR.GetLength(1))
        Next
        Return stat
    End Function

    ''' <summary>
    ''' Fonction qui calcule la P-Valeur d'un test de Student.
    ''' </summary>
    ''' <param name="testHyp">La statistique d'un test de Student.</param>
    ''' <param name="tailleEchant">La taille de l'échantillon.</param>
    ''' <returns>La P-Valeur du test de Student.</returns>
    ''' <remarks></remarks>
    Public Function calculPValeurStudent(testHyp As Double, tailleEchant As Integer) As Double
        Return Globals.ThisAddIn.Application.WorksheetFunction.T_Dist_2T(Math.Abs(testHyp), tailleEchant - 1)
    End Function


    '***************************** Test de Patell *****************************

    ''' <summary>
    ''' Fonction qui calcule, pour chaque entreprise, le nombre de jours où elle est côtée sur la période d'estimation.
    ''' </summary>
    ''' <param name="tabAREst">Tableau des AR sur la période d'estimation.</param>
    ''' <returns>Pour chaque entreprise, le nombre de jours de cotation sur la période d'estimation.</returns>
    ''' <remarks></remarks>
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
    ''' Procédure qui calcule les statistiques associées au test de Patell.
    ''' </summary>
    ''' <param name="tabAREst">Tableau des AR sur la période d'estimation.</param>
    ''' <param name="tabAREv">Tableau des AR sur la période d'événement.</param>
    ''' <param name="tabDateEst">Tableau des dates de la période d'estimation.</param>
    ''' <param name="tabDateEv">Tableau des dates sur la période d'événement.</param>
    ''' <param name="tabRentaClassiquesMarcheEst">Tableau des rentabilités calculées classiquement (ie pas de la même façon 
    ''' que les rentabilités des entreprises) sur la période d'estimation.</param>
    ''' <param name="tabRentaClassiquesMarcheEv">Tableau des rentabilités calculées classiquement (ie pas de la même façon 
    ''' que les rentabilités des entreprises) sur la période d'événement.</param>
    ''' <param name="Mj">Tableau du nombre de jours de cotation pour chaque entreprise.</param>
    ''' <param name="testHypAAR">(Sortie) Statistiques associées au test d'hypothèse "H0 : AAR = 0" pour chaque date 
    ''' sur la période d'événement.</param>
    ''' <param name="testHypCAAR">(Sortie) Statistiques associées au test d'hypothèse "H0 : CAAR = 0" pour chaque date 
    ''' sur la période d'événement.</param>
    ''' <remarks></remarks>
    Public Sub patellTest(ByRef tabAREst(,) As Double, ByRef tabAREv(,) As Double, _
                               ByRef tabDateEst() As Integer, ByRef tabDateEv() As Integer, _
                               ByRef tabRentaClassiquesMarcheEst(,) As Double, ByRef tabRentaClassiquesMarcheEv(,) As Double, _
                               ByRef Mj() As Integer, ByRef testHypAAR() As Double, ByRef testHypCAAR() As Double)
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

    ''' <summary>
    ''' Fonction calculant la variance sur la période d'estimation pour chaque titre (variance qui doit être ajustée 
    ''' pour être incorporée dans la formule du test de Patell).
    ''' </summary>
    ''' <param name="tabAREst">Tableau des AR sur la période d'estimation.</param>
    ''' <param name="Mj">Tableau du nombre de jours de cotation pour chaque entreprise.</param>
    ''' <returns>Pour chaque titre, la variance sur la période d'estimation.</returns>
    ''' <remarks></remarks>
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

    ''' <summary>
    ''' Fonction qui calcule la moyenne des rentabilités de marché sur la période d'estimation, et ce, pour chaque titre.
    ''' </summary>
    ''' <param name="tabRentaClassiquesMarcheEst">Tableau des rentabilités calculées classiquement (ie pas de la même façon 
    ''' que les rentabilités des entreprises) sur la période d'estimation.</param>
    ''' <returns>Pour chaque entreprise, la moyenne des rentabilités de marché sur la période d'estimation.</returns>
    ''' <remarks></remarks>
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

    ''' <summary>
    ''' Fonction calculant la somme au carré des rentabilités de marché moins leur moyenne, sur la période d'estimation, et ce, 
    ''' pour chaque entreprise.
    ''' </summary>
    ''' <param name="tabRentaClassiquesMarcheEst">Tableau des rentabilités calculées classiquement (ie pas de la même façon 
    ''' que les rentabilités des entreprises) sur la période d'estimation.</param>
    ''' <param name="rmEst">Moyenne des rentabilités de marché sur la période d'estimation pour chaque entreprise.</param>
    ''' <returns>Pour chaque entreprise, la somme au carré des rentabilités de marché moins leur moyenne sur la période 
    ''' d'estimation.</returns>
    ''' <remarks></remarks>
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

    ''' <summary>
    ''' Fonction qui calcule la statistique du test de Patell (H0 : CAAR = 0).
    ''' </summary>
    ''' <param name="SAR">Tableau des AR normalisés pour chaque entreprise, à chaque date de la période d'événement.</param>
    ''' <param name="Mj">Tableau du nombre de jours de cotation pour chaque entreprise.</param>
    ''' <returns>Les statistiques du test de Patell (H0 : CAAR = 0), à chaque date de la période d'événement</returns>
    ''' <remarks></remarks>
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

    ''' <summary>
    ''' Fonction qui calcule la statistique du test de Patell (H0 : AAR = 0).
    ''' </summary>
    ''' <param name="SAR">Tableau des AR normalisés pour chaque entreprise, à chaque date de la période d'événement.</param>
    ''' <param name="Mj">Tableau du nombre de jours de cotation pour chaque entreprise.</param>
    ''' <returns>Les statistiques du test de Patell (H0 : AAR = 0), à chaque date de la période d'événement</returns>
    ''' <remarks></remarks>
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

    ''' <summary>
    ''' Fonction qui calcule la statistique du test de signe pour chaque date de la période d'événement.
    ''' </summary>
    ''' <param name="tabEstAR">Tableau des AR sur la période d'estimation.</param>
    ''' <param name="tabEvAR">Tableau des AR sur la période d'événement.</param>
    ''' <returns>Pour chaque date de la période d'événement, la statistique du test de signe.</returns>
    ''' <remarks></remarks>
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

    ''' <summary>
    ''' Fonction calculant la P-Valeur du test de signe.
    ''' </summary>
    ''' <param name="stat">Statistique du test de signe.</param>
    ''' <returns>La P-Valeur du test de signe.</returns>
    ''' <remarks></remarks>
    Public Function calculPValeurTestSigne(stat As Double) As Double
        Return 2 * (1 - Globals.ThisAddIn.Application.WorksheetFunction.Norm_S_Dist(Math.Abs(stat), True))
    End Function

End Module
