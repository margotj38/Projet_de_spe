
''' <summary>
''' Module de gestion des opérations (création et traitement) sur les AR.
''' </summary>
''' <remarks></remarks>
Module RentaAnormales

    ''' <summary>
    ''' Calcule les AR avec le modèle considéré.
    ''' </summary>
    ''' <param name="tabRentaMarcheEst"> Rentabilités de marché sur la période d'estimation. </param>
    ''' <param name="tabRentaMarcheEv"> Rentabilités de marché sur la période d'événement. </param>
    ''' <param name="tabRentaEst"> Rentabilités des entreprises sur la période d'estimation. </param>
    ''' <param name="tabRentaEv"> Rentabilités des entreprises sur la période d'événement. </param>
    ''' <param name="tabAREst"> (Sortie) AR sur la période d'estimation. </param>
    ''' <param name="tabAREv"> (Sortie) AR sur la période d'événement. </param>
    ''' <param name="tabDateEst"> (Sortie) Dates correspondantes sur la période d'estimation. </param>
    ''' <param name="tabDateEv"> (Sortie) Dates correspondantes sur la période d'événement. </param>
    ''' <remarks></remarks>
    Public Sub calculAR(ByRef tabRentaMarcheEst(,) As Double, ByRef tabRentaMarcheEv(,) As Double, ByRef tabRentaEst(,) As Double, _
                             ByRef tabRentaEv(,) As Double, ByRef tabAREst(,) As Double, ByRef tabAREv(,) As Double, _
                             ByRef tabDateEst() As Integer, ByRef tabDateEv() As Integer)

        'On appelle une fonction selon le modèle choisi
        Select Case Globals.Ribbons.Ruban.selFenetres.modele
            Case 0
                modeleMoyenne(tabRentaEst, tabRentaEv, tabAREst, tabAREv, tabDateEst, tabDateEv)
            Case 1
                modeleMarcheSimple(tabRentaMarcheEst, tabRentaMarcheEv, tabRentaEst, tabRentaEv, tabAREst, tabAREv, tabDateEst, tabDateEv)
            Case 2
                'Création des tableaux pour pouvoir faire les régressions en tenant compte des N/A
                Dim tabRentaReg(,,)() = OpPrixRenta.constructionTableauxReg(OpPrixRenta.maxRentAbs, tabRentaEst, tabRentaMarcheEst)
                modeleMarche(tabRentaMarcheEst, tabRentaMarcheEv, tabRentaEst, tabRentaEv, tabRentaReg, tabAREst, tabAREv, tabDateEst, tabDateEv)
            Case Else
                MsgBox("Erreur interne : numero de modèle incorrect dans ChoixSeuilFenetre", 16)
        End Select

    End Sub

    '***************************** Les différents modèles d'estimation des AR *****************************

    ''' <summary>
    ''' Estimation des AR à partir du modèle des rentabilités moyennes : K = moyenne des rentabilités
    ''' </summary>
    ''' <param name="tabRentaEst"> Rentabilités des entreprises sur la période d'estimation. </param>
    ''' <param name="tabRentaEv"> Rentabilités des entreprises sur la période d'événement. </param>
    ''' <param name="tabAREst"> (Sortie) AR sur la période d'estimation. </param>
    ''' <param name="tabAREv"> (Sortie) AR sur la période d'événement. </param>
    ''' <param name="tabDateEst"> (Sortie) Dates correspondantes sur la période d'estimation. </param>
    ''' <param name="tabDateEv"> (Sortie) Dates correspondantes sur la période d'événement. </param>
    ''' <remarks> Attention : la premiere colonne de tabRentaEst et de tabRentaEv sont les dates. </remarks>
    Public Sub modeleMoyenne(ByRef tabRentaEst(,) As Double, ByRef tabRentaEv(,) As Double, _
                             ByRef tabAREst(,) As Double, ByRef tabAREv(,) As Double, _
                             ByRef tabDateEst() As Integer, ByRef tabDateEv() As Integer)

        'On dimensionne les tableaux de AR
        'On ne range pas les dates d'événement dans les tableaux de AR
        ReDim tabAREst(tabRentaEst.GetUpperBound(0), tabRentaEst.GetUpperBound(1) - 1)
        ReDim tabAREv(tabRentaEv.GetUpperBound(0), tabRentaEv.GetUpperBound(1) - 1)

        'Et ceux de dates
        ReDim tabDateEst(tabRentaEst.GetUpperBound(0))
        ReDim tabDateEv(tabRentaEv.GetUpperBound(0))

        'On range les dates dans les tableaux de dates
        'Pour la période d'estimation
        For indDate = 0 To tabRentaEst.GetUpperBound(0)
            tabDateEst(indDate) = tabRentaEst(indDate, 0)
        Next indDate
        'Puis pour la période d'événement
        For indDate = 0 To tabRentaEv.GetUpperBound(0)
            tabDateEv(indDate) = tabRentaEv(indDate, 0)
        Next indDate

        'Tableau des moyennes de chaque titre
        Dim tabMoy(tabRentaEst.GetUpperBound(1) - 1) As Double

        'Calcul des moyennes sur la fenêtre d'estimation
        'Variable pour savoir si un #N/A précédait
        Dim prixPresent As Integer = 1
        For colonne = 1 To tabRentaEst.GetUpperBound(1)
            For i = 0 To tabRentaEst.GetUpperBound(0)
                'S'il y a un NA, on incrémente prixPresent
                If Double.IsNaN(tabRentaEst(i, colonne)) Then
                    prixPresent = prixPresent + 1
                Else
                    'Sinon on somme en multipliant par le nombre de #N/A présents + 1 (ie prixPresent)
                    tabMoy(colonne - 1) = tabMoy(colonne - 1) + tabRentaEst(i, colonne) * prixPresent
                    prixPresent = 1
                End If
            Next i
            'On divise par la taille de la fenêtre d'estimation moins le nombre de #N/A finaux (ie prixPresent - 1)
            If Not tabRentaEst.GetLength(0) - (prixPresent - 1) = 0 Then
                tabMoy(colonne - 1) = tabMoy(colonne - 1) / (tabRentaEst.GetLength(0) - (prixPresent - 1))
            End If
            prixPresent = 1
        Next colonne

        'Calcul des AR sur la fenêtre d'estimation
        For colonne = 1 To tabRentaEst.GetUpperBound(1)
            For i = 0 To tabRentaEst.GetUpperBound(0)
                If Double.IsNaN(tabRentaEst(i, colonne)) Then
                    tabAREst(i, colonne - 1) = Double.NaN
                Else
                    'On obtient des AR sur une période
                    tabAREst(i, colonne - 1) = (tabRentaEst(i, colonne) - tabMoy(colonne - 1))
                End If
            Next i
        Next colonne

        'Calcul des AR sur la fenêtre d'événement
        For colonne = 1 To tabRentaEv.GetUpperBound(1)
            For i = 0 To tabRentaEv.GetUpperBound(0)
                If Double.IsNaN(tabRentaEv(i, colonne)) Then
                    tabAREv(i, colonne - 1) = Double.NaN
                Else
                    'On obtient des AR sur une période
                    tabAREv(i, colonne - 1) = (tabRentaEv(i, colonne) - tabMoy(colonne - 1))
                End If
            Next i
        Next colonne
    End Sub

    ''' <summary>
    ''' Estimation des AR à partir du modèle de marché simplifié : K = Rm
    ''' </summary>
    ''' <param name="tabRentaMarcheEst"> Rentabilités de marché sur la période d'estimation. </param>
    ''' <param name="tabRentaMarcheEv"> Rentabilités de marché sur la période d'événement. </param>
    ''' <param name="tabRentaEst"> Rentabilités des entreprises sur la période d'estimation. </param>
    ''' <param name="tabRentaEv"> Rentabilités des entreprises sur la période d'événement. </param>
    ''' <param name="tabAREst"> (Sortie) AR sur la période d'estimation. </param>
    ''' <param name="tabAREv"> (Sortie) AR sur la période d'événement. </param>
    ''' <param name="tabDateEst"> (Sortie) Dates correspondantes sur la période d'estimation. </param>
    ''' <param name="tabDateEv"> (Sortie) Dates correspondantes sur la période d'événement. </param>
    ''' <remarks> Attention : la premiere colonne de tabRentaEst et de tabRentaEv sont les dates. </remarks>
    Public Sub modeleMarcheSimple(ByRef tabRentaMarcheEst(,) As Double, ByRef tabRentaMarcheEv(,) As Double, _
                                ByRef tabRentaEst(,) As Double, ByRef tabRentaEv(,) As Double, _
                                ByRef tabAREst(,) As Double, ByRef tabAREv(,) As Double, _
                                ByRef tabDateEst() As Integer, ByRef tabDateEv() As Integer)

        'On dimensionne les tableaux de AR
        'On ne range pas les dates d'événement dans les tableaux de AR
        ReDim tabAREst(tabRentaEst.GetUpperBound(0), tabRentaEst.GetUpperBound(1) - 1)
        ReDim tabAREv(tabRentaEv.GetUpperBound(0), tabRentaEv.GetUpperBound(1) - 1)

        'Et ceux de dates
        ReDim tabDateEst(tabRentaEst.GetUpperBound(0))
        ReDim tabDateEv(tabRentaEv.GetUpperBound(0))

        'On range les dates dans les tableaux de dates
        'Pour la période d'estimation
        For indDate = 0 To tabRentaEst.GetUpperBound(0)
            tabDateEst(indDate) = tabRentaEst(indDate, 0)
        Next indDate
        'Puis pour la période d'événement
        For indDate = 0 To tabRentaEv.GetUpperBound(0)
            tabDateEv(indDate) = tabRentaEv(indDate, 0)
        Next indDate

        For colonne = 1 To tabRentaEst.GetUpperBound(1)
            'remplissage des AR sur la fenetre d'estimation
            For i = 0 To tabRentaEst.GetUpperBound(0)
                If Double.IsNaN(tabRentaEst(i, colonne)) Then
                    tabAREst(i, colonne - 1) = Double.NaN
                Else
                    tabAREst(i, colonne - 1) = tabRentaEst(i, colonne) - tabRentaMarcheEst(i, colonne)
                End If
            Next i

            'remplissage des AR sur la fenetre d'événement
            For i = 0 To tabRentaEv.GetUpperBound(0)
                If Double.IsNaN(tabRentaEv(i, colonne)) Then
                    tabAREv(i, colonne - 1) = Double.NaN
                Else
                    tabAREv(i, colonne - 1) = tabRentaEv(i, colonne) - tabRentaMarcheEv(i, colonne)
                End If
            Next i
        Next
    End Sub

    
    ''' <summary>
    ''' Estimation des AR à partir du modèle de marché classique : K = alpha + beta*Rm
    ''' </summary>
    ''' <param name="tabRentaMarcheEst"> Rentabilités de marché sur la période d'estimation. </param>
    ''' <param name="tabRentaMarcheEv"> Rentabilités de marché sur la période d'événement. </param>
    ''' <param name="tabRentaEst"> Rentabilités des entreprises sur la période d'estimation. </param>
    ''' <param name="tabRentaEv"> Rentabilités des entreprises sur la période d'événement. </param>
    ''' <param name="tabRentaReg"> Structure de données particulière pour organiser les données de rentabilités
    ''' de façon à pouvoir réaliser différentes régressions linéaires cohérentes. </param>
    ''' <param name="tabAREst"> (Sortie) AR sur la période d'estimation. </param>
    ''' <param name="tabAREv"> (Sortie) AR sur la période d'événement. </param>
    ''' <param name="tabDateEst"> (Sortie) Dates correspondantes sur la période d'estimation. </param>
    ''' <param name="tabDateEv"> (Sortie) Dates correspondantes sur la période d'événement. </param>
    ''' <remarks> Attention : la premiere colonne de tabRentaEst et de tabRentaEv sont les dates. </remarks>
    Public Sub modeleMarche(ByRef tabRentaMarcheEst(,) As Double, ByRef tabRentaMarcheEv(,) As Double, _
                                ByRef tabRentaEst(,) As Double, ByRef tabRentaEv(,) As Double, ByRef tabRentaReg(,,)() As Double, _
                                ByRef tabAREst(,) As Double, ByRef tabAREv(,) As Double, _
                                ByRef tabDateEst() As Integer, ByRef tabDateEv() As Integer)

        'On dimensionne les tableaux de AR
        'On ne range pas les dates d'événement dans les tableaux de AR
        ReDim tabAREst(tabRentaEst.GetUpperBound(0), tabRentaEst.GetUpperBound(1) - 1)
        ReDim tabAREv(tabRentaEv.GetUpperBound(0), tabRentaEv.GetUpperBound(1) - 1)

        'Et ceux de dates
        ReDim tabDateEst(tabRentaEst.GetUpperBound(0))
        ReDim tabDateEv(tabRentaEv.GetUpperBound(0))

        'On range les dates dans les tableaux de dates
        'Pour la période d'estimation
        For indDate = 0 To tabRentaEst.GetUpperBound(0)
            tabDateEst(indDate) = tabRentaEst(indDate, 0)
        Next indDate
        'Puis pour la période d'événement
        For indDate = 0 To tabRentaEv.GetUpperBound(0)
            tabDateEv(indDate) = tabRentaEv(indDate, 0)
        Next indDate

        'nombre de différentes régressions
        Dim nbReg = tabRentaReg.GetLength(1)
        'déclaration des tableaux contenant les alpha et beta de la régression
        Dim a(nbReg) As Double
        Dim b(nbReg) As Double
        'moyenne pondérée pour obtenir les véritables alpha et beta
        Dim alpha As Double = 0
        Dim beta As Double = 0
        'pour chaque entreprise...
        For colonne = 1 To tabRentaReg.GetUpperBound(0)
            'nombre de rentabilités totale (sans NA)
            Dim nbRent As Integer = 0
            'pour chaque tableau
            For reg = 0 To nbReg - 1
                If Not tabRentaReg(colonne - 1, reg, 0).GetLength(0) = 0 Then
                    'extraction des Rt
                    Dim Y() As Double = tabRentaReg(colonne - 1, reg, 0)
                    Dim X() As Double = tabRentaReg(colonne - 1, reg, 1)
                    'calcul des coefficients des différentes régressions
                    a(reg) = Globals.ThisAddIn.Application.WorksheetFunction.Index(Globals.ThisAddIn.Application.WorksheetFunction.LinEst(Y, X), 2) / (reg + 1)
                    b(reg) = Globals.ThisAddIn.Application.WorksheetFunction.Index(Globals.ThisAddIn.Application.WorksheetFunction.LinEst(Y, X), 1) / (reg + 1)
                    'somme pondérée
                    alpha = alpha + a(reg) * tabRentaReg(colonne - 1, reg, 1).GetLength(0)
                    beta = beta + b(reg) * tabRentaReg(colonne - 1, reg, 1).GetLength(0)
                    nbRent = nbRent + tabRentaReg(colonne - 1, reg, 1).GetLength(0)
                End If
            Next
            'moyenne pondérée
            alpha = alpha / nbRent
            beta = beta / nbRent

            'remplissage des AR sur la fenetre d'estimation
            For i = 0 To tabRentaEst.GetUpperBound(0)
                If Double.IsNaN(tabRentaEst(i, colonne)) Then
                    tabAREst(i, colonne - 1) = Double.NaN
                Else
                    tabAREst(i, colonne - 1) = (tabRentaEst(i, colonne) - (alpha + beta * tabRentaMarcheEst(i, colonne)))
                End If
            Next i

            'remplissage des AR sur la fenetre d'événement
            For i = 0 To tabRentaEv.GetUpperBound(0)
                If Double.IsNaN(tabRentaEv(i, colonne)) Then
                    tabAREv(i, colonne - 1) = Double.NaN
                Else
                    tabAREv(i, colonne - 1) = (tabRentaEv(i, colonne) - (alpha + beta * tabRentaMarcheEv(i, colonne)))
                End If
            Next i
        Next
    End Sub


    '***************************** Opérations sur les AR *****************************
    ''' <summary>
    ''' Fonction qui calcule les moyennes empiriques (sur le temps) des AR sur la fenêtre d'estimation.
    ''' </summary>
    ''' <param name="tabEstAR"> AR calculés sur la fenêtre d'estimation. </param>
    ''' <returns> Tableau des moyennes des AR pour chaque entreprise. </returns>
    ''' <remarks></remarks>
    Public Function moyEstAR(ByRef tabEstAR(,) As Double) As Double()
        Dim tabMoy(tabEstAR.GetUpperBound(1)) As Double
        'Variable pour savoir si un #N/A précédait
        Dim ArPresent As Integer = 1
        For colonne = 0 To tabEstAR.GetUpperBound(1)
            tabMoy(colonne) = 0
            For i = 0 To tabEstAR.GetUpperBound(0)
                'S'il y a un NA, on incrémente ArPresent
                If Double.IsNaN(tabEstAR(i, colonne)) Then
                    ArPresent = ArPresent + 1
                Else
                    'Sinon on somme en multipliant par le nombre de #N/A présents + 1 (ie prixPresent)
                    tabMoy(colonne) = tabMoy(colonne) + tabEstAR(i, colonne) * ArPresent
                    ArPresent = 1
                End If
            Next i
            'On divise par la taille de la fenêtre d'estimation moins le nombre de #N/A finaux (ie prixPresent - 1)
            If Not tabEstAR.GetLength(0) - (ArPresent - 1) = 0 Then
                tabMoy(colonne) = tabMoy(colonne) / (tabEstAR.GetLength(0) - (ArPresent - 1))
            End If
            ArPresent = 1
        Next colonne
        Return tabMoy
    End Function

    ''' <summary>
    ''' Fonction qui calcule les variances empiriques (sur le temps) des AR sur la fenêtre d'estimation.
    ''' </summary>
    ''' <param name="tabEstAR"> AR calculés sur la fenêtre d'estimation. </param>
    ''' <param name="tabMoy"> Moyennes pour chaque entreprise calculées au préalable. </param>
    ''' <returns> Tableau des variances empiriques des AR pour chaque entreprise. </returns>
    ''' <remarks></remarks>
    Public Function varEstAR(ByRef tabEstAR(,) As Double, ByRef tabMoy As Double()) As Double()
        'Calcul des variances sur la fenêtre d'estimation
        Dim tabVar(tabEstAR.GetUpperBound(1)) As Double
        'Variable pour savoir si un #N/A précédait
        Dim ArPresent As Integer = 1
        For colonne = 0 To tabEstAR.GetUpperBound(1)
            tabVar(colonne) = 0
            For i = 0 To tabEstAR.GetUpperBound(0)
                'S'il y a un NA, on incrémente ArPresent
                If Double.IsNaN(tabEstAR(i, colonne)) Then
                    ArPresent = ArPresent + 1
                Else
                    'Sinon on somme en multipliant par le nombre de #N/A présents + 1 (ie prixPresent)
                    tabVar(colonne) = tabVar(colonne) + Math.Pow(((tabEstAR(i, colonne) / ArPresent) - tabMoy(colonne)), 2) * ArPresent
                    ArPresent = 1
                End If
            Next i
            'On divise par la taille de la fenêtre d'estimation moins le nombre de #N/A finaux (ie prixPresent - 1) - 1
            If Not tabEstAR.GetLength(0) - (ArPresent - 1) - 1 = 0 Then
                tabVar(colonne) = tabVar(colonne) / (tabEstAR.GetLength(0) - (ArPresent - 1) - 1)
            End If
            ArPresent = 1
        Next colonne
        Return tabVar
    End Function

    ''' <summary>
    ''' Calcule la moyenne empirique après avoir normalisé les AR.
    ''' </summary>
    ''' <param name="tabEstAR"> AR calculés sur la période d'estimation. </param>
    ''' <param name="tabEvAR"> AR calculés sur la période d'événement. </param>
    ''' <returns> Moyenne empirique des AR normalisés en chaque temps de la fenêtre d'événement. </returns>
    ''' <remarks></remarks>
    Public Function moyNormAR(ByRef tabEstAR(,) As Double, ByRef tabEvAR(,) As Double) As Double()
        Dim tailleFenetreEv As Integer = tabEvAR.GetLength(0)
        'tableau à retourner
        Dim tabMoyNormAR(tailleFenetreEv - 1) As Double
        'Récupération des variances sur la fenêtre d'estimation 
        Dim tabMoyAR() As Double = moyEstAR(tabEstAR)
        Dim tabVarAR() As Double = varEstAR(tabEstAR, tabMoyAR)
        'remplissage du tableau
        For i = 0 To tabEvAR.GetUpperBound(0)
            'tableau des ARi/si (i.e AR normalisés)
            Dim tabNormAR(tabEvAR.GetLength(1) - 1) As Double
            For e = 0 To tabEvAR.GetUpperBound(1)
                'Gestion des NA dans le tableau des AR
                If Double.IsNaN(tabEvAR(i, e)) Then
                    tabNormAR(e) = Double.NaN
                Else
                    'normalisation
                    tabNormAR(e) = tabEvAR(i, e) / Math.Sqrt(tabVarAR(e))
                End If
            Next
            'moyenne sur les ARi/si
            tabMoyNormAR(i) = Utilitaires.calcul_moyenne(tabNormAR)
        Next
        Return tabMoyNormAR
    End Function

    ''' <summary>
    ''' Calcule la variance empirique après avoir normalisé les AR.
    ''' </summary>
    ''' <param name="tabEstAR"> AR calculés sur la période d'estimation. </param>
    ''' <param name="tabEvAR"> AR calculés sur la période d'événement. </param>
    ''' <param name="tabMoyNormAR"> Moyenne des AR normalisés déjà calculée au préalable. </param>
    ''' <returns> Variance empirique des AR normalisés en chaque temps de la fenêtre d'événement. </returns>
    ''' <remarks></remarks>
    Public Function ecartNormAR(ByRef tabEstAR(,) As Double, ByRef tabEvAR(,) As Double, ByRef tabMoyNormAR As Double()) As Double()
        Dim tailleFenetreEv As Integer = tabEvAR.GetLength(0)
        'tableau à retourner
        Dim tabEcartNormAR(tailleFenetreEv - 1) As Double
        'Récupération des variances sur la fenêtre d'estimation 
        Dim tabMoyAR() As Double = moyEstAR(tabEstAR)
        Dim tabVarAR() As Double = varEstAR(tabEstAR, tabMoyAR)
        'remplissage du tableau
        For i = 0 To tabEvAR.GetUpperBound(0)
            'tableau des ARi/si (i.e AR normalisés)
            Dim tabNormAR(tabEvAR.GetLength(1) - 1) As Double
            For e = 0 To tabEvAR.GetUpperBound(1)
                'Gestion des NA dans le tableau des AR
                If Double.IsNaN(tabEvAR(i, e)) Then
                    tabNormAR(e) = Double.NaN
                Else
                    'normalisation
                    tabNormAR(e) = tabEvAR(i, e) / Math.Sqrt(tabVarAR(e))
                End If
            Next
            'écart-type sur les ARi/si
            tabEcartNormAR(i) = Math.Sqrt(Utilitaires.calcul_variance(tabNormAR, tabMoyNormAR(i)))
        Next
        Return tabEcartNormAR
    End Function

    ' ''' <summary>
    ' ''' Calcule la variance des AR par entreprise sur la période d'estimation.
    ' ''' </summary>
    ' ''' <param name="tabEstAR"> AR calculés sur la période d'estimation. </param>
    ' ''' <returns> Variance empirique des AR sur la passé pour chaque entreprise. </returns>
    ' ''' <remarks></remarks>
    'Public Function calcVarEstAR(ByRef tabEstAR(,) As Double) As Double()
    '    'tableau à retourner
    '    Dim tabVarAR(tabEstAR.GetLength(1) - 1) As Double

    '    'pour chaque entreprise...
    '    For e = 0 To tabEstAR.GetUpperBound(1)
    '        Dim vectAR(tabEstAR.GetLength(0) - 1) As Double
    '        For t = 0 To tabEstAR.GetUpperBound(0)
    '            vectAR(t) = tabEstAR(t, e)
    '        Next
    '        tabVarAR(e) = Utilitaires.calcul_variance(vectAR, Utilitaires.calcul_moyenne(vectAR))
    '    Next
    '    Return tabVarAR
    'End Function

    ''' <summary>
    ''' Fonction qui calcule les moyennes de AR (AAR) en chaque temps de la fenêtre d'événement.
    ''' </summary>
    ''' <param name="tabEvAR"> AR sur la fenêtre d'événement. </param>
    ''' <returns> Moyenne des AR sur les entreprises en chaque temps de la fenêtre d'événement. </returns>
    ''' <remarks></remarks>
    Public Function moyAR(ByRef tabEvAR(,) As Double) As Double()
        Dim tailleFenetreEv As Integer = tabEvAR.GetLength(0)
        'tableau à retourner
        Dim tabMoyAR(tailleFenetreEv - 1) As Double
        'remplissage du tableau
        For t = 0 To tabEvAR.GetUpperBound(0)
            'tableau des ARi
            Dim tabAR(tabEvAR.GetUpperBound(1)) As Double
            For i = 0 To tabEvAR.GetUpperBound(1)
                'extraction des ARi
                tabAR(i) = tabEvAR(t, i)
            Next
            'moyenne sur les ARi/si
            tabMoyAR(t) = Utilitaires.calcul_moyenne(tabAR)
        Next
        Return tabMoyAR
    End Function


    '********************************************** Opérations sur les CAR et CAAR **********************************************

    ''' <summary>
    ''' Calcule les CAR sur la fenêtre d'événement.
    ''' </summary>
    ''' <param name="tabEvAR"> AR calculés sur la période d'événement. </param>
    ''' <returns> CAR par entreprise en chaque temps cumulé sur la fenêtre d'événement. </returns>
    ''' <remarks></remarks>
    Function CalculCar(ByRef tabEvAR(,) As Double) As Double(,)
        Dim tailleFenetreEv As Integer = tabEvAR.GetLength(0)
        Dim N As Integer = tabEvAR.GetLength(1)

        'tableau à retourner
        Dim tabCAR(tailleFenetreEv - 1, N - 1) As Double

        'Variable pour savoir si un #N/A précédait
        Dim prixPresent As Integer = 1
        For e = 0 To tabEvAR.GetUpperBound(1)
            Dim somme As Double = 0
            Dim onlyNan = True
            For i = 0 To tabEvAR.GetUpperBound(0)
                If Double.IsNaN(tabEvAR(i, e)) Then
                    'S'il y a un NA, on incrémente prixPresent
                    prixPresent = prixPresent + 1
                Else
                    onlyNan = False
                    'Sinon on somme en multipliant par le nombre de #N/A présents + 1 (ie prixPresent)
                    somme = somme + tabEvAR(i, e) * prixPresent
                    prixPresent = 1
                End If
                'En fonction de si on a eu des données pour cumuler les AR...
                If onlyNan Then
                    tabCAR(i, e) = Double.NaN
                Else
                    tabCAR(i, e) = somme
                End If
            Next
            prixPresent = 1
        Next

        CalculCar = tabCAR
    End Function

    ''' <summary>
    ''' Calcule les CAAR sur la fenêtre d'événement.
    ''' </summary>
    ''' <param name="tabAAR"> Tableau des AAR sur la période d'événement. </param>
    ''' <returns> CAAR en chaque temps cumulé sur la fenêtre d'événement. </returns>
    ''' <remarks></remarks>
    Function CalculCAAR(ByRef tabAAR() As Double) As Double()
        Dim tabCAAR(tabAAR.GetUpperBound(0)) As Double
        tabCAAR(0) = tabAAR(0)
        For i = 1 To tabAAR.GetUpperBound(0)
            tabCAAR(i) = tabCAAR(i - 1) + tabAAR(i)
        Next
        Return tabCAAR
    End Function

    ''' <summary>
    ''' Calcule la moyenne empirique après avoir normalisé les CAR.
    ''' </summary>
    ''' <param name="tabEstAR"> AR calculés sur la période d'estimation. </param>
    ''' <param name="tabCAR"> CAR par entreprise en chaque temps cumulé sur la fenêtre d'événement. </param>
    ''' <returns> Moyenne empirique des CAR normalisés en chaque temps cumulé de la fenêtre d'événement. </returns>
    ''' <remarks></remarks>
    Function moyNormCar(ByRef tabEstAR(,) As Double, ByRef tabCAR(,) As Double) As Double()

        Dim tailleFenetreEv As Integer = tabCAR.GetLength(0)
        'tableau à retourner
        Dim tabMoyNormCAR(tailleFenetreEv - 1) As Double
        'récupération des  des variances
        Dim tabMoyAR() As Double = moyEstAR(tabEstAR)
        Dim tabVarAR() As Double = varEstAR(tabEstAR, tabMoyAR)

        For i = 0 To tabCAR.GetUpperBound(0)
            Dim tabNormCAR(tabCAR.GetUpperBound(1)) As Double
            For e = 0 To tabCAR.GetUpperBound(1)
                'Gestion des NA dans le tableau des AR
                If Double.IsNaN(tabCAR(i, e)) Then
                    tabNormCAR(e) = Double.NaN
                Else
                    'normalisation du CAR sur i+1 périodes
                    tabNormCAR(e) = tabCAR(i, e) / ((i + 1) * Math.Sqrt(tabVarAR(e)))
                End If
            Next
            tabMoyNormCAR(i) = Utilitaires.calcul_moyenne(tabNormCAR)
        Next
        Return tabMoyNormCAR
    End Function

    ''' <summary>
    ''' Calcule la variance empirique après avoir normalisé les CAR.
    ''' </summary>
    ''' <param name="tabEstAR"> AR calculés sur la période d'estimation. </param>
    ''' <param name="tabCAR"> CAR par entreprise en chaque temps cumulé sur la fenêtre d'événement. </param>
    ''' <param name="tabMoy"> Moyenne des CAR normalisés déjà calculée au préalable. </param>
    ''' <returns> Variance empirique des CAR normalisés en chaque temps cumulé de la fenêtre d'événement. </returns>
    ''' <remarks></remarks>
    Function ecartNormCar(ByRef tabEstAR(,) As Double, tabCAR(,) As Double, tabMoy() As Double) As Double()
        Dim tailleFenetreEv As Integer = tabCAR.GetLength(0)
        Dim tabEcartNormCAR(tailleFenetreEv - 1) As Double

        'récupération des  des variances
        Dim tabMoyAR() As Double = moyEstAR(tabEstAR)
        Dim tabVarAR() As Double = varEstAR(tabEstAR, tabMoyAR)

        For i = 0 To tailleFenetreEv - 1
            Dim tabNormCAR(tabCAR.GetLength(1) - 1) As Double
            For e = 0 To tabCAR.GetUpperBound(1)
                'Gestion des NA dans le tableau des AR
                If Double.IsNaN(tabCAR(i, e)) Then
                    tabNormCAR(e) = Double.NaN
                Else
                    'normalisation du CAR sur i+1 périodes
                    tabNormCAR(e) = tabCAR(i, e) / ((i + 1) * Math.Sqrt(tabVarAR(e)))
                End If
            Next
            tabEcartNormCAR(i) = Math.Sqrt(Utilitaires.calcul_variance(tabNormCAR, tabMoy(i)))
        Next
        Return tabEcartNormCAR
    End Function

    ''' <summary>
    ''' Produit et affiche les résultats des tests statiques sur les AR et les CAR
    ''' </summary>
    ''' <param name="tabEvAR"> AR calculés sur la période d'estimation. </param>
    ''' <param name="tabEstAR"> AR calculés sur la période d'événement. </param>
    ''' <param name="datesEvAR"> Dates correspondantes sur la période d'événement. </param>
    ''' <remarks></remarks>
    Public Sub traitementTabAR(tabEvAR(,) As Double, tabEstAR(,) As Double, datesEvAR() As Integer)

        'La création d'une nouvelle feuille
        Dim nom As String
        nom = InputBox("Entrer Le nom de la feuille des résultats de l'étude d'événements: ")
        'Si l'utilisateur n'entre pas un nom
        If nom Is "" Then nom = "Resultat"
        Globals.ThisAddIn.Application.Sheets.Add(After:=Globals.ThisAddIn.Application.Worksheets(Globals.ThisAddIn.Application.Worksheets.Count))
        Globals.ThisAddIn.Application.ActiveSheet.Name = nom

        Dim N As Integer = tabEvAR.GetLength(1)

        '----------------- AR -----------------
        'tableau des AR moyen normalisés
        Dim tabMoyAR() As Double = RentaAnormales.moyNormAR(tabEstAR, tabEvAR)
        'tableau des écart-types des AR normalisés
        Dim tabEcartAR() As Double = RentaAnormales.ecartNormAR(tabEstAR, tabEvAR, tabMoyAR)

        Dim statAR() As Double = TestsStatistiques.calculStatStudent(tabMoyAR, tabEcartAR, N)

        '----------------- CAR -----------------
        Dim tailleFenetreEv As Integer = tabEvAR.GetLength(0)
        'Remplissage des tableaux : CAR, moyenne, variance
        Dim tabCAR(,) As Double = RentaAnormales.CalculCar(tabEvAR)
        Dim tabMoyCAR() As Double = RentaAnormales.moyNormCar(tabEstAR, tabCAR)
        Dim tabEcartCAR() As Double = RentaAnormales.ecartNormCar(tabEstAR, tabCAR, tabMoyCAR)

        Dim statCAR() As Double = TestsStatistiques.calculStatStudent(tabMoyCAR, tabEcartCAR, N)

        'affichage des résultats des CAR
        ExcelDialogue.afficheResAsympt(datesEvAR, tabMoyAR, tabEcartAR, statAR, tabMoyCAR, tabEcartCAR, statCAR, N, nom, 0)

        '----------------- AAR -----------------
        Dim statAAR() As Double = TestsStatistiques.calculStatStudentAAR(tabEstAR, tabEvAR)

        '----------------- CAAR -----------------
        Dim statCAAR() As Double = TestsStatistiques.calculStatStudentCAAR(tabEstAR, tabEvAR)

        'affichage des résultats des résultats AR
        ExcelDialogue.afficheResExact(datesEvAR, statAAR, statCAAR, N, nom, 9)

    End Sub

End Module
