Module RentaAnormales

    'Calcule les AR avec le modèle considéré
    Public Sub calculAR(ByRef tabRentaMarcheEst(,) As Double, ByRef tabRentaMarcheEv(,) As Double, ByRef tabRentaEst(,) As Double, _
                             ByRef tabRentaEv(,) As Double, ByRef tabAREst(,) As Double, ByRef tabAREv(,) As Double, _
                             ByRef tabDateEst() As Integer, ByRef tabDateEv() As Integer)

        'appelle une fonction pour chaque modèle
        Select Case Globals.Ribbons.Ruban.selFenetres.modele
            Case 0
                modeleMoyenne(tabRentaEst, tabRentaEv, tabAREst, tabAREv, tabDateEst, tabDateEv)
            Case 1
                modeleMarcheSimple(tabRentaEst, tabRentaEv, tabRentaMarcheEst, tabRentaMarcheEv, tabAREst, tabAREv, tabDateEst, tabDateEv)
            Case 2
                'Création des tableaux pour pouvoir les X et Y de la régression
                Dim tabRentaReg(,,)() = UtilitaireRentabilites.constructionTableauxReg(UtilitaireRentabilites.maxPrixAbs, tabRentaEst, tabRentaMarcheEst)
                modeleMarche(tabRentaEst, tabRentaEv, tabRentaReg, tabRentaMarcheEst, tabRentaMarcheEv, tabAREst, tabAREv, tabDateEst, tabDateEv)
            Case Else
                MsgBox("Erreur interne : numero de modèle incorrect dans ChoixSeuilFenetre", 16)
        End Select

    End Sub


    '***************************** Modèle de marché *****************************

    'Estimation des AR à partir du modèle de marché : K = alpha + beta*Rm
    'Premiere colonne de tabRentaEst et tabRentaEv : dates
    Public Sub modeleMarche(ByRef tabRentaEst(,) As Double, ByRef tabRentaEv(,) As Double, ByRef tabRentaReg(,,)() As Double, _
                             ByRef tabRentaMarcheEst(,) As Double, ByRef tabRentaMarcheEv(,) As Double, ByRef tabAREst(,) As Double, _
                             ByRef tabAREv(,) As Double, ByRef tabDateEst() As Integer, ByRef tabDateEv() As Integer)

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
                    'Dim Y(rentaEst(reg, colonne).GetUpperBound(1)) As Double
                    'Dim X(rentaEst(reg, colonne).GetUpperBound(1)) As Double
                    'For t = 0 To rentaEst(reg, colonne).GetUpperBound(1)
                    '    Y(t) = rentaEst(reg, colonne)(1, t)
                    '    'extraction des Rm
                    '    X(t) = rentaEst(reg, colonne)(0, t)
                    'Next
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
            'Variable pour savoir si des AR précédents sont manquants
            'Dim prixPresent As Integer = 1
            For i = 0 To tabRentaEst.GetUpperBound(0)
                If tabRentaEst(i, colonne) = -2146826246 Then
                    tabAREst(i, colonne - 1) = -2146826246
                    'prixPresent = prixPresent + 1
                Else
                    tabAREst(i, colonne - 1) = (tabRentaEst(i, colonne) - (alpha + beta * tabRentaMarcheEst(i, colonne))) '* prixPresent
                    'prixPresent = 1
                End If
            Next i

            'remplissage des AR sur la fenetre d'événement
            'Variable pour savoir si des AR précédents sont manquants
            'Dim prixPresent As Integer = 1
            For i = 0 To tabRentaEv.GetUpperBound(0)
                If tabRentaEv(i, colonne) = -2146826246 Then
                    tabAREv(i, colonne - 1) = -2146826246
                    'prixPresent = prixPresent + 1
                Else
                    tabAREv(i, colonne - 1) = (tabRentaEv(i, colonne) - (alpha + beta * tabRentaMarcheEv(i, colonne))) '* prixPresent
                    'prixPresent = 1
                End If
            Next i
        Next
    End Sub


    '***************************** Modèle de marché simplifié *****************************

    'Estimation des AR à partir du modèle de marché simplifié : K = moyenne des rentabilités
    Public Sub modeleMarcheSimple(ByRef tabRentaEst(,) As Double, ByRef tabRentaEv(,) As Double, _
                             ByRef tabRentaMarcheEst(,) As Double, ByRef tabRentaMarcheEv(,) As Double, ByRef tabAREst(,) As Double, _
                             ByRef tabAREv(,) As Double, ByRef tabDateEst() As Integer, ByRef tabDateEv() As Integer)

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
                If tabRentaEst(i, colonne) = -2146826246 Then
                    tabAREst(i, colonne - 1) = -2146826246
                Else
                    tabAREst(i, colonne - 1) = tabRentaEst(i, colonne) - tabRentaMarcheEst(i, colonne)
                End If
            Next i

            'remplissage des AR sur la fenetre d'événement
            For i = 0 To tabRentaEv.GetUpperBound(0)
                If tabRentaEv(i, colonne) = -2146826246 Then
                    tabAREv(i, colonne - 1) = -2146826246
                Else
                    tabAREv(i, colonne - 1) = tabRentaEv(i, colonne) - tabRentaMarcheEv(i, colonne)
                End If
            Next i
        Next
    End Sub


    '***************************** Modèle des rentabilités moyennes *****************************

    'Estimation des AR à partir du modèle de la moyenne : K = R
    'Premiere colonne de tabRentaEst et tabRentaEv : dates
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
                If tabRentaEst(i, colonne) = -2146826246 Then
                    prixPresent = prixPresent + 1
                Else
                    'Sinon on somme en multipliant par le nombre de #N/A absents + 1 (ie prixPresent)
                    tabMoy(colonne - 1) = tabMoy(colonne - 1) + tabRentaEst(i, colonne) * prixPresent
                    prixPresent = 1
                End If
            Next i
            'On divise par la taille de la fenêtre d'estimation moins le nombre de #N/A finaux (ie prixPresent - 1)
            tabMoy(colonne - 1) = tabMoy(colonne - 1) / (tabRentaEst.GetLength(0) - (prixPresent - 1))
        Next colonne

        'Calcul des AR sur la fenêtre sur la fenêtre d'estimation
        For colonne = 1 To tabRentaEst.GetUpperBound(1)
            For i = 0 To tabRentaEst.GetUpperBound(0)
                If tabRentaEst(i, colonne) = -2146826246 Then
                    tabAREst(i, colonne - 1) = -2146826246
                Else
                    'On obtient des AR sur une période
                    tabAREst(i, colonne - 1) = (tabRentaEst(i, colonne) - tabMoy(colonne - 1))
                    prixPresent = 1
                End If
            Next i
        Next colonne

        'Calcul des AR sur la fenêtre sur la fenêtre d'événement
        For colonne = 1 To tabRentaEv.GetUpperBound(1)
            For i = 0 To tabRentaEv.GetUpperBound(0)
                If tabRentaEv(i, colonne) = -2146826246 Then
                    tabAREv(i, colonne - 1) = -2146826246
                Else
                    'On obtient des AR sur une période
                    tabAREv(i, colonne - 1) = (tabRentaEv(i, colonne) - tabMoy(colonne - 1))
                End If
            Next i
        Next colonne
    End Sub

    '***************************** Opérations sur les AR *****************************

    Public Function moyNormAR(ByRef tabEstAR(,) As Object, ByRef tabEvAR(,) As Object) As Double()
        Dim tailleFenetreEv As Integer = tabEvAR.GetLength(0)
        'tableau à retourner
        Dim tabMoyNormAR(tailleFenetreEv - 1) As Double
        'tableau des variances
        Dim tabVarAR() As Double = calcVarEstAR(tabEstAR)
        'remplissage du tableau
        For i = 1 To tabEvAR.GetUpperBound(0)
            'tableau des ARi/si
            Dim tabNormAR(tabEvAR.GetLength(1) - 1) As Double
            For e = 1 To tabEvAR.GetUpperBound(1)
                tabNormAR(e - 1) = tabEvAR(i, e) / Math.Sqrt(tabVarAR(e - 1))
            Next
            'moyenne sur les ARi/si
            tabMoyNormAR(i - 1) = TestsStatistiques.calcul_moyenne(tabNormAR)
        Next
        Return tabMoyNormAR
    End Function

    Public Function ecartNormAR(ByRef tabEstAR(,) As Object, ByRef tabEvAR(,) As Object, ByRef tabMoyNormAR As Double()) As Double()
        Dim tailleFenetreEv As Integer = tabEvAR.GetLength(0)
        'tableau à retourner
        Dim tabEcartNormAR(tailleFenetreEv - 1) As Double
        'tableau des variances
        Dim tabVarAR() As Double = calcVarEstAR(tabEstAR)
        'remplissage du tableau
        For i = 1 To tabEvAR.GetUpperBound(0)
            'tableau des ARi/si
            Dim tabNormAR(tabEvAR.GetLength(1) - 1) As Double
            For e = 1 To tabEvAR.GetUpperBound(1)
                tabNormAR(e - 1) = tabEvAR(i, e) / Math.Sqrt(tabVarAR(e - 1))
            Next
            'moyenne sur les ARi/si
            tabEcartNormAR(i - 1) = Math.Sqrt(TestsStatistiques.calcul_variance(tabNormAR, tabMoyNormAR(i - 1)))
        Next
        Return tabEcartNormAR
    End Function

    'calcule la variance des AR par entreprise sur la période d'estimation pour toutes les entreprises
    Public Function calcVarEstAR(ByRef tabEstAR(,) As Object) As Double()
        'tableau à retourner
        Dim tabVarAR(tabEstAR.GetLength(1) - 1) As Double
        'pour chaque entreprise...
        For e = 1 To tabEstAR.GetUpperBound(1)
            Dim vectAR(tabEstAR.GetLength(0) - 1) As Double
            For t = 1 To tabEstAR.GetUpperBound(0)
                vectAR(t - 1) = CDbl(tabEstAR(t, e))
            Next
            tabVarAR(e - 1) = TestsStatistiques.calcul_variance(vectAR, TestsStatistiques.calcul_moyenne(vectAR))
        Next
        Return tabVarAR
    End Function

    '**********************************************Opérations sur les CAR

    'Fonction qui calcule les CAR sur la fenetre d'événements
    Function CalculCar(ByRef tabEvAR(,) As Object) As Double(,)
        Dim tailleFenetreEv As Integer = tabEvAR.GetLength(0)
        Dim N As Integer = tabEvAR.GetLength(1)

        'tableau à retourner
        Dim tabCAR(tailleFenetreEv - 1, N - 1) As Double

        For e = 1 To N
            Dim somme As Double = 0
            For i = 1 To tailleFenetreEv
                'si RA = NA => On ignore la valeur de RA
                If (tabEvAR(i, e) = -2146826246) Then
                    tabEvAR(i, e) = 0
                End If
                somme = somme + tabEvAR(i, e)
                tabCAR(i - 1, e - 1) = somme
            Next
        Next

        CalculCar = tabCAR
    End Function

    Function calculMoyenneCar(ByRef tabCAR(,) As Double) As Double()
        Dim tailleFenetreEv As Integer = tabCAR.GetLength(0)
        Dim tabMoyCar(tailleFenetreEv - 1) As Double

        For i = 1 To tailleFenetreEv
            Dim tab(tabCAR.GetLength(1) - 1) As Double
            For e = 1 To tabCAR.GetUpperBound(1)
                tab(e - 1) = tabCAR(
                    i - 1, e - 1)
            Next
            tabMoyCar(i - 1) = TestsStatistiques.calcul_moyenne(tab)
        Next
        Return tabMoyCar
    End Function

    Function calculVarianceCar(tabCAR(,) As Double, tabMoy() As Double) As Double()
        Dim tailleFenetreEv As Integer = tabCAR.GetLength(0)
        Dim tabMoyCar(tailleFenetreEv - 1) As Double

        For i = 1 To tailleFenetreEv
            Dim tab(tabCAR.GetLength(1) - 1) As Double
            For e = 1 To tabCAR.GetUpperBound(1)
                tab(e - 1) = tabCAR(i - 1, e - 1)
            Next
            tabMoyCar(i - 1) = TestsStatistiques.calcul_variance(tab, tabMoy(i - 1))
        Next
        Return tabMoyCar
    End Function
End Module
