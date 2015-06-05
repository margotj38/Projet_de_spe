Module RentaAnormales

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
        '''''''''''''''Comenté pour debug
        'PretraitementPrix.constructionTableauRenta(nbLignes, nbColonnes, maxPrixAbsent, tabRenta, tabRentaMarche)

        'On calcule maintenant les AR
        Dim tailleComplete As Integer = fenetreEstFin - fenetreEstDebut + 1 + fenetreEvFin - fenetreEvDebut + 1
        calculARAvecNA = calculAR(tailleComplete, maxPrixAbsent, fenetreEstDebut, fenetreEstFin, fenetreEvDebut, fenetreEvFin, _
                                  currentSheet.Cells(2, 1).Value + 1, tabRenta, tabRentaMarche)
    End Function

    'Calcule les AR avec le modèle considéré
    Public Function calculAR(tailleComplete As Integer, maxPrixAbsent As Integer, fenetreEstDebut As Integer, _
                             fenetreEstFin As Integer, fenetreEvDebut As Integer, fenetreEvFin As Integer, premiereDate As Integer, Optional tabRenta(,) As Double = Nothing, Optional tabRentaMarche(,) As Double = Nothing) As Double(,)
        Dim tabAR(,) As Double
        'appelle une fonction pour chaque modèle
        Select Case Globals.Ribbons.Ruban.choixSeuilFenetre.modele
            Case 0
                tabAR = modeleMoyenne(tailleComplete, premiereDate, fenetreEstDebut, fenetreEstFin, tabRenta)
            Case 1
                tabAR = modeleMarcheSimple()
            Case 2
                'Création des tableaux pour pouvoir les X et Y de la régression
                Dim tabRentaReg(,,)() = PretraitementPrix.constructionTableauxNA(maxPrixAbsent, fenetreEstDebut, fenetreEstFin, tabRenta, tabRentaMarche)
                tabAR = modeleMarche(tailleComplete, premiereDate, fenetreEstDebut, fenetreEstFin, tabRenta, tabRentaMarche, tabRentaReg)
            Case Else
                MsgBox("Erreur interne : numero de modèle incorrect dans ChoixSeuilFenetre", 16)
                tabAR = Nothing
        End Select
        'affichage des AR dans une nouvelle feuille excel
        Globals.ThisAddIn.Application.Sheets.Add()
        Globals.ThisAddIn.Application.ActiveSheet.Name = "AR"
        Dim currentSheet As Excel.Worksheet = CType(Globals.ThisAddIn.Application.Worksheets("AR"), Excel.Worksheet)
        For i = fenetreEvDebut To fenetreEvFin
            currentSheet.Cells(i - fenetreEvDebut + 1, 1).Value = i
            For j = 0 To tabAR.GetUpperBound(1)
                currentSheet.Cells(i - fenetreEvDebut + 1, j + 2).Value = tabAR(i - fenetreEvDebut, j)
            Next
        Next
        Return tabAR
    End Function


    '***************************** Modèle de marché *****************************

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
                    a(reg) = Globals.ThisAddIn.Application.WorksheetFunction.Index(Globals.ThisAddIn.Application.WorksheetFunction.LinEst(Y, X), 2) / (reg + 1)
                    b(reg) = Globals.ThisAddIn.Application.WorksheetFunction.Index(Globals.ThisAddIn.Application.WorksheetFunction.LinEst(Y, X), 1) / (reg + 1)
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
        modeleMarche = tabAR
    End Function


    '***************************** Modèle de marché simplifié *****************************

    'Estimation des AR à partir du modèle de marché simplifié : K = moyenne des rentabilités
    Public Function modeleMarcheSimple() As Double(,)
        Dim currentSheet As Excel.Worksheet = CType(Globals.ThisAddIn.Application.Worksheets("Rt"), Excel.Worksheet)
        'compte le nombre de lignes et de colonnes
        Dim nbLignes As Integer = currentSheet.UsedRange.Rows.Count
        Dim nbColonnes As Integer = currentSheet.UsedRange.Columns.Count
        'tableau stockant les AR calculés grâce à la régression
        Dim tabAR(nbLignes - 2, nbColonnes - 2) As Double

        For i = 0 To nbColonnes - 2
            'remplissage du tableau
            For t = 0 To nbLignes - 2
                currentSheet = CType(Globals.ThisAddIn.Application.Worksheets("Rm"), Excel.Worksheet)
                Dim k As Double = currentSheet.Cells(t + 2, i + 2).Value
                currentSheet = CType(Globals.ThisAddIn.Application.Worksheets("Rt"), Excel.Worksheet)
                tabAR(t, i) = currentSheet.Cells(t + 2, i + 2).Value - k
            Next
        Next
        modeleMarcheSimple = tabAR
    End Function


    '***************************** Modèle des rentabilités moyennes *****************************

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

        For i = 1 To tailleFenetreEv
            Dim somme As Double = 0
            For e = 1 To N
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
                tab(e - 1) = tabCAR(i - 1, e - 1)
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
