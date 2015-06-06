﻿Module UtilitaireRentabilites

    'Variables globales
    'Le tableau des rentabilités est nécessaire car on ne les affiche pas et on ne peut donc pas les récupérer
    'lorsque l'utilisateur fait sa sélection des rentabilités des entreprises
    Public tabRentaMarche(,) As Double = Nothing
    'De même pour les entreprises
    Public tabRenta(,) As Double = Nothing
    'Le nombre maximum de #N/A présents à la suite sont nécessaires pouvoir construire les tableaux de rentabilités
    'spécifiques à la régression
    Public maxPrixAbs As Integer = 0

    '***************************** Construction des rentabilités *****************************

    'Entrée : plages des rentabilités pour les entreprises pour les période d'événement et d'estimation, 
    'tableaux des rentabilités du marché et des entreprises
    'Sortie : tableaux des rentabilités des entreprises et du marché pour les période d'événement et d'estimation
    Public Sub constructionTabRenta(plageEst As String, plageEv As String, _
                                    ByRef tabRentaMarche(,) As Double, ByRef tabRenta(,) As Double, _
                                    ByRef tabRentaMarcheEst(,) As Double, ByRef tabRentaMarcheEv(,) As Double, _
                                    ByRef tabRentaEst(,) As Double, ByRef tabRentaEv(,) As Double)

        'On parse les plages pour récupérer les indices de la fenêtre
        Dim premiereCol As Integer, derniereCol As Integer
        Dim debutEst As Integer, finEst As Integer, debutEv As Integer, finEv As Integer
        Utilitaires.parserPlageColonnes(plageEst, premiereCol, derniereCol)
        parserPlageLignes(plageEst, debutEst, finEst)
        parserPlageLignes(plageEv, debutEv, finEv)

        'On met les tableaux à la bonne dimension
        ReDim tabRentaEst(finEst - debutEst, derniereCol - premiereCol)
        ReDim tabRentaEv(finEv - debutEv, derniereCol - premiereCol)
        ReDim tabRentaMarcheEst(finEst - debutEst, derniereCol - premiereCol)
        ReDim tabRentaMarcheEv(finEv - debutEv, derniereCol - premiereCol)

        'Pour chaque colonne
        For colonne = premiereCol To derniereCol
            'On remplit le tableau d'estimation
            For i = debutEst To finEst
                tabRentaEst(i - debutEst, colonne - premiereCol) = tabRenta(i - debutEst, colonne - premiereCol)
                tabRentaMarcheEst(i - debutEst, colonne - premiereCol) = tabRentaMarche(i - debutEst, colonne - premiereCol)
            Next i
            'Et celui d'événement
            For i = debutEv To finEv
                tabRentaEv(i - debutEv, colonne - premiereCol) = tabRenta(i - debutEst, colonne - premiereCol)
                tabRentaMarcheEv(i - debutEv, colonne - premiereCol) = tabRentaMarche(i - debutEst, colonne - premiereCol)
            Next i
        Next colonne
    End Sub

    'Entrée : tableaux centrés des cours et du marché (1ère colonne : dates)
    'Sortie : tableaux des rentabilités des entreprises et du marché (1ère colonne : dates) + maxPrixAbsent
    Public Sub calculTabRenta(ByRef tabPrixCentres(,) As Double, ByRef tabMarcheCentre(,) As Double, _
                              ByRef tabRenta(,) As Double, ByRef tabRentaMarche(,) As Double, ByRef maxPrixAbsent As Integer)

        'On recopie la colonne des dates dans les tableaux
        For indDate = 1 To tabPrixCentres.GetUpperBound(0)
            tabRenta(indDate - 1, 0) = tabPrixCentres(indDate, 0)
            tabRentaMarche(indDate - 1, 0) = tabMarcheCentre(indDate, 0)
        Next indDate

        'On calcule les rentabilités et les rentabilités de marché associées
        Dim prixPresent As Integer = 0
        'Pour savoir combien de tableaux stockant les Rt et Rm on va déclaré
        maxPrixAbsent = 0

        For titre = 1 To tabPrixCentres.GetUpperBound(1)
            For indDate = 0 To tabPrixCentres.GetUpperBound(0)
                If prixPresent = 0 Then
                    'Si on est sur le premier prix
                    '(-2146826246 est la valeur obtenue lorsqu'un ".Value" est fait sur une cellule #N/A)
                    If Not (tabPrixCentres(indDate, titre) = -2146826246) Then
                        prixPresent = prixPresent + 1
                        If prixPresent > maxPrixAbsent Then
                            maxPrixAbsent = prixPresent
                        End If
                    End If
                ElseIf tabPrixCentres(indDate, titre) = -2146826246 Then
                    'Si il n'y a pas de prix à cette date
                    'On met un équivalent de #N/A dans les tableaux
                    tabRenta(indDate - 1, titre) = -2146826246
                    tabRentaMarche(indDate - 1, titre) = -2146826246
                    prixPresent = prixPresent + 1
                    If prixPresent > maxPrixAbsent Then
                        maxPrixAbsent = prixPresent
                    End If
                Else
                    'Sinon on fait le calcul en remontant au dernier prix disponible
                    tabRenta(indDate - 1, titre) = ((tabPrixCentres(indDate, titre) - tabPrixCentres(indDate - prixPresent, titre)) / _
                        tabPrixCentres(indDate - prixPresent, titre)) / prixPresent
                    'On fait de même pour les rentabilités de marché
                    tabRentaMarche(indDate - 1, titre) = ((tabMarcheCentre(indDate, titre) - tabMarcheCentre(indDate - prixPresent, titre)) / _
                        tabMarcheCentre(indDate - prixPresent, titre)) / prixPresent
                    'Et on indique qu'un prix était présent
                    prixPresent = 1
                End If
            Next indDate
            prixPresent = 0
        Next titre
    End Sub

    Public Function constructionTableauxReg(maxPrixAbsent As Integer, ByRef tabRentaEst(,) As Double, _
                                           ByRef tabRentaMarcheEst(,) As Double) As Double(,,)()

        'Déclaration du tableau à retourner
        Dim tabRentaReg(tabRentaEst.GetUpperBound(1), maxPrixAbsent - 1, 1)() As Double
        For i = 0 To tabRentaEst.GetUpperBound(1)
            For j = 0 To maxPrixAbsent - 1
                For k = 0 To 1
                    tabRentaReg(i, j, k) = New Double(tabRentaEst.GetUpperBound(0)) {}
                Next
            Next
        Next

        Dim prixPresent As Integer = 1
        For titre = 0 To tabRentaEst.GetUpperBound(1)
            'Tableau permettant de savoir si un redimensionnement est nécessaire
            Dim tabRedimEst(maxPrixAbsent - 1) As Integer
            For indDate = 0 To tabRentaEst.GetUpperBound(0)
                If tabRentaEst(indDate, titre) = -2146826246 Then
                    'Si il n'y a pas de prix à cette date
                    prixPresent = prixPresent + 1
                Else
                    'Sinon, on range les rentabilités dans le tableau

                    'On ajoute Rt et Rm au tableau
                    'Les rentabilités sont ramenées en équivalent à une période (par division par prixPresent)
                    tabRentaReg(titre, prixPresent - 1, 0)(tabRedimEst(prixPresent - 1)) = tabRentaEst(indDate, titre)
                    tabRentaReg(titre, prixPresent - 1, 1)(tabRedimEst(prixPresent - 1)) = tabRentaMarcheEst(indDate, titre)

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


    '***************************** Pour centrer les prix/rentabilités autour des dates d'événement *****************************

    Sub donneesCentrees(plageDate As String, feuilleDates As String, feuilleDonnees As String, ByRef tabEntreprisesCentre(,) As Double, ByRef tabMarcheCentre(,) As Double)

        Dim currentSheet As Excel.Worksheet = CType(Globals.ThisAddIn.Application.Worksheets(feuilleDates), Excel.Worksheet)

        Dim datesEv(currentSheet.Range(plageDate).Rows.Count - 1) As Date
        For i = 1 To currentSheet.Range(plageDate).Rows.Count
            datesEv(i - 1) = currentSheet.Range(plageDate).Cells(i, 1).Value
        Next i

        'Deux tableaux 1 dimension à trier selon le premier tableau
        Dim tabDate(datesEv.GetLength(0) - 1) As Date
        Dim tabInd(datesEv.GetLength(0) - 1) As Integer
        For i = 0 To datesEv.GetLength(0) - 1
            tabDate(i) = datesEv(i)
            tabInd(i) = i + 1
        Next

        'Tri des tableaux selon les dates
        TriDoubleTab(tabDate, tabInd, tabDate.GetLowerBound(0), tabDate.GetUpperBound(0))

        'on se positionne sur la feuille des prix
        currentSheet = CType(Globals.ThisAddIn.Application.Worksheets(feuilleDonnees), Excel.Worksheet)
        Dim nbLignes As Integer = currentSheet.UsedRange.Rows.Count
        Dim nbColonnes As Integer = currentSheet.UsedRange.Columns.Count

        'calul taille fenetre globale
        Dim minUp As Integer, minDown As Integer
        'indice premiere date evenement - indice premiere date
        minUp = currentSheet.Range("A:A").Find(Format(tabDate(0), "Short date").ToString).Row - 2
        'indice derniere date - derniere date evenement
        minDown = nbLignes - currentSheet.Columns("A:A").Find(Format(tabDate(tabDate.GetUpperBound(0)), "Short date").ToString).Row

        'Redimensionnement des tableaux de retour
        ReDim tabEntreprisesCentre(minDown + minUp, tabDate.GetUpperBound(0) + 1)
        ReDim tabMarcheCentre(minDown + minUp, tabDate.GetUpperBound(0) + 1)

        For i = -minUp To minDown
            tabEntreprisesCentre(i + minUp, 0) = i
            tabMarcheCentre(i + minUp, 0) = i
        Next

        For colonne = 1 To tabDate.GetLength(0)
            'on se positionne sur la feuille contenant les prix
            currentSheet = CType(Globals.ThisAddIn.Application.Worksheets(feuilleDonnees), Excel.Worksheet)
            Dim fenetreInf As Integer, fenetreSup As Integer
            Dim dateCour As Excel.Range
            Dim data As Excel.Range, marche As Excel.Range
            dateCour = currentSheet.Columns("A:A").Find(Format(tabDate(colonne - 1), "Short date").ToString)
            fenetreInf = dateCour.Row - minUp
            fenetreSup = dateCour.Row + minDown
            'récupération des prix centrés autour de l'évènement
            data = currentSheet.Range(currentSheet.Cells(fenetreInf, tabInd(colonne - 1) + 2), currentSheet.Cells(fenetreSup, tabInd(colonne - 1) + 2))
            'récupération des indices de marché correspondants
            marche = currentSheet.Range(currentSheet.Cells(fenetreInf, 2), currentSheet.Cells(fenetreSup, 2))

            For i = -minUp To minDown
                tabEntreprisesCentre(i + minUp, tabInd(colonne - 1)) = data.Cells(i + minUp + 1, 1).Value
                tabMarcheCentre(i + minUp, tabInd(colonne - 1)) = marche.Cells(i + minUp + 1, 1).Value
            Next i
        Next
    End Sub

End Module
