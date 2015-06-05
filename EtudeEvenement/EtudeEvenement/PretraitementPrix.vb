Module PretraitementPrix

    '***************************** Pour centrer les prix *****************************

    'Génère le tableau des prix centrés autour de la date d'évènement
    '2 nouveaux onglets sont créés, un pour les prix, un pour le marché
    Sub prixCentres(plageDate As String, feuille As String, ByRef tabPrixCentres(,) As Double, ByRef tabMarcheCentre(,) As Double)

        Dim currentSheet As Excel.Worksheet = CType(Globals.ThisAddIn.Application.Worksheets(feuille), Excel.Worksheet)

        'tableau des dates d'évènements
        'Dim datesEv()
        'datesEv = currentSheet.Range(plageDate).Value
        Dim datesEv(currentSheet.Range(plageDate).Rows.Count - 1) As Date
        For i = 1 To currentSheet.Range(plageDate).Rows.Count
            datesEv(i - 1) = currentSheet.Range(plageDate).Cells(i, 1).Value
        Next i


        'tableau 2 dimensions à trier
        'Dim aTrier(datesEv.GetLength(0) - 1, 1) As Object
        'For i = 0 To datesEv.GetLength(0) - 1
        '    aTrier(i, 0) = i
        '    aTrier(i, 1) = datesEv(i)
        'Next

        'Deux tableaux 1 dimension à trier selon le premier tableau
        Dim tabDate(datesEv.GetLength(0) - 1) As Date
        Dim tabInd(datesEv.GetLength(0) - 1) As Integer
        For i = 0 To datesEv.GetLength(0) - 1
            tabDate(i) = datesEv(i)
            tabInd(i) = i + 1
        Next

        'Tri des tableaux selon les dates
        TriDoubleTab(tabDate, tabInd, tabDate.GetLowerBound(0), tabDate.GetUpperBound(0))

        'Tri du tableau selon les dates
        'Tri(aTrier, 1, LBound(aTrier, 1), UBound(aTrier, 1))

        'on se positionne sur la feuille des prix
        currentSheet = CType(Globals.ThisAddIn.Application.Worksheets("Prix"), Excel.Worksheet)
        Dim nbLignes As Integer = currentSheet.UsedRange.Rows.Count
        Dim nbColonnes As Integer = currentSheet.UsedRange.Columns.Count

        'calul taille fenetre globale
        Dim minUp As Integer, minDown As Integer
        'indice premiere date evenement - indice premiere date
        minUp = currentSheet.Range("A:A").Find(Format(tabDate(0), "Short date").ToString).Row - 2
        'indice derniere date - derniere date evenement
        minDown = nbLignes - currentSheet.Columns("A:A").Find(Format(tabDate(tabDate.GetUpperBound(0)), "Short date").ToString).Row

        ''écritures des entêtes de lignes et colonnes sur la nouvelle feuille prixCentres
        'currentSheet = CType(Globals.ThisAddIn.Application.Worksheets("prixCentres"), Excel.Worksheet)
        'currentSheet.Cells(1, 1).Value = "Date"
        'For i = 2 To nbColonnes - 1
        '    currentSheet.Cells(1, i).Value = "P" & i - 1
        'Next

        'Redimensionnement des tableaux de retour
        ReDim tabPrixCentres(minDown + minUp, nbColonnes - 2)
        ReDim tabMarcheCentre(minDown + minUp, nbColonnes - 2)

        For i = -minUp To minDown
            'currentSheet.Cells(i + minUp + 2, 1).Value = i
            tabPrixCentres(i + minUp, 0) = i
            tabMarcheCentre(i + minUp, 0) = i
        Next
        ''de même pour marcheCentre
        'currentSheet = CType(Globals.ThisAddIn.Application.Worksheets("marcheCentre"), Excel.Worksheet)
        'currentSheet.Cells(1, 1).Value = "Date"
        'For i = 2 To nbColonnes - 1
        '    currentSheet.Cells(1, i).Value = "Pm pour P" & i - 1
        'Next
        'For i = -minUp To minDown
        '    currentSheet.Cells(i + minUp + 2, 1).Value = i
        'Next

        For colonne = 1 To nbColonnes - 2
            'on se positionne sur la feuille contenant les prix
            currentSheet = CType(Globals.ThisAddIn.Application.Worksheets("Prix"), Excel.Worksheet)
            Dim fenetreInf As Integer, fenetreSup As Integer
            Dim dateCour As Excel.Range, firmeCour As Excel.Range
            Dim data As Excel.Range, marche As Excel.Range
            dateCour = currentSheet.Columns("A:A").Find(Format(tabDate(colonne - 1), "Short date").ToString)
            fenetreInf = dateCour.Row - minUp
            fenetreSup = dateCour.Row + minDown
            'firmeCour = currentSheet.Rows("1:1").Find(aTrier(colonne - 1, 0).ToString)
            firmeCour = currentSheet.Rows("1:1").Find("P" & tabInd(colonne - 1).ToString)
            'récupération des prix centrés autour de l'évènement
            data = currentSheet.Range(currentSheet.Cells(fenetreInf, firmeCour.Column), currentSheet.Cells(fenetreSup, firmeCour.Column))
            'récupération des indices de marché correspondants
            marche = currentSheet.Range(currentSheet.Cells(fenetreInf, 2), currentSheet.Cells(fenetreSup, 2))

            For i = -minUp To minDown
                tabPrixCentres(i + minUp, tabInd(colonne - 1)) = data.Cells(i + minUp + 1, 1).Value
                tabMarcheCentre(i + minUp, tabInd(colonne - 1)) = marche.Cells(i + minUp + 1, 1).Value
            Next i

            'on se positionne sur la feuille contenant les prix centrés pour écrire dedans
            'currentSheet = CType(Globals.ThisAddIn.Application.Worksheets("prixCentres"), Excel.Worksheet)
            'currentSheet.Range(currentSheet.Cells(2, i + 1), currentSheet.Cells(minUp + minDown + 2, i + 1)).Value = data.Value
            'on se positionne sur la feuille contenant les indices de marché pour écrire dedans
            'currentSheet = CType(Globals.ThisAddIn.Application.Worksheets("marcheCentre"), Excel.Worksheet)
            'currentSheet.Range(currentSheet.Cells(2, i + 1), currentSheet.Cells(minUp + minDown + 2, i + 1)).Value = marche.Value
        Next
    End Sub

    Sub TriDoubleTab(tabDate() As Date, tabInd() As Integer, gauche As Integer, droite As Integer) ' Quick sort
        Dim ref As Date = tabDate((gauche + droite) \ 2)
        Dim g As Integer = gauche
        Dim d As Integer = droite
        Do
            Do While tabDate(g) < ref : g = g + 1 : Loop
            Do While ref < tabDate(d) : d = d - 1 : Loop
            If g <= d Then
                Dim tempDate As Date = tabDate(g)
                tabDate(g) = tabDate(d)
                tabDate(d) = tempDate
                Dim temp As String = tabInd(g)
                tabInd(g) = tabInd(d)
                tabInd(d) = temp
                g = g + 1
                d = d - 1
            End If
        Loop While g <= d
        If g < droite Then TriDoubleTab(tabDate, tabInd, g, droite)
        If gauche < d Then TriDoubleTab(tabDate, tabInd, gauche, d)
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


    '***************************** Transformation en rentabilités *****************************

    Public Function constructionTableauxNA(maxPrixAbsent As Integer, fenetreEstDebut As Integer, fenetreEstFin As Integer, _
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

        Dim currentSheet As Excel.Worksheet = CType(Globals.ThisAddIn.Application.Worksheets("prixCentres"), Excel.Worksheet)
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

    'Entrée : tableaux centrés des cours et du marché (1ère colonne : dates)
    'Sortie : tableaux des rentabilités des entreprises et du marché (1ère colonne : dates)
    Public Sub calculTabRenta(ByRef tabPrixCentres(,) As Double, ByRef tabMarcheCentre(,) As Double, _
                              ByRef tabRenta(,) As Double, ByRef tabRentaMarche(,) As Double)

        'On recopie la colonne des dates dans les tableaux
        For indDate = 1 To tabPrixCentres.GetUpperBound(0)
            tabRenta(indDate - 1, 0) = tabPrixCentres(indDate, 0)
            tabRentaMarche(indDate - 1, 0) = tabMarcheCentre(indDate, 0)
        Next indDate

        'On calcule les rentabilités et les rentabilités de marché associées
        Dim prixPresent As Integer = 0
        'Pour savoir combien de tableaux stockant les Rt et Rm on va déclaré
        Dim maxPrixAbsent As Integer = 0

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
                    tabRenta(indDate - 1, titre) = (tabPrixCentres(indDate, titre) - tabPrixCentres(indDate - prixPresent, titre)) / _
                        tabPrixCentres(indDate - prixPresent, titre)
                    'On fait de même pour les rentabilités de marché
                    tabRentaMarche(indDate - 1, titre) = (tabMarcheCentre(indDate, titre) - tabMarcheCentre(indDate - prixPresent, titre)) / _
                        tabMarcheCentre(indDate - prixPresent, titre)
                    'Et on indique qu'un prix était présent
                    prixPresent = 1
                End If
            Next indDate
            prixPresent = 0
        Next titre
    End Sub

    'Entrée : plages des rentabilités pour les entreprises pour les période d'événement et d'estimation, 
    'tableaux des rentabilités du marché
    'Sortie : tableaux des rentabilités du marché pour les période d'événement et d'estimation
    Public Sub constructionTabRenta(plageEst As String, plageEv As String, ByRef tabRentaMarche(,) As Double, _
                                    ByRef tabRentaMarcheEst(,) As Double, tabRentaMarcheEv(,) As Double)
        'On parse les plages pour récupérer les indices de la fenêtre
        Dim premiereCol As Integer, derniereCol As Integer
        Dim debutEst As Integer, finEst As Integer, debutEv As Integer, finEv As Integer
        parserPlageColonnes(plageEst, premiereCol, derniereCol)
        parserPlageLignes(plageEst, debutEst, finEst)
        parserPlageLignes(plageEv, debutEv, finEv)

        'Pour chaque colonne
        For colonne = premiereCol - 2 To derniereCol - 2
            'On remplit le tableau d'estimation
            For i = debutEst To finEst
                tabRentaMarcheEst(i - debutEst, colonne - 2) = tabRentaMarche(i - 2, colonne - 2)
            Next i
            'Et celui d'événement
            For i = debutEv To finEv
                tabRentaMarcheEv(i - debutEv, colonne - 2) = tabRentaMarche(i - 2, colonne - 2)
            Next i
        Next colonne
    End Sub

    Private Sub parserPlageColonnes(plage As String, ByRef premiereCol As Integer, ByRef derniereCol As Integer)
        Dim rangePlage As Excel.Range = Globals.ThisAddIn.Application.Range(plage)
        premiereCol = rangePlage.Cells(1, 1).Column()
        derniereCol = rangePlage.Cells(1, rangePlage.Columns.Count).Column()
    End Sub

    Private Sub parserPlageLignes(plage As String, ByRef debut As Integer, ByRef fin As Integer)
        Dim rangePlage As Excel.Range = Globals.ThisAddIn.Application.Range(plage)
        debut = rangePlage.Cells(1, 1).Row()
        fin = rangePlage.Cells(rangePlage.Rows.Count, 1).Row()
    End Sub

    Public Sub recupererFeuillePlage(textRefEdit As String, ByRef feuille As String, ByRef plage As String)
        Dim tabString() As String = Split(textRefEdit, "'")
        feuille = tabString(1)
        'On enlève le '!' 
        tabString = Split(tabString(2), "!")
        plage = tabString(1)
    End Sub

End Module
