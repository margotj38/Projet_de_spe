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
        ReDim tabPrixCentres(minDown + minUp, tabDate.GetUpperBound(0) + 1)
        ReDim tabMarcheCentre(minDown + minUp, tabDate.GetUpperBound(0) + 1)

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

        For colonne = 1 To tabDate.GetLength(0)
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


    

End Module
