﻿
''' <summary>
''' Module de gestion des opérations sur les prix et les rentabilités.
''' </summary>
''' <remarks></remarks>
Module UtilitaireRentabilites

    'Variables globales

    ''' <summary>
    ''' Rentabilités de marché.
    ''' </summary>
    ''' <remarks> Le tableau des rentabilités est nécessaire car on ne les affiche pas et on ne peut donc pas les récupérer 
    ''' lorsque l'utilisateur fait sa sélection des rentabilités des entreprises</remarks>
    Public tabRentaMarche(,) As Double = Nothing

    ''' <summary>
    ''' Rentabilités des entreprises.
    ''' </summary>
    ''' <remarks> On conserve le tableau des rentabilités des entreprises après l'affichage pour ne pas avoir à récupérer 
    ''' de nouveau les données de l'affichage par la suite. </remarks>
    Public tabRenta(,) As Double = Nothing

    ''' <summary>
    ''' Le nombre maximum de #N/A présents à la suite dans les rentabilités.
    ''' </summary>
    ''' <remarks> Ce nombre est nécessaire pour construire le tableau des rentabilité spécifique au modèle de marché. </remarks>
    Public maxRentAbs As Integer = 0

    ''' <summary>
    ''' Tableau des rentabilités de marché calculé simplement sur chaque date.
    ''' </summary>
    ''' <remarks> Utilisé pour le test de Patell. </remarks>
    Public tabRentaClassiquesMarche(,) As Double = Nothing


    '***************************** Construction des rentabilités *****************************

    ''' <summary>
    ''' Méthode qui construit les six tableaux de rentabilités : rentabilités des entreprises, rentabilités de marché 
    ''' calculées de la même façon que les rentabilités d'entreprises, rentabilités de marché calculées normalement 
    ''' (sans tenir compte des #N/A des rentabilités d'entreprise). A chacune de ces rentabilités est associé un tableau 
    ''' pour la période d'estimation et un autre pour la période d'événement.
    ''' </summary>
    ''' <param name="plageEst">Plage contenant les données de la période d'estimation.</param>
    ''' <param name="plageEv">Plage contenant les données de la période d'événement.</param>
    ''' <param name="tabRentaMarche">Tableau des rentabilités de marché (la première colonne étant les dates).</param>
    ''' <param name="tabRenta">Tableau des rentabilités des entreprises calculées de la même façon que les rentabilités 
    ''' des entreprises (la première colonne étant les dates).</param>
    ''' <param name="tabRentaClassiquesMarche">Tableau des rentabilité de marché calculées classiquement 
    ''' (la première colonne étant les dates).</param>
    ''' <param name="tabRentaMarcheEst">(Sortie) Tableau des rentabilités de marché sur la période d'estimation, 
    ''' calculées de la même façon que les rentabilités des entreprises (la première colonne étant les dates).</param>
    ''' <param name="tabRentaMarcheEv">(Sortie) Tableau des rentabilités de marché sur la période d'événement, 
    ''' calculées de la même façon que les rentabilités des entreprises (la première colonne étant les dates).</param>
    ''' <param name="tabRentaEst">(Sortie) Tableau des rentabilités des entreprises sur la période d'estimation 
    ''' (la première colonne étant les dates).</param>
    ''' <param name="tabRentaEv">(Sortie) Tableau des rentabilités des entreprises sur la période d'événement 
    ''' (la première colonne étant les dates).</param>
    ''' <param name="tabRentaClassiquesMarcheEst">(Sortie) Tableau des rentabilités de marché sur la période d'estimation, 
    ''' calculées classiquement (la première colonne étant les dates).</param>
    ''' <param name="tabRentaClassiquesMarcheEv">(Sortie) Tableau des rentabilités de marché sur la période d'événement, 
    ''' calculées classiquement (la première colonne étant les dates).</param>
    ''' <remarks></remarks>
    Public Sub constructionTabRenta(plageEst As String, plageEv As String, _
                                    ByRef tabRentaMarche(,) As Double, ByRef tabRenta(,) As Double, _
                                    ByRef tabRentaClassiquesMarche(,) As Double, _
                                    ByRef tabRentaMarcheEst(,) As Double, ByRef tabRentaMarcheEv(,) As Double, _
                                    ByRef tabRentaEst(,) As Double, ByRef tabRentaEv(,) As Double, _
                                    ByRef tabRentaClassiquesMarcheEst(,) As Double, ByRef tabRentaClassiquesMarcheEv(,) As Double)

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
        ReDim tabRentaClassiquesMarcheEst(finEst - debutEst, derniereCol - premiereCol)
        ReDim tabRentaClassiquesMarcheEv(finEv - debutEv, derniereCol - premiereCol)

        'Pour chaque colonne
        For colonne = premiereCol To derniereCol
            'On remplit le tableau d'estimation
            For i = debutEst To finEst
                tabRentaEst(i - debutEst, colonne - premiereCol) = tabRenta(i - 2, colonne - 1)
                tabRentaMarcheEst(i - debutEst, colonne - premiereCol) = tabRentaMarche(i - 2, colonne - 1)
                tabRentaClassiquesMarcheEst(i - debutEst, colonne - premiereCol) = tabRentaClassiquesMarche(i - 2, colonne - 1)
            Next i
            'Et celui d'événement
            For i = debutEv To finEv
                tabRentaEv(i - debutEv, colonne - premiereCol) = tabRenta(i - 2, colonne - 1)
                tabRentaMarcheEv(i - debutEv, colonne - premiereCol) = tabRentaMarche(i - 2, colonne - 1)
                tabRentaClassiquesMarcheEv(i - debutEv, colonne - premiereCol) = tabRentaClassiquesMarche(i - 2, colonne - 1)
            Next i
        Next colonne
    End Sub

    ''' <summary>
    ''' Méthode calculant les rentabilités des entreprises et du marché à partir des cours.
    ''' </summary>
    ''' <param name="tabPrixCentres">Tableau des cours des entreprises, centrés autour de l'événement.</param>
    ''' <param name="tabMarcheCentre">Tableau des prix du marché, centrés autour de l'événement.</param>
    ''' <param name="tabRenta">(Sortie) Tableau des rentabilités des entreprises.</param>
    ''' <param name="tabRentaMarche">(Sortie) Tableau des rentabilités de marché calculées de la même façon que les 
    ''' rentabilités des entreprises.</param>
    ''' <param name="tabRentaClassiquesMarche">(Sortie) Tableau des rentabilités de marché calculées classiquement.</param>
    ''' <param name="maxPrixAbsent">(Sortie) Nombre maximal de données manquantes consécutives + 1.</param>
    ''' <param name="rentaLog">Valeur déterminant le type de calcul des rentabilités (False : calcul arithmétique, 
    ''' True : calcul logarithmique).</param>
    ''' <remarks></remarks>
    Public Sub calculTabRenta(ByRef tabPrixCentres(,) As Double, ByRef tabMarcheCentre(,) As Double, _
                              ByRef tabRenta(,) As Double, ByRef tabRentaMarche(,) As Double, _
                              ByRef tabRentaClassiquesMarche(,) As Double, ByRef maxPrixAbsent As Integer, _
                              rentaLog As Boolean)

        'On recopie la colonne des dates dans les tableaux
        For indDate = 1 To tabPrixCentres.GetUpperBound(0)
            tabRenta(indDate - 1, 0) = tabPrixCentres(indDate, 0)
            tabRentaMarche(indDate - 1, 0) = tabMarcheCentre(indDate, 0)
            tabRentaClassiquesMarche(indDate - 1, 0) = tabMarcheCentre(indDate, 0)
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
                    If Not (Double.IsNaN(tabPrixCentres(indDate, titre))) Then
                        prixPresent = prixPresent + 1
                        If prixPresent > maxPrixAbsent Then
                            maxPrixAbsent = prixPresent
                        End If
                    End If
                ElseIf Double.IsNaN(tabPrixCentres(indDate, titre)) Then
                    'Si il n'y a pas de prix à cette date
                    'On met un équivalent de #N/A dans les tableaux
                    tabRenta(indDate - 1, titre) = Double.NaN
                    tabRentaMarche(indDate - 1, titre) = Double.NaN
                    'Dans tabRentaClassiquesMarche, on fait le calcul classique selon le mode
                    If rentaLog Then
                        tabRentaClassiquesMarche(indDate - 1, titre) = Math.Log(tabMarcheCentre(indDate, titre) / _
                                                                 tabMarcheCentre(indDate - 1, titre))
                    Else
                        tabRentaClassiquesMarche(indDate - 1, titre) = (tabMarcheCentre(indDate, titre) - tabMarcheCentre(indDate - 1, titre)) / _
                                                                 tabMarcheCentre(indDate - 1, titre)
                    End If

                    prixPresent = prixPresent + 1
                    If prixPresent > maxPrixAbsent Then
                        maxPrixAbsent = prixPresent
                    End If
                Else
                    'Sinon on fait le calcul en remontant au dernier prix disponible (avec le bon mode de calcul (log ou arithmétique))
                    If rentaLog Then
                        'Calcul des rentabilités des entreprises
                        tabRenta(indDate - 1, titre) = (Math.Log(tabPrixCentres(indDate, titre) / _
                                                                 tabPrixCentres(indDate - prixPresent, titre))) / prixPresent
                        'On fait de même pour les rentabilités de marché
                        tabRentaMarche(indDate - 1, titre) = (Math.Log(tabMarcheCentre(indDate, titre) / _
                                                                 tabMarcheCentre(indDate - prixPresent, titre))) / prixPresent
                        'Dans tabRentaClassiquesMarche, on fait le calcul classique
                        tabRentaClassiquesMarche(indDate - 1, titre) = Math.Log(tabMarcheCentre(indDate, titre) / _
                                                                 tabMarcheCentre(indDate - 1, titre))
                    Else
                        'Calcul des rentabilités des entreprises
                        tabRenta(indDate - 1, titre) = ((tabPrixCentres(indDate, titre) - tabPrixCentres(indDate - prixPresent, titre)) / _
                            tabPrixCentres(indDate - prixPresent, titre)) / prixPresent
                        'On fait de même pour les rentabilités de marché
                        tabRentaMarche(indDate - 1, titre) = ((tabMarcheCentre(indDate, titre) - tabMarcheCentre(indDate - prixPresent, titre)) / _
                            tabMarcheCentre(indDate - prixPresent, titre)) / prixPresent
                        'Dans tabRentaClassiquesMarche, on fait le calcul classique
                        tabRentaClassiquesMarche(indDate - 1, titre) = (tabMarcheCentre(indDate, titre) - tabMarcheCentre(indDate - 1, titre)) / _
                            tabMarcheCentre(indDate - 1, titre)
                    End If

                    'Et on indique qu'un prix était présent
                    prixPresent = 1
                End If
            Next indDate
            prixPresent = 0
        Next titre
    End Sub

    ''' <summary>
    ''' Fonction qui construit le tableau des rentabilités spécifique au modèle de marché.
    ''' </summary>
    ''' <param name="maxRentAbsent"> Nombre maximum de données manquantes consécutives dans les rentabilités. </param>
    ''' <param name="tabRentaEst"> Rentabilités sur la période d'estimation. </param>
    ''' <param name="tabRentaMarcheEst"> Rentabilités de marché sur la période d'estimation. </param>
    ''' <returns> Tableau de rentabilités spécifiques au modèle de marché. Détail de chacune des dimensions : 
    ''' la 1ère dimension correspond aux entreprises, la 2ème dimension correspond au nombre de données manquantes
    ''' consécutives précédent une certaine rentabilité, la 3ème dimension correspond au type de rentabilité (indice 0 pour 
    ''' les entreprises et 1 pour le marché). Chaque élément de ce tableau est un tableau de rentabilités. </returns>
    ''' <remarks> Exemple d'accès au données : tabDeRetour(3, 2, 1)() est la tableau des rentabilités de marché précédées par deux données 
    ''' manquantes pour la quatrième entreprise. </remarks>
    Public Function constructionTableauxReg(maxRentAbsent As Integer, ByRef tabRentaEst(,) As Double, _
                                           ByRef tabRentaMarcheEst(,) As Double) As Double(,,)()

        'Déclaration du tableau à retourner
        Dim tabRentaReg(tabRentaEst.GetUpperBound(1), maxRentAbsent - 1, 1)() As Double
        For i = 0 To tabRentaEst.GetUpperBound(1)
            For j = 0 To maxRentAbsent - 1
                For k = 0 To 1
                    tabRentaReg(i, j, k) = New Double(tabRentaEst.GetUpperBound(0)) {}
                Next
            Next
        Next

        Dim prixPresent As Integer = 1
        For titre = 0 To tabRentaEst.GetUpperBound(1)
            'Tableau permettant de savoir si un redimensionnement est nécessaire
            Dim tabRedimEst(maxRentAbsent - 1) As Integer
            For indDate = 0 To tabRentaEst.GetUpperBound(0)
                If Double.IsNaN(tabRentaEst(indDate, titre)) Then
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
            For prixPres = 0 To maxRentAbsent - 1
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

    ''' <summary>
    ''' Fonction qui calcule le nombre de #NA consécutifs maximal présent dans les rentabilités.
    ''' </summary>
    ''' <param name="tabRenta"> Tableau des rentabilités à analyser. </param>
    ''' <returns> Nombre de #NA consécutifs maximal présent dans les rentabilités. </returns>
    ''' <remarks></remarks>
    Public Function calculMaxRentAbs(tabRenta(,) As Double) As Integer
        Dim prixPresent As Integer = 1
        Dim maxPrixAbsent As Integer = 1

        For titre = 1 To tabRenta.GetUpperBound(1)
            For indDate = 0 To tabRenta.GetUpperBound(0)
                If tabRenta(indDate, titre) = -2146826246 Then
                    prixPresent = prixPresent + 1
                    If prixPresent > maxPrixAbsent Then
                        maxPrixAbsent = prixPresent
                    End If
                Else
                    prixPresent = 1
                End If
            Next indDate
            prixPresent = 1
        Next titre

        Return maxPrixAbsent
    End Function


    '***************************** Pour centrer les prix/rentabilités autour des dates d'événement *****************************

    ''' <summary>
    ''' Centre des données de prix ou de rentabilités autour de la date d'événement pour chaque entreprise.
    ''' </summary>
    ''' <param name="plageDate"> Plage des données Excel des dates d'événement pour chaque entreprise. </param>
    ''' <param name="feuilleDates"> Nom de la feuille contenant les dates. </param>
    ''' <param name="feuilleDonnees"> Nom de la feuille contenant les données de prix ou de rentabilités. </param>
    ''' <param name="tabEntreprisesCentre"> (Sortie) Données centrées pour les entreprises. </param>
    ''' <param name="tabMarcheCentre"> (Sortie) Données centrées correspondantes (pour chaque entreprise) au niveau du marché. </param>
    ''' <param name="coursOuv"> Booléen indiquant s'il s'agit de cours d'ouverture. Attention : ne peut valoir 
    ''' true que si les données à centrer sont des prix ! </param>
    ''' <remarks></remarks>
    Sub donneesCentrees(plageDate As String, feuilleDates As String, feuilleDonnees As String, ByRef tabEntreprisesCentre(,) As Double, _
                        ByRef tabMarcheCentre(,) As Double, Optional coursOuv As Boolean = False)

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
        'si ce sont des cours d'ouverture, on modifie le centrage
        If coursOuv Then
            minUp = minUp + 1
            minDown = minDown - 1
        End If

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
            If coursOuv Then
                'si ce sont des cours d'ouverture, on centre par rapport au lendemain de la date d'événement
                fenetreInf = dateCour.Row + 1 - minUp
                fenetreSup = dateCour.Row + 1 + minDown
            Else
                'sinon on centre par rapport à la date d'événement
                fenetreInf = dateCour.Row - minUp
                fenetreSup = dateCour.Row + minDown
            End If

            'récupération des prix centrés autour de l'évènement
            data = currentSheet.Range(currentSheet.Cells(fenetreInf, tabInd(colonne - 1) + 2), currentSheet.Cells(fenetreSup, tabInd(colonne - 1) + 2))
            'récupération des indices de marché correspondants
            marche = currentSheet.Range(currentSheet.Cells(fenetreInf, 2), currentSheet.Cells(fenetreSup, 2))

            For i = -minUp To minDown
                If Globals.ThisAddIn.Application.WorksheetFunction.IsNA(data.Cells(i + minUp + 1, 1)) Then
                    tabEntreprisesCentre(i + minUp, tabInd(colonne - 1)) = Double.NaN
                    tabMarcheCentre(i + minUp, tabInd(colonne - 1)) = Double.NaN
                Else
                    tabEntreprisesCentre(i + minUp, tabInd(colonne - 1)) = data.Cells(i + minUp + 1, 1).Value
                    tabMarcheCentre(i + minUp, tabInd(colonne - 1)) = marche.Cells(i + minUp + 1, 1).Value
                End If
            Next i
        Next
    End Sub

End Module
