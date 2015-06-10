Imports System.Windows.Forms.DataVisualization.Charting
Imports System.Diagnostics

''' <summary>
''' Module d'intéraction avec l'interface Excel. Il regroupe principalement les fonctions d'affichage des résultats
''' dans les feuilles Excel.
''' </summary>
''' <remarks></remarks>

Module ExcelDialogue

    ''' <summary>
    ''' Conversion d'une plage de données dont la première colonnes contient des dates 
    ''' en un tableau de données et un tableau de dates.
    ''' </summary>
    ''' <param name="plage"> Plage des données Excel sous forme de chaine de caractères. </param>
    ''' <param name="feuille"> Nom de la feuille où sélectionner la plage de données. </param>
    ''' <param name="tabDonnees"> Tableau de données en sortie. </param>
    ''' <param name="tabDates"> Tableau de dates en sortie. Attention les dates sont sous forme entières
    ''' (dates relatives après centrage) et non de type Date. </param>
    ''' <remarks></remarks>
    Public Sub convertPlageTab(plage As String, feuille As String, ByRef tabDonnees(,) As Double, ByRef tabDates() As Integer)

        Dim currentSheet As Excel.Worksheet = CType(Globals.ThisAddIn.Application.Worksheets(feuille), Excel.Worksheet)

        Dim tmpRange As Excel.Range
        tmpRange = currentSheet.Range(plage)
        'extraction de la première colonne correspondant aux dates
        tmpRange = currentSheet.Range(plage)
        ReDim tabDates(0 To tmpRange.Rows.Count - 1)
        For ligne = 0 To tabDates.GetUpperBound(0)
            tabDates(ligne) = tmpRange.Cells(ligne + 1, 1).Value
        Next
        'tableau des données pour l'estimation
        ReDim tabDonnees(0 To tmpRange.Rows.Count - 1, 0 To tmpRange.Columns.Count - 2)
        For ligne = 0 To tabDonnees.GetUpperBound(0)
            For colonne = 0 To tabDonnees.GetUpperBound(1)
                If Globals.ThisAddIn.Application.WorksheetFunction.IsNA(tmpRange.Cells(ligne + 1, colonne + 2)) Then
                    tabDonnees(ligne, colonne) = Double.NaN
                Else
                    tabDonnees(ligne, colonne) = tmpRange.Cells(ligne + 1, colonne + 2).Value
                End If
            Next
        Next
    End Sub

    ''' <summary>
    ''' Affichage des résultats des tests statistiques asymptotiques sur les AR et les CAR 
    ''' sur la période autour de l'événement.
    ''' </summary>
    ''' <param name="datesEvAR">  Dates de la période d'événement sur laquelle les tests sont réalisés. </param>
    ''' <param name="tabMoyAR"> Moyennes des AR en chaque temps de la fenêtre d'événement. </param>
    ''' <param name="tabEcartAR"> Ecart-types des AR en chaque temps de la fenêtre d'événement. </param>
    ''' <param name="statAR"> Statistiques de test des AR en chaque temps de la fenêtre d'événement. </param>
    ''' <param name="tabMoyCAR"> Moyennes des CAR en chaque temps de la fenêtre d'événement. </param>
    ''' <param name="tabEcartCAR"> Ecart-types des CAR en chaque temps de la fenêtre d'événement. </param>
    ''' <param name="statCAR"> Statistiques de test des CAR en chaque temps de la fenêtre d'événement. </param>
    ''' <param name="tailleEch"> Taille de l'échantillon (i.e le nombre d'entreprises). </param>
    ''' <param name="nomFeuille"> Nom de la feuille où afficher les résultats. </param>
    ''' <param name="decal"> Décalage du nombre de colonnes pour l'affichage.  </param>
    ''' <remarks></remarks>
    Public Sub afficheResAsympt(datesEvAR() As Integer, tabMoyAR() As Double, tabEcartAR() As Double, statAR() As Double, _
                                tabMoyCAR() As Double, tabEcartCAR() As Double, statCAR() As Double, _
                                tailleEch As Integer, nomFeuille As String, decal As Integer)

        'Affichage en-tête
        Globals.ThisAddIn.Application.Worksheets(nomFeuille).Cells(1, decal + 1).Value = "Résultats du T-test asymptotique"
        Globals.ThisAddIn.Application.Worksheets(nomFeuille).Cells(1, decal + 1).Font.Bold = True
        Globals.ThisAddIn.Application.Worksheets(nomFeuille).Cells(1, decal + 1).Interior.ColorIndex = 50
        Globals.ThisAddIn.Application.Worksheets(nomFeuille).Cells(1, decal + 2).Interior.ColorIndex = 50
        Globals.ThisAddIn.Application.Worksheets(nomFeuille).Cells(1, decal + 3).Interior.ColorIndex = 50

        Dim tailleFenetreEv As Integer = datesEvAR.GetLength(0)
        'indice pour l'écriture dans les cellules
        Dim j As Integer

        nomCellule(Globals.ThisAddIn.Application.Worksheets(nomFeuille).Cells(3, decal + 2), "Moyenne")
        nomCellule(Globals.ThisAddIn.Application.Worksheets(nomFeuille).Cells(3, decal + 3), "Ecart-type")
        nomCellule(Globals.ThisAddIn.Application.Worksheets(nomFeuille).Cells(3, decal + 4), "T-statistique")
        nomCellule(Globals.ThisAddIn.Application.Worksheets(nomFeuille).Cells(3, decal + 5), "P-valeur (%)")

        'affichage des résultats sur les AR
        For i = 0 To tailleFenetreEv - 1
            j = i + 4

            nomCellule(Globals.ThisAddIn.Application.Worksheets(nomFeuille).Cells(j, decal + 1), "AR(" & datesEvAR(i) & ")")

            'La colonne des moyennes
            valeurCellule(Globals.ThisAddIn.Application.Worksheets(nomFeuille).Cells(j, decal + 2), tabMoyAR(i))

            'La colonne des écart-type
            valeurCellule(Globals.ThisAddIn.Application.Worksheets(nomFeuille).Cells(j, decal + 3), tabEcartAR(i))

            'La colonne des T-statistiques
            valeurCellule(Globals.ThisAddIn.Application.Worksheets(nomFeuille).Cells(j, decal + 4), statAR(i))

            'La colonne des p-valeurs
            Dim pValeur As Double = TestsStatistiques.calculPValeurStudent(statAR(i), tailleEch)
            valeurCellule(Globals.ThisAddIn.Application.Worksheets(nomFeuille).Cells(j, decal + 5), pValeur * 100)
            'La signification du test
            Globals.ThisAddIn.Application.Worksheets(nomFeuille).Cells(j, decal + 6).Value = signification(pValeur)
        Next i

        'affichage des résultats sur les CAR
        nomCellule(Globals.ThisAddIn.Application.Worksheets(nomFeuille).Cells(tailleFenetreEv + 6, decal + 2), "Moyenne")
        nomCellule(Globals.ThisAddIn.Application.Worksheets(nomFeuille).Cells(tailleFenetreEv + 6, decal + 3), "Ecart-type")
        nomCellule(Globals.ThisAddIn.Application.Worksheets(nomFeuille).Cells(tailleFenetreEv + 6, decal + 4), "T-statistique")
        nomCellule(Globals.ThisAddIn.Application.Worksheets(nomFeuille).Cells(tailleFenetreEv + 6, decal + 5), "P-valeur (%)")

        For i = 0 To tailleFenetreEv - 1
            j = i + tailleFenetreEv + 7

            nomCellule(Globals.ThisAddIn.Application.Worksheets(nomFeuille).Cells(j, decal + 1), "CAR(" & datesEvAR(i) & ")")

            'La colonne des moyennes
            valeurCellule(Globals.ThisAddIn.Application.Worksheets(nomFeuille).Cells(j, decal + 2), tabMoyCAR(i))

            'La colonne des écart-type
            valeurCellule(Globals.ThisAddIn.Application.Worksheets(nomFeuille).Cells(j, decal + 3), tabEcartCAR(i))

            'La colonne des T-statistiques
            valeurCellule(Globals.ThisAddIn.Application.Worksheets(nomFeuille).Cells(j, decal + 4), statCAR(i))

            'La colonne des p-valeurs
            Dim pValeur As Double = TestsStatistiques.calculPValeurStudent(statCAR(i), tailleEch)
            valeurCellule(Globals.ThisAddIn.Application.Worksheets(nomFeuille).Cells(j, decal + 5), pValeur * 100)
            'La signification du test
            Globals.ThisAddIn.Application.Worksheets(nomFeuille).Cells(j, decal + 6).Value = signification(pValeur)
        Next i
    End Sub

    ''' <summary>
    ''' Affichage des résultats des tests statistiques sur les CAR sur la période autour de l'événement.
    ''' </summary>
    ''' <param name="datesEvAR"> Dates de la période d'événement sur laquelle les tests sont réalisés. </param>
    ''' <param name="statAAR"> Statistiques de test des AAR en chaque temps de la fenêtre d'événement. </param>
    ''' <param name="statCAAR"> Statistiques de test des CAAR en chaque temps de la fenêtre d'événement. </param>
    ''' <param name="tailleEch"> Taille de l'échantillon (i.e le nombre d'entreprises). </param>
    ''' <param name="nomFeuille"> Nom de la feuille où afficher les résultats. </param>
    ''' <param name="decal"> Décalage du nombre de colonnes pour l'affichage.  </param>
    ''' <remarks></remarks>
    Public Sub afficheResExact(datesEvAR() As Integer, statAAR() As Double, statCAAR() As Double, tailleEch As Integer, nomFeuille As String, decal As Integer)

        'Affichage en-tête
        Globals.ThisAddIn.Application.Worksheets(nomFeuille).Cells(1, decal + 1).Value = "Résultats du T-test exact"
        Globals.ThisAddIn.Application.Worksheets(nomFeuille).Cells(1, decal + 1).Font.Bold = True
        Globals.ThisAddIn.Application.Worksheets(nomFeuille).Cells(1, decal + 1).Interior.ColorIndex = 50
        Globals.ThisAddIn.Application.Worksheets(nomFeuille).Cells(1, decal + 2).Interior.ColorIndex = 50

        Dim tailleFenetreEv As Integer = datesEvAR.GetLength(0)
        'indice pour l'écriture dans les cellules
        Dim j As Integer

        nomCellule(Globals.ThisAddIn.Application.Worksheets(nomFeuille).Cells(3, decal + 2), "T-statistique")
        nomCellule(Globals.ThisAddIn.Application.Worksheets(nomFeuille).Cells(3, decal + 3), "P-valeur (%)")

        'affichage des résultats sur les AR
        For i = 0 To tailleFenetreEv - 1
            j = i + 4

            nomCellule(Globals.ThisAddIn.Application.Worksheets(nomFeuille).Cells(j, decal + 1), "AR(" & datesEvAR(i) & ")")

            'La colonne des moyennes
            valeurCellule(Globals.ThisAddIn.Application.Worksheets(nomFeuille).Cells(j, decal + 2), statAAR(i))

            'La colonne des p-valeurs
            Dim pValeur As Double = TestsStatistiques.calculPValeurStudent(statAAR(i), tailleEch)
            valeurCellule(Globals.ThisAddIn.Application.Worksheets(nomFeuille).Cells(j, decal + 3), pValeur * 100)
            'La signification du test
            Globals.ThisAddIn.Application.Worksheets(nomFeuille).Cells(j, decal + 4).Value = signification(pValeur)
        Next i

        'affichage des résultats sur les CAR
        nomCellule(Globals.ThisAddIn.Application.Worksheets(nomFeuille).Cells(tailleFenetreEv + 6, decal + 2), "T-statistique")
        nomCellule(Globals.ThisAddIn.Application.Worksheets(nomFeuille).Cells(tailleFenetreEv + 6, decal + 3), "P-valeur (%)")

        For i = 0 To tailleFenetreEv - 1
            j = i + tailleFenetreEv + 7

            nomCellule(Globals.ThisAddIn.Application.Worksheets(nomFeuille).Cells(j, decal + 1), "CAR(" & datesEvAR(i) & ")")

            'La colonne des moyennes
            valeurCellule(Globals.ThisAddIn.Application.Worksheets(nomFeuille).Cells(j, decal + 2), statCAAR(i))

            'La colonne des p-valeurs
            Dim pValeur As Double = TestsStatistiques.calculPValeurStudent(statCAAR(i), tailleEch)
            valeurCellule(Globals.ThisAddIn.Application.Worksheets(nomFeuille).Cells(j, decal + 3), pValeur * 100)
            'La signification du test
            Globals.ThisAddIn.Application.Worksheets(nomFeuille).Cells(j, decal + 4).Value = signification(pValeur)
        Next i
    End Sub

    ''' <summary>
    ''' Assigne un texte à une cellule avec une mise en forme.
    ''' </summary>
    ''' <param name="cell"> Cellule à mettre en forme. </param>
    ''' <param name="texte"> Texte à insérer dans la cellule. </param>
    ''' <remarks></remarks>
    Private Sub nomCellule(cell As Excel.Range, texte As String)
        cell.Value = texte
        cell.Font.Bold = True
        cell.Borders.Value = 1
        cell.Interior.ColorIndex = 27
    End Sub

    ''' <summary>
    ''' Assigne un texte à une cellule avec des bordures sur le tableau.
    ''' </summary>
    ''' <param name="cell"> Cellule à mettre en forme. </param>
    ''' <param name="texte"> Texte à insérer dans la cellule. </param>
    ''' <remarks></remarks>
    Private Sub valeurCellule(cell As Excel.Range, texte As Double)
        cell.Value = texte
        cell.Borders.Value = 1
    End Sub

    ''' <summary>
    ''' Renvoie le nombre d'étoiles en fonction de la P-valeur (même convention qu'en R).
    ''' </summary>
    ''' <param name="pValeur"> P-Valeur du test. </param>
    ''' <returns> Code de signification du test. </returns>
    ''' <remarks></remarks>
    Function signification(pValeur As Double) As String
        Select Case pValeur
            Case Is < 0.001
                signification = "***"
            Case Is < 0.01
                signification = "**"
            Case Is < 0.05
                signification = "*"
            Case Is < 0.1
                signification = "."
            Case Else
                signification = ""
        End Select
    End Function

    ''' <summary>
    ''' Affichage des rentabilités des entreprises centrées autour de la date d'événement pour chaque entreprise.
    ''' La date 0 correspond à la date d'événement.
    ''' </summary>
    ''' <param name="tabrenta"> Rentabilités des entreprises centrées. </param>
    ''' <remarks></remarks>
    Public Sub affichageRentaCentrees(tabrenta(,) As Double)
        'Création d'une nouvelle feuille
        Dim nom As String
        nom = InputBox("Entrer le nom de la feuille des rentabilités centrées : ")
        'Si l'utilisateur n'entre pas un nom
        If nom Is "" Then nom = "Rentabilités centrées"
        Globals.ThisAddIn.Application.Sheets.Add(After:=Globals.ThisAddIn.Application.Worksheets(Globals.ThisAddIn.Application.Worksheets.Count))
        Try
            Globals.ThisAddIn.Application.ActiveSheet.Name = nom
        Catch ex As System.Runtime.InteropServices.COMException
            MsgBox(ex.Message, 16)
            nom = Globals.ThisAddIn.Application.ActiveSheet.Name
        End Try

        'Affichage des dates
        Globals.ThisAddIn.Application.Worksheets(nom).Range("A1").Value = "Dates"
        For i = 0 To tabrenta.GetUpperBound(0)
            Globals.ThisAddIn.Application.Worksheets(nom).Range("A" & i + 2).Value = tabrenta(i, 0)
            Globals.ThisAddIn.Application.Worksheets(nom).Range("A" & i + 2).Borders.Value = 1
        Next i

        'On écrit la première ligne
        For colonne = 1 To tabrenta.GetUpperBound(1)
            Globals.ThisAddIn.Application.Worksheets(nom).Cells(1, colonne + 1).Value = "R" & colonne
        Next colonne

        'Affichage des rentabilités
        For colonne = 1 To tabrenta.GetUpperBound(1)
            For i = 0 To tabrenta.GetUpperBound(0)
                If Double.IsNaN(tabrenta(i, colonne)) Then
                    Globals.ThisAddIn.Application.Worksheets(nom).Cells(i + 2, colonne + 1).Value = "#N/A"
                Else
                    Globals.ThisAddIn.Application.Worksheets(nom).Cells(i + 2, colonne + 1).Value = tabrenta(i, colonne)
                End If
            Next i
        Next colonne
    End Sub

    ''' <summary>
    ''' Affichage des AR sur les fenêtres d'estimation et d'événement.
    ''' </summary>
    ''' <param name="tabAREst"> AR sur la fenêtre d'estimation. </param>
    ''' <param name="tabAREv"> AR sur la fenêtre d'événement. </param>
    ''' <param name="tabDateEst"> Dates correspondantes sur la fenêtre d'estimation. </param>
    ''' <param name="tabDateEv"> Dates correspondantes sur la fenêtre d'événement. </param>
    ''' <remarks></remarks>
    Public Sub affichageAR(ByRef tabAREst(,) As Double, ByRef tabAREv(,) As Double, _
                           ByRef tabDateEst() As Integer, ByRef tabDateEv() As Integer)
        'Création d'une nouvelle feuille
        Dim nom As String
        nom = InputBox("Entrer le nom de la feuille des rentabilités anormales : ")
        'Si l'utilisateur n'entre pas un nom
        If nom Is "" Then nom = "Rentabilités anormales"
        Globals.ThisAddIn.Application.Sheets.Add(After:=Globals.ThisAddIn.Application.Worksheets(Globals.ThisAddIn.Application.Worksheets.Count))
        Try
            Globals.ThisAddIn.Application.ActiveSheet.Name = nom
        Catch ex As System.Runtime.InteropServices.COMException
            MsgBox(ex.Message, 16)
            nom = Globals.ThisAddIn.Application.ActiveSheet.Name
        End Try

        Dim currentSheet As Excel.Worksheet = CType(Globals.ThisAddIn.Application.Worksheets(nom), Excel.Worksheet)

        'Affichage en-tête estimation
        currentSheet.Cells(1, 1).Value = "AR sur la période d'estimation"
        currentSheet.Cells(1, 1).Font.Bold = True
        currentSheet.Cells(1, 1).Interior.ColorIndex = 50
        currentSheet.Cells(1, 2).Interior.ColorIndex = 50
        currentSheet.Cells(1, 3).Interior.ColorIndex = 50

        'Affichage de la première ligne
        For i = 0 To tabAREst.GetUpperBound(1)
            nomCellule(currentSheet.Cells(3, i + 2), "AR" & i + 1)
        Next i

        'Affichage des dates pour la période d'estimation
        nomCellule(currentSheet.Cells(3, 1), "Dates")
        For i = 0 To tabDateEst.GetUpperBound(0)
            nomCellule(currentSheet.Cells(i + 4, 1), tabDateEst(i).ToString)
        Next i

        'Affichage des données pour la période d'estimation
        For colonne = 0 To tabAREst.GetUpperBound(1)
            For i = 0 To tabAREst.GetUpperBound(0)
                If Double.IsNaN(tabAREst(i, colonne)) Then
                    currentSheet.Cells(i + 4, colonne + 2).Value = "#N/A"
                    currentSheet.Cells(i + 4, colonne + 2).Borders.Value = 1
                Else
                    valeurCellule(currentSheet.Cells(i + 4, colonne + 2), tabAREst(i, colonne))
                End If
            Next i
        Next colonne

        'Affichage en-tête événement
        currentSheet.Cells(tabAREst.GetUpperBound(0) + 7, 1).Value = "AR sur la période d'événement"
        currentSheet.Cells(tabAREst.GetUpperBound(0) + 7, 1).Font.Bold = True
        currentSheet.Cells(tabAREst.GetUpperBound(0) + 7, 1).Interior.ColorIndex = 50
        currentSheet.Cells(tabAREst.GetUpperBound(0) + 7, 2).Interior.ColorIndex = 50
        currentSheet.Cells(tabAREst.GetUpperBound(0) + 7, 3).Interior.ColorIndex = 50

        'Affichage de la première ligne
        For i = 0 To tabAREst.GetUpperBound(1)
            nomCellule(currentSheet.Cells(8 + tabDateEst.GetLength(0), i + 2), "AR" & i + 1)
        Next i

        'Affichage des dates pour la période d'événement
        nomCellule(currentSheet.Cells(8 + tabDateEst.GetLength(0), 1), "Dates")
        For i = 0 To tabDateEv.GetUpperBound(0)
            nomCellule(currentSheet.Cells(9 + tabDateEst.GetLength(0) + i, 1), tabDateEv(i).ToString)
        Next i

        'Affichage des données pour la période d'événement
        For colonne = 0 To tabAREv.GetUpperBound(1)
            For i = 0 To tabAREv.GetUpperBound(0)
                If Double.IsNaN(tabAREv(i, colonne)) Then
                    currentSheet.Cells(9 + tabDateEst.GetLength(0) + i, colonne + 2).Value = "#N/A"
                    currentSheet.Cells(9 + tabDateEst.GetLength(0) + i, colonne + 2).Borders.Value = 1
                Else
                    valeurCellule(currentSheet.Cells(9 + tabDateEst.GetLength(0) + i, colonne + 2), tabAREv(i, colonne))
                End If
            Next i
        Next colonne

    End Sub

    ''' <summary>
    ''' Affichage des résultats du test de Patell.
    ''' </summary>
    ''' <param name="tabDateEv"> Dates correspondant à la fenêtre d'événement. </param>
    ''' <param name="testHypAAR"> Statistiques de test pour les AAR. </param>
    ''' <param name="testHypCAAR"> Statistiques de test pour les CAAR. </param>
    ''' <remarks></remarks>
    Public Sub affichagePatell(ByRef tabDateEv() As Integer, ByRef testHypAAR() As Double, ByRef testHypCAAR() As Double)
        'Création d'une nouvelle feuille
        Dim nom As String
        nom = InputBox("Entrer le nom de la feuille des résultats : ")
        'Si l'utilisateur n'entre pas un nom
        If nom Is "" Then nom = "Résultats Patell"
        Globals.ThisAddIn.Application.Sheets.Add(After:=Globals.ThisAddIn.Application.Worksheets(Globals.ThisAddIn.Application.Worksheets.Count))
        Try
            Globals.ThisAddIn.Application.ActiveSheet.Name = nom
        Catch ex As System.Runtime.InteropServices.COMException
            MsgBox(ex.Message, 16)
            nom = Globals.ThisAddIn.Application.ActiveSheet.Name
        End Try

        'Affichage en-tête
        Globals.ThisAddIn.Application.Worksheets(nom).Cells(1, 1).Value = "Résultats du test de Patell"
        Globals.ThisAddIn.Application.Worksheets(nom).Cells(1, 1).Font.Bold = True
        Globals.ThisAddIn.Application.Worksheets(nom).Cells(1, 1).Interior.ColorIndex = 50
        Globals.ThisAddIn.Application.Worksheets(nom).Cells(1, 2).Interior.ColorIndex = 50
        Globals.ThisAddIn.Application.Worksheets(nom).Cells(1, 3).Interior.ColorIndex = 50

        '*** Test AAR = 0 ***

        'Le nom de chaque colonne
        nomCellule(Globals.ThisAddIn.Application.Worksheets(nom).Range("B3"), "Test Patell")
        nomCellule(Globals.ThisAddIn.Application.Worksheets(nom).Range("C3"), "P-valeur (%)")

        'Affichage des dates et des statistiques du test de Patell et de la P-Valeur
        For i = 0 To tabDateEv.GetUpperBound(0)
            nomCellule(Globals.ThisAddIn.Application.Worksheets(nom).cells(i + 4, 1), "AR(" & tabDateEv(i) & ")")
            valeurCellule(Globals.ThisAddIn.Application.Worksheets(nom).cells(i + 4, 2), testHypAAR(i))
            Dim pValeur As Double
            pValeur = 2 * (1 - Globals.ThisAddIn.Application.WorksheetFunction.Norm_S_Dist(Math.Abs(testHypAAR(i)), True))
            valeurCellule(Globals.ThisAddIn.Application.Worksheets(nom).cells(i + 4, 3), pValeur * 100)
            'La signification du test
            Globals.ThisAddIn.Application.Worksheets(nom).Cells(i + 4, 4).Value = signification(pValeur)
        Next i

        '*** Test CAAR = 0 ***

        Dim debutAffichage As Integer = tabDateEv.GetLength(0) + 6

        'Le nom de chaque colonne
        nomCellule(Globals.ThisAddIn.Application.Worksheets(nom).Cells(debutAffichage, 2), "Test Patell")
        nomCellule(Globals.ThisAddIn.Application.Worksheets(nom).Cells(debutAffichage, 3), "P-valeur (%)")

        'Affichage des dates et des statistiques du test de Patell et de la P-Valeur
        For i = 0 To tabDateEv.GetUpperBound(0)
            nomCellule(Globals.ThisAddIn.Application.Worksheets(nom).cells(i + debutAffichage + 1, 1), "CAR(" & tabDateEv(i) & ")")
            valeurCellule(Globals.ThisAddIn.Application.Worksheets(nom).cells(i + debutAffichage + 1, 2), testHypCAAR(i))
            Dim pValeur As Double
            pValeur = 2 * (1 - Globals.ThisAddIn.Application.WorksheetFunction.Norm_S_Dist(Math.Abs(testHypCAAR(i)), True))
            valeurCellule(Globals.ThisAddIn.Application.Worksheets(nom).cells(i + debutAffichage + 1, 3), pValeur * 100)
            'La signification du test
            Globals.ThisAddIn.Application.Worksheets(nom).Cells(i + debutAffichage + 1, 4).Value = signification(pValeur)
        Next i

    End Sub

    ''' <summary>
    ''' Affichage des résultats du test de signe.
    ''' </summary>
    ''' <param name="tabDateEv"> Dates correspondant à la fenêtre d'événement. </param>
    ''' <param name="tabEstAR"> AR sur la fenêtre d'estimation. </param>
    ''' <param name="tabEvAR"> AR sur la fenêntre d'événement. </param>
    ''' <remarks></remarks>
    Public Sub affichageSigne(ByRef tabDateEv() As Integer, ByRef tabEstAR(,) As Double, ByRef tabEvAR(,) As Double)
        Dim tailleFenetreEv As Integer = tabEvAR.GetLength(0)

        'Création d'une nouvelle feuille
        Dim nom As String
        nom = InputBox("Entrer le nom de la feuille des résultats : ")
        'Si l'utilisateur n'entre pas un nom
        If nom Is "" Then nom = "Résultats Signe"
        Globals.ThisAddIn.Application.Sheets.Add(After:=Globals.ThisAddIn.Application.Worksheets(Globals.ThisAddIn.Application.Worksheets.Count))
        Try
            Globals.ThisAddIn.Application.ActiveSheet.Name = nom
        Catch ex As System.Runtime.InteropServices.COMException
            MsgBox(ex.Message, 16)
            nom = Globals.ThisAddIn.Application.ActiveSheet.Name
        End Try


        'Affichage en-tête
        Globals.ThisAddIn.Application.Worksheets(nom).Cells(1, 1).Value = "Résultats du test de signe"
        Globals.ThisAddIn.Application.Worksheets(nom).Cells(1, 1).Font.Bold = True
        Globals.ThisAddIn.Application.Worksheets(nom).Cells(1, 1).Interior.ColorIndex = 50
        Globals.ThisAddIn.Application.Worksheets(nom).Cells(1, 2).Interior.ColorIndex = 50

        'Le nom de chaque colonne
        nomCellule(Globals.ThisAddIn.Application.Worksheets(nom).Range("B3"), "Test signe")
        nomCellule(Globals.ThisAddIn.Application.Worksheets(nom).Range("C3"), "P-valeur (%)")

        'Appel de la fonction qui calcule la statistique du test de signe
        Dim stat() As Double = TestsStatistiques.statTestSigne(tabEstAR, tabEvAR)

        'Affichage des dates et des statistiques du test de Patell et de la P-Valeur
        For i = 0 To tailleFenetreEv - 1
            nomCellule(Globals.ThisAddIn.Application.Worksheets(nom).cells(i + 4, 1), tabDateEv(i))
            valeurCellule(Globals.ThisAddIn.Application.Worksheets(nom).cells(i + 4, 2), stat(i))
            Dim pValeur As Double
            'Calcul de la p-valeur
            pValeur = TestsStatistiques.calculPValeurTestSigne(stat(i))
            valeurCellule(Globals.ThisAddIn.Application.Worksheets(nom).cells(i + 4, 3), pValeur * 100)
            'La signification du test
            Globals.ThisAddIn.Application.Worksheets(nom).Cells(i + 4, 4).Value = signification(pValeur)
        Next i

    End Sub

End Module
