''' <summary>
''' Module regroupant des fonctions annexes diverses.
''' </summary>
''' <remarks></remarks>

Module Utilitaires

    '***************************** Calcul *****************************

    ''' <summary>
    ''' Fonction calculant une moyenne sur un tableau en tenant compte de possible "Double.Nan".
    ''' </summary>
    ''' <param name="tab">Tableau dont on veut calculer la moyenne.</param>
    ''' <returns>La moyenne des éléments du tableau.</returns>
    ''' <remarks></remarks>
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

    ''' <summary>
    ''' Fonction calculant une variance sur un tableau en tenant compte de possible "Double.Nan".
    ''' </summary>
    ''' <param name="tab">Tableau dont on veut calculer la variance.</param>
    ''' <param name="moyenne">Moyenne du tableau.</param> 
    ''' <returns>La variance des éléments du tableau.</returns>
    ''' <remarks></remarks>
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

    '***************************** Pour parser les données de refEdit *****************************

    ''' <summary>
    ''' Extrait la première et dernière colonne d'une plage de données.
    ''' </summary>
    ''' <param name="plage"> Plage de données Excel sous forme de chaine de caractères. </param>
    ''' <param name="premiereCol"> Indice de la première colonne de la plage de données. </param>
    ''' <param name="derniereCol"> Indice de la dernière colonne de la plage de données. </param>
    ''' <remarks></remarks>
    Public Sub parserPlageColonnes(plage As String, ByRef premiereCol As Integer, ByRef derniereCol As Integer)
        Dim rangePlage As Excel.Range = Globals.ThisAddIn.Application.Range(plage)
        premiereCol = rangePlage.Cells(1, 1).Column()
        derniereCol = rangePlage.Cells(1, rangePlage.Columns.Count).Column()
    End Sub

    ''' <summary>
    ''' Extrait la première et dernière ligne d'une plage de données.
    ''' </summary>
    ''' <param name="plage"> Plage de données Excel sous forme de chaine de caractères. </param>
    ''' <param name="premiereLigne"> Indice de la première ligne de la plage de données. </param>
    ''' <param name="derniereLigne"> Indice de la dernière ligne de la plage de données. </param>
    ''' <remarks></remarks>
    Public Sub parserPlageLignes(plage As String, ByRef premiereLigne As Integer, ByRef derniereLigne As Integer)
        Dim rangePlage As Excel.Range = Globals.ThisAddIn.Application.Range(plage)
        premiereLigne = rangePlage.Cells(1, 1).Row()
        derniereLigne = rangePlage.Cells(rangePlage.Rows.Count, 1).Row()
    End Sub

    ''' <summary>
    ''' Sépare le nom de la feuille et la plage de données à partir de la sélection avec RefEdit.
    ''' </summary>
    ''' <param name="textRefEdit"> Chaine de caractères renvoyée par RefEdit. </param>
    ''' <param name="feuille"> Nom de la feuille associé à la sélection avec RefEdit. </param>
    ''' <param name="plage"> Plage de données sous forme de chaine de caractères associée à la sélection avec RefEdit. </param>
    ''' <remarks></remarks>
    Public Sub recupererFeuillePlage(textRefEdit As String, ByRef feuille As String, ByRef plage As String)
        Dim tabString() As String = Split(textRefEdit, "'")
        feuille = tabString(1)
        'On enlève le '!' 
        tabString = Split(tabString(2), "!")
        plage = tabString(1)
    End Sub


    '***************************** Algo de tri *****************************

    ''' <summary>
    ''' Tri (Quick sort) deux tableaux une dimension selon l'ordre chronologique sur le premier tableau de dates.
    ''' </summary>
    ''' <param name="tabDate"> Tableau de dates. </param>
    ''' <param name="tabInd"> (Entrée-sortie) Tableau de permutations. </param>
    ''' <param name="gauche"> Indice inférieur du tableau. </param>
    ''' <param name="droite"> Indice supérieur du tableau. </param>
    ''' <remarks> Il s'agit d'une fonction récursive d'où les deux derniers paramètres. </remarks>
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

End Module
