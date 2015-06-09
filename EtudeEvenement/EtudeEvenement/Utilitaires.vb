''' <summary>
''' Module regroupant des fonctions annexes diverses.
''' </summary>
''' <remarks></remarks>

Module Utilitaires

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
