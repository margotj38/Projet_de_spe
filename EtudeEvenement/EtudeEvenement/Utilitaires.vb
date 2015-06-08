''' <summary>
''' bla bla bla
''' </summary>
''' <remarks></remarks>

Module Utilitaires

    '***************************** Pour parser les données de refEdit *****************************

    Public Sub parserPlageColonnes(plage As String, ByRef premiereCol As Integer, ByRef derniereCol As Integer)
        Dim rangePlage As Excel.Range = Globals.ThisAddIn.Application.Range(plage)
        premiereCol = rangePlage.Cells(1, 1).Column()
        derniereCol = rangePlage.Cells(1, rangePlage.Columns.Count).Column()
    End Sub

    Public Sub parserPlageLignes(plage As String, ByRef debut As Integer, ByRef fin As Integer)
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


    '***************************** Algo de tri *****************************

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


End Module
