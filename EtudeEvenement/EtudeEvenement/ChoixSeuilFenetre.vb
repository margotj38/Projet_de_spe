Imports System.Diagnostics
Imports System.Windows.Forms.DataVisualization.Charting

Public Class ChoixSeuilFenetre

    Private textFenetreDebut As String
    Private textFenetreFin As String

    Private model As Integer ' 0 => ModeleMoyenne; 1 => ModeleMarcheSimple; 2 => ModeleMarche
    Private numTest As Integer ' 0 => TestSimple; 1 => TestPatell; 2 => TestSigne

    'constructeur
    Public Sub New(ByVal model As Integer, ByVal test As Integer)
        InitializeComponent()
        Me.model = model
        Me.numTest = test
    End Sub

    'accesseur sur model
    Public Property modele() As Integer
        Get
            Return model
        End Get
        Set(value As Integer)
            If value < 0 Or value > 2 Then
                MsgBox("Erreur interne : numéro de modèle incorrect", 16)
            End If
            model = value
        End Set
    End Property

    'accesseur sur numTest
    Public Property test() As Integer
        Get
            Return numTest
        End Get
        Set(value As Integer)
            If value < 0 Or value > 2 Then
                MsgBox("Erreur interne : numéro de test incorrect", 16)
            End If
            numTest = value
        End Set
    End Property

    Private Sub FenetreDebBox_TextChanged(sender As Object, e As EventArgs) Handles FenetreDebBox.TextChanged
        textFenetreDebut = FenetreDebBox.Text
    End Sub

    Private Sub FenetreFinBox_TextChanged(sender As Object, e As EventArgs) Handles FenetreFinBox.TextChanged
        textFenetreFin = FenetreFinBox.Text
    End Sub

    Private Sub PValeur_Click(sender As Object, e As EventArgs) Handles PValeur.Click
        Try
            Dim fenetreDebut As Integer = CInt(textFenetreDebut)
            Dim fenetreFin As Integer = CInt(textFenetreFin)
            Dim currentSheet As Excel.Worksheet = CType(Globals.ThisAddIn.Application.Worksheets("Rt"), Excel.Worksheet)
            Dim premiereDate As Integer = currentSheet.Cells(2, 1).Value
            Dim tailleEchant As Integer = currentSheet.UsedRange.Columns.Count - 1
            If fenetreDebut > fenetreFin Or fenetreDebut <= premiereDate Or fenetreFin > premiereDate + currentSheet.UsedRange.Rows.Count - 1 Then
                MsgBox("Erreur : La fenêtre de temps de l'événement doit être cohérente avec les données", 16)
            Else
                'Calcul des AR
                Dim tabAR As Double(,)
                tabAR = Globals.ThisAddIn.calculAR(fenetreDebut)
                For i = 0 To tabAR.GetUpperBound(0)
                    Debug.WriteLine(tabAR(i, 0))
                Next
                'Calcul de la pValeur
                Dim pValeur As Double
                Select Case test
                    Case 0
                        'test simple'
                        Dim tabCAR As Double()
                        tabCAR = Globals.ThisAddIn.calculCAR(tabAR, fenetreDebut, fenetreFin)
                        Dim testHyp As Double = Globals.ThisAddIn.calculStatistique(tabCAR)
                        pValeur = Globals.ThisAddIn.calculPValeur(tailleEchant, testHyp) * 100
                    Case 1
                        'test de Patell'
                        Dim testHyp As Double = Globals.ThisAddIn.patellTest(tabAR, fenetreDebut, fenetreFin)
                        pValeur = 2 * (1 - Globals.ThisAddIn.Application.WorksheetFunction.Norm_S_Dist(Math.Abs(testHyp), True)) * 100
                    Case 2
                        'test de signe : à compléter'
                        pValeur = 0
                End Select
                
                MsgBox("P-Valeur : " & pValeur.ToString("0.0000") & "%")
                Globals.Ribbons.Ruban.seuilFenetreTaskPane.Visible = False
            End If
        Catch erreur As InvalidCastException
            MsgBox("Erreur : Vous devez entrer des données correctes (utiliser la virgule pour les nombres décimaux)", 16)
        End Try
    End Sub

    Private Sub PValeurFenetre_Click(sender As Object, e As EventArgs) Handles PValeurFenetre.Click
        Try
            Dim fenetreDebut As Integer = CInt(textFenetreDebut)
            Dim fenetreFin As Integer = CInt(textFenetreFin)
            Dim currentSheet As Excel.Worksheet = CType(Globals.ThisAddIn.Application.Worksheets("Rt"), Excel.Worksheet)
            Dim tailleEchant As Integer = currentSheet.UsedRange.Columns.Count - 1
            Dim premiereDate As Integer = currentSheet.Cells(2, 1).Value
            Dim derniereDate As Integer = premiereDate + currentSheet.UsedRange.Rows.Count - 2
            Dim maxFenetre As Integer = Math.Min(Math.Abs(premiereDate), Math.Abs(derniereDate))

            If fenetreDebut > fenetreFin Or fenetreDebut <= premiereDate Or fenetreFin > premiereDate + currentSheet.UsedRange.Rows.Count - 1 Then
                MsgBox("Erreur : La fenêtre de temps de l'événement doit être cohérente avec les données", 16)
            Else
                'Calcul des pvaleurs et affichage de la courbe
                Globals.ThisAddIn.tracerPValeur(tailleEchant, maxFenetre)
                Globals.Ribbons.Ruban.seuilFenetreTaskPane.Visible = False
                Globals.Ribbons.Ruban.graphPVal.Visible = True
            End If
        Catch erreur As InvalidCastException
            MsgBox("Erreur : Vous devez entrer des données correctes (utiliser la virgule pour les nombres décimaux)", 16)
        End Try
    End Sub

    'Private Sub BoutonPatell_Click(sender As Object, e As EventArgs) Handles BoutonPatell.Click
    '    Try
    '        Dim fenetreDebut As Integer = CInt(textFenetreDebut)
    '        Dim fenetreFin As Integer = CInt(textFenetreFin)
    '        Dim currentSheet As Excel.Worksheet = CType(Globals.ThisAddIn.Application.Worksheets("Rt"), Excel.Worksheet)
    '        Dim premiereDate As Integer = currentSheet.Cells(2, 1).Value
    '        Dim tailleEchant As Integer = currentSheet.UsedRange.Columns.Count - 1
    '        If fenetreDebut > fenetreFin Or fenetreDebut <= premiereDate Or fenetreFin > premiereDate + currentSheet.UsedRange.Rows.Count - 1 Then
    '            MsgBox("Erreur : La fenêtre de temps de l'événement doit être cohérente avec les données", 16)
    '        Else
    '            'Calcul de la pvaleur
    '            Dim tabAR As Double(,)
    '            tabAR = Globals.ThisAddIn.calculAR(fenetreDebut)

    '            Dim testHyp As Double = Globals.ThisAddIn.patellTest(tabAR, fenetreDebut, fenetreFin)
    '            Dim pValeur As Double = 2 * (1 - Globals.ThisAddIn.Application.WorksheetFunction.Norm_S_Dist(Math.Abs(testHyp), True)) * 100

    '            MsgBox("P-Valeur : " & pValeur.ToString("0.0000") & "%")
    '            Globals.Ribbons.Ruban.seuilFenetreTaskPane.Visible = False
    '        End If
    '    Catch erreur As InvalidCastException
    '        MsgBox("Erreur : Vous devez entrer des données correctes (utiliser la virgule pour les nombres décimaux)", 16)
    '    End Try
    'End Sub

End Class
