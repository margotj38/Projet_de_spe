﻿
Imports System.Windows.Forms.DataVisualization.Charting

Public Class ChoixSeuilFenetre

    Private textFenetreDebut As String
    Private textFenetreFin As String

    Private model As Integer ' 0 => ModeleMoyenne; 1 => ModeleRentaMarche; 2 => ModeleMarche

    Public Sub New(ByVal model As Integer)
        InitializeComponent()
        Me.model = model
    End Sub

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

    Private Sub FenetreBox_TextChanged(sender As Object, e As EventArgs) Handles FenetreBox.TextChanged
        textFenetreDebut = FenetreBox.Text
    End Sub

    Private Sub FenetreFinBox_TextChanged(sender As Object, e As EventArgs) Handles FenetreFinBox.TextChanged
        textFenetreFin = FenetreFinBox.Text
    End Sub

    Private Sub PValeur_Click(sender As Object, e As EventArgs) Handles PValeur.Click
        'Try
        Dim fenetreDebut As Integer = CInt(textFenetreDebut)
        Dim fenetreFin As Integer = CInt(textFenetreFin)
        Dim currentSheet As Excel.Worksheet = CType(Globals.ThisAddIn.Application.Worksheets("Rt"), Excel.Worksheet)
        Dim premiereDate As Integer = currentSheet.Cells(2, 1).Value
        Dim tailleEchant As Integer = currentSheet.UsedRange.Columns.Count - 1
        If fenetreDebut > fenetreFin Or fenetreDebut <= premiereDate Or fenetreFin > premiereDate + currentSheet.UsedRange.Rows.Count - 1 Then
            MsgBox("Erreur : La fenêtre de temps de l'événement doit être cohérente avec les données", 16)
        Else
            'Calcul de la pvaleur
            Dim tabAR As Double(,)
            tabAR = Globals.ThisAddIn.calculAR(fenetreDebut, fenetreFin)
            ''
            'Pour test Patell
            Dim testHyp As Double = Globals.ThisAddIn.patellTest(tabAR, fenetreFin - fenetreDebut + 1, fenetreDebut, fenetreFin)
            Dim pValeur As Double = Globals.ThisAddIn.Application.WorksheetFunction.Norm_S_Dist(testHyp, True)
            'Dim tabCAR As Double()
            'tabCAR = Globals.ThisAddIn.calculCAR(tabAR, fenetreDebut, fenetreFin)
            'Dim testHyp As Double = Globals.ThisAddIn.calculStatistique(tabCAR)
            '
            'Dim pValeur As Double = Globals.ThisAddIn.calculPValeur(tailleEchant, testHyp) * 100
            ''
            MsgBox("P-Valeur : " & PValeur.ToString("0.0000") & "%")
            Globals.Ribbons.Ruban.seuilFenetreTaskPane.Visible = False
        End If
        'Catch erreur As InvalidCastException
        'MsgBox("Erreur : Vous devez entrer des données correctes (utiliser la virgule pour les nombres décimaux)", 16)
        'End Try
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
End Class
