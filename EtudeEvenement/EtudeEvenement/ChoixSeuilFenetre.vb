﻿Imports System.Diagnostics
Imports System.Windows.Forms.DataVisualization.Charting

Public Class ChoixSeuilFenetre

    Private textFenetreEstDebut As String
    Private textFenetreEstFin As String
    Private textFenetreEvDebut As String
    Private textFenetreEvFin As String

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

    Private Sub FenetreEstDebBox_TextChanged(sender As Object, e As EventArgs) Handles FenetreEstDebBox.TextChanged
        textFenetreEstDebut = FenetreEstDebBox.Text
    End Sub

    Private Sub FenetreEstFinBox_TextChanged(sender As Object, e As EventArgs) Handles FenetreEstFinBox.TextChanged
        textFenetreEstFin = FenetreEstFinBox.Text
    End Sub

    Private Sub FenetreDebBox_TextChanged(sender As Object, e As EventArgs) Handles FenetreDebBox.TextChanged
        textFenetreEvDebut = FenetreDebBox.Text
    End Sub

    Private Sub FenetreFinBox_TextChanged(sender As Object, e As EventArgs) Handles FenetreFinBox.TextChanged
        textFenetreEvFin = FenetreFinBox.Text
    End Sub

    Private Sub PValeur_Click(sender As Object, e As EventArgs) Handles PValeur.Click
        Try
            Dim fenetreEstDebut As Integer = CInt(textFenetreEstDebut)
            Dim fenetreEstFin As Integer = CInt(textFenetreEstFin)
            Dim fenetreEvDebut As Integer = CInt(textFenetreEvDebut)
            Dim fenetreEvFin As Integer = CInt(textFenetreEvFin)

            Dim currentSheet As Excel.Worksheet = CType(Globals.ThisAddIn.Application.Worksheets("Rt"), Excel.Worksheet)
            Dim premiereDate As Integer = currentSheet.Cells(2, 1).Value
            Dim tailleEchant As Integer = currentSheet.UsedRange.Columns.Count - 1

            If fenetreEvDebut > fenetreEvFin Or fenetreEvFin > premiereDate + currentSheet.UsedRange.Rows.Count - 1 _
                Or fenetreEstDebut > fenetreEstFin Or fenetreEstDebut < premiereDate Or fenetreEstFin >= fenetreEvDebut Then
                MsgBox("Erreur : La fenêtre de temps de l'événement doit être cohérente avec les données", 16)
            Else
                'Calcul des AR
                Dim tabAR As Double(,)
                tabAR = Globals.ThisAddIn.calculAR(fenetreEstDebut, fenetreEstFin)
                'For i = 0 To tabAR.GetUpperBound(0)
                'Debug.WriteLine(tabAR(i, 0))
                'Next
                'Calcul de la pValeur
                Dim pValeur As Double
                Select Case test
                    Case 0
                        'test simple'
                        Dim tabCAR As Double()
                        tabCAR = Globals.ThisAddIn.calculCAR(tabAR, fenetreEstDebut, fenetreEstFin, fenetreEvDebut, fenetreEvFin)
                        Dim testHyp As Double = Globals.ThisAddIn.calculStatistique(tabCAR)
                        pValeur = Globals.ThisAddIn.calculPValeur(tailleEchant, testHyp) * 100
                    Case 1
                        'test de Patell'
                        Dim testHyp As Double = Globals.ThisAddIn.patellTest(tabAR, fenetreEstDebut, fenetreEstFin, fenetreEvDebut, fenetreEvFin)
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
        Dim currentSheet As Excel.Worksheet = CType(Globals.ThisAddIn.Application.Worksheets("Rt"), Excel.Worksheet)
        Dim premiereDate As Integer = currentSheet.Cells(2, 1).Value
        Dim derniereDate As Integer = premiereDate + currentSheet.UsedRange.Rows.Count - 2
        Dim tailleEchant As Integer = currentSheet.UsedRange.Columns.Count - 1

        'Calcul des pvaleurs et affichage de la courbe
        Globals.ThisAddIn.tracerPValeur(tailleEchant, derniereDate)
        Globals.Ribbons.Ruban.seuilFenetreTaskPane.Visible = False
        Globals.Ribbons.Ruban.graphPVal.Visible = True
    End Sub

End Class
