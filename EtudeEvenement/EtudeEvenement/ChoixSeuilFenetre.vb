Imports System.Diagnostics
Imports System.Windows.Forms.DataVisualization.Charting
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices

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

    Private Sub FenetreEstDebBox_TextChanged(sender As Object, e As EventArgs)
        'textFenetreEstDebut = FenetreEstDebBox.Text
    End Sub

    Private Sub FenetreEstFinBox_TextChanged(sender As Object, e As EventArgs)
        'textFenetreEstFin = FenetreEstFinBox.Text
    End Sub

    Private Sub FenetreDebBox_TextChanged(sender As Object, e As EventArgs)
        'textFenetreEvDebut = FenetreDebBox.Text
    End Sub

    Private Sub FenetreFinBox_TextChanged(sender As Object, e As EventArgs)
        'textFenetreEvFin = FenetreFinBox.Text
    End Sub

    Private Sub LancementEtEv_Click(sender As Object, e As EventArgs) Handles LancementEtEv.Click
        'Try
        Dim fenetreEstDebut As Integer = 0 'CInt(textFenetreEstDebut)
        Dim fenetreEstFin As Integer = 0 'CInt(textFenetreEstFin)
        Dim fenetreEvDebut As Integer = 0 'CInt(textFenetreEvDebut)
        Dim fenetreEvFin As Integer = 0 'CInt(textFenetreEvFin)

        'On effectue le centrage des prix autour des événements
        '''''Commenté pour debug 
        'PretraitementPrix.prixCentres()

        Dim currentSheet As Excel.Worksheet = CType(Globals.ThisAddIn.Application.Worksheets("prixCentres"), Excel.Worksheet)
        Dim premiereDate As Integer = currentSheet.Cells(2, 1).Value
        Dim tailleEchant As Integer = currentSheet.UsedRange.Columns.Count - 1

        If fenetreEvDebut > fenetreEvFin Or fenetreEvFin > premiereDate + currentSheet.UsedRange.Rows.Count - 1 _
            Or fenetreEstDebut > fenetreEstFin Or fenetreEstDebut < premiereDate + 1 Or fenetreEstFin >= fenetreEvDebut Then
            MsgBox("Erreur : La fenêtre de temps de l'événement doit être cohérente avec les données", 16)
        Else
            'Calcul des AR
            Dim tabAR As Double(,)
            'tabAR = RentaAnormales.calculARAvecNA(fenetreEstDebut, fenetreEstFin, fenetreEvDebut, fenetreEvFin)
            'tabAR = Globals.ThisAddIn.calculAR(fenetreEstDebut, fenetreEstFin)
            'Calcul de la pValeur
            Dim pValeur As Double
            Select Case test
                Case 0
                    'test simple'
                    Dim tabCAR As Double()
                    tabCAR = TestsStatistiques.calculCAR(tabAR, premiereDate + 1, fenetreEstDebut, fenetreEstFin, fenetreEvDebut, fenetreEvFin)
                    Dim testHyp As Double = TestsStatistiques.calculStatStudent(tabCAR)
                    pValeur = TestsStatistiques.calculPValeur(tailleEchant, testHyp) * 100
                Case 1
                    'test de Patell'
                    'Dim testHyp As Double = TestsStatistiques.patellTest(tabAR, fenetreEstDebut, fenetreEstFin, fenetreEvDebut, fenetreEvFin)
                    'pValeur = 2 * (1 - Globals.ThisAddIn.Application.WorksheetFunction.Norm_S_Dist(Math.Abs(testHyp), True)) * 100
                Case 2
                    'test de signe'
                    'Dim testHyp As Double = TestsStatistiques.statTestSigne(tabAR, fenetreEstDebut, fenetreEstFin, fenetreEvDebut, fenetreEvFin)
                    'pValeur = 2 * (1 - Globals.ThisAddIn.Application.WorksheetFunction.Norm_S_Dist(Math.Abs(testHyp), True)) * 100
            End Select

            MsgBox("P-Valeur : " & pValeur.ToString("0.0000") & "%")
            'Globals.Ribbons.Ruban.seuilFenetreTaskPane.Visible = False
        End If
        'Catch erreur As InvalidCastException
        'MsgBox("Erreur : Vous devez entrer des données correctes (utiliser la virgule pour les nombres décimaux)", 16)
        'End Try
    End Sub

    Private Sub PValeurFenetre_Click(sender As Object, e As EventArgs) Handles PValeurFenetre.Click
        Dim currentSheet As Excel.Worksheet = CType(Globals.ThisAddIn.Application.Worksheets("Rt"), Excel.Worksheet)
        Dim premiereDate As Integer = currentSheet.Cells(2, 1).Value
        Dim derniereDate As Integer = premiereDate + currentSheet.UsedRange.Rows.Count - 2
        Dim tailleEchant As Integer = currentSheet.UsedRange.Columns.Count - 1

        'Calcul des pvaleurs et affichage de la courbe
        ExcelDialogue.tracerPValeur(tailleEchant, derniereDate)
        'Globals.Ribbons.Ruban.seuilFenetreTaskPane.Visible = False
        Globals.Ribbons.Ruban.graphPVal.Visible = True
    End Sub

End Class
