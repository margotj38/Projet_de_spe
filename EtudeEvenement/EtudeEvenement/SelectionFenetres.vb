Imports System.Diagnostics
Imports System.Windows.Forms.DataVisualization.Charting
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices

Public Class SelectionFenetres

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

    Private Sub SelectionFenetres_Load(sender As Object, e As EventArgs) Handles Me.Load
        Dim excelApp As Excel.Application = Nothing

        ' Create an Excel App
        Try
            excelApp = Marshal.GetActiveObject("Excel.Application")
        Catch ex As COMException
            ' An exception is thrown if there is not an open excel instance.                    
        Finally
            If excelApp Is Nothing Then
                excelApp = New Microsoft.Office.Interop.Excel.Application
                excelApp.Workbooks.Add()
            End If
            excelApp.Visible = True

            Me.refEditEst.ExcelConnector = excelApp
            Me.refEditEv.ExcelConnector = excelApp
        End Try

        Me.refEditEst.Focus()
        Me.refEditEv.Focus()
    End Sub

    Private Sub LancementEtEv_Click(sender As Object, e As EventArgs) Handles LancementEtEv.Click

        'On récupère les plages des périodes d'estimation et d'événement + la feuille sur laquelle elles sont
        Dim plageEst As String = ""
        Dim plageEv As String = ""
        Dim feuille As String = ""
        Utilitaires.recupererFeuillePlage(Me.refEditEst.Address, feuille, plageEst)
        Utilitaires.recupererFeuillePlage(Me.refEditEst.Address, feuille, plageEv)

        'On construit les 4 tableaux des rentabilités (entreprises et marché, période d'estimation et d'événement)
        Dim currentSheet As Excel.Worksheet = CType(Globals.ThisAddIn.Application.Worksheets(feuille), Excel.Worksheet)
        Dim tabRentaEst(,) As Double
        Dim tabRentaEv(,) As Double
        Dim tabRentaMarcheEst(,) As Double
        Dim tabRentaMarcheEv(,) As Double
        UtilitaireRentabilites.constructionTabRenta(plageEst, plageEv, feuille, UtilitaireRentabilites.tabRentaMarche, _
                                                    tabRentaMarcheEst, tabRentaMarcheEv, tabRentaEst, tabRentaEv)

        'Calcul des AR
        Dim tabAREst(,) As Double
        Dim tabAREv(,) As Double

        RentaAnormales.calculAR(tabRentaMarcheEst, tabRentaMarcheEv, tabRentaEst, tabRentaEv, tabAREst, tabAREv)

        'Dim pValeur As Double
        'Select Case test
        '    Case 0
        '        'test simple'
        '        Dim tabCAR As Double()
        '        tabCAR = TestsStatistiques.calculCAR(tabAR, premiereDate + 1, fenetreEstDebut, fenetreEstFin, fenetreEvDebut, fenetreEvFin)
        '        Dim testHyp As Double = TestsStatistiques.calculStatStudent(tabCAR)
        '        pValeur = TestsStatistiques.calculPValeur(tailleEchant, testHyp) * 100
        '    Case 1
        '        'test de Patell'
        '        Dim testHyp As Double = TestsStatistiques.patellTest(tabAR, fenetreEstDebut, fenetreEstFin, fenetreEvDebut, fenetreEvFin)
        '        pValeur = 2 * (1 - Globals.ThisAddIn.Application.WorksheetFunction.Norm_S_Dist(Math.Abs(testHyp), True)) * 100
        '    Case 2
        '        'test de signe'
        '        Dim testHyp As Double = TestsStatistiques.statTestSigne(tabAR, fenetreEstDebut, fenetreEstFin, fenetreEvDebut, fenetreEvFin)
        '        pValeur = 2 * (1 - Globals.ThisAddIn.Application.WorksheetFunction.Norm_S_Dist(Math.Abs(testHyp), True)) * 100
        'End Select

        'MsgBox("P-Valeur : " & pValeur.ToString("0.0000") & "%")
        'Globals.Ribbons.Ruban.seuilFenetreTaskPane.Visible = False
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