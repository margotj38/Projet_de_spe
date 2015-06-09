Imports System.Diagnostics
Imports System.Windows.Forms.DataVisualization.Charting
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices

''' <summary>
''' Windows Form permettant la gestion du lancement de l'étude d'événement à partir de AR.
''' </summary>
''' <remarks>L'interface graphique de cette Windows Form contient deux objets de type refEdit (un pour sélectionner la 
''' plage correspondant à la période d'estimation, un second pour la sélection de la période d'événement) et un bouton 
''' pour lancer le test choisi à partir des AR sélectionnés.</remarks>
Public Class SelectionAR

    ''' <summary>
    ''' Numéro du test à utiliser (0 => test de Student; 1 => test de signe).
    ''' </summary>
    ''' <remarks></remarks>
    Private numTest As Integer

    ''' <summary>
    ''' Constructeur initialisant le numéro du test au test choisi.
    ''' </summary>
    ''' <param name="valTest">Numéro du test à utiliser.</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal valTest As Integer)
        InitializeComponent()
        test = valTest
    End Sub

    ''' <summary>
    ''' Accesseur sur le numéro du test.
    ''' </summary>
    ''' <value>Numéro du test choisi</value>
    ''' <returns>Le numéro du test</returns>
    ''' <remarks></remarks>
    Public Property test() As Integer
        Get
            Return numTest
        End Get
        Set(value As Integer)
            If value < 0 Or value > 1 Then
                MsgBox("Erreur interne : numéro de test incorrect", 16)
            End If
            numTest = value
        End Set
    End Property

    ''' <summary>
    ''' Méthode appelée au chargement de cette classe. Elle connecte les objets refEdit à l'application Excel en cours
    ''' d'exécution.
    ''' </summary>
    ''' <param name="sender">Non utilisé</param>
    ''' <param name="e">Non utilisé</param>
    ''' <remarks></remarks>
    Private Sub SelectionAR_Load(sender As Object, e As EventArgs) Handles Me.Load
        Dim excelApp As Excel.Application = Nothing

        'Create an Excel App
        Try
            excelApp = Marshal.GetActiveObject("Excel.Application")
        Catch ex As COMException
            'An exception is thrown if there is not an open excel instance.                    
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
    End Sub

    ''' <summary>
    ''' Méthode associée au clic sur le bouton "Lancement". Elle récupère les AR, effectue le test choisi et affiche 
    ''' les résultats.
    ''' </summary>
    ''' <param name="sender">Non utilisé</param>
    ''' <param name="e">Non utilisé</param>
    ''' <remarks></remarks>
    Private Sub LancementEtEv_Click(sender As Object, e As EventArgs) Handles LancementEtEv.Click

        'On récupère les plages des périodes d'estimation et d'événement + la feuille sur laquelle elles sont
        'Les plages ont pour premiere colonne les dates
        Dim plageEst As String = ""
        Dim plageEv As String = ""
        Dim feuille As String = ""
        'Si une des deux plages n'a pas été sélectionnée, on lève une erreur
        If Me.refEditEst.Address = "" Or Me.refEditEv.Address = "" Then
            MsgBox("Erreur : Une plage de données n'a pas été renseignée", 16)
            Return
        End If
        Utilitaires.recupererFeuillePlage(Me.refEditEst.Address, feuille, plageEst)
        Utilitaires.recupererFeuillePlage(Me.refEditEv.Address, feuille, plageEv)

        'On compte le nombre de colonnes de chaque plage
        Dim premColEst As Integer = 0, dernColEst As Integer = 0
        Utilitaires.parserPlageColonnes(plageEst, premColEst, dernColEst)
        Dim premColEv As Integer = 0, dernColEv As Integer = 0
        Utilitaires.parserPlageColonnes(plageEv, premColEv, dernColEv)
        'Si le nombre de colonnes n'est pas le même, on lève une erreur
        If Not dernColEst - premColEst = dernColEv - premColEv Then
            MsgBox("Erreur : Le nombre de colonnes sélectionnées pour la période d'estimation n'est pas le même que " &
                   "pour la période d'événement", 16)
            Return
        End If

        'On récupère les données dans des tableaux adaptés
        Dim tabEstAR(,) As Double = Nothing
        Dim tabEvAR(,) As Double = Nothing
        Dim tabDateEst() As Integer = Nothing
        Dim tabDateEv() As Integer = Nothing
        ExcelDialogue.convertPlageTab(plageEst, feuille, tabEstAR, tabDateEst)
        ExcelDialogue.convertPlageTab(plageEv, feuille, tabEvAR, tabDateEv)

        Select Case test
            Case 0
                'test simple
                RentaAnormales.traitementTabAR(tabEvAR, tabEstAR, tabDateEv)
            Case 1
                'test de signe
                ExcelDialogue.affichageSigne(tabDateEv, tabEstAR, tabEvAR)
        End Select

    End Sub

    Private Sub refEditEst_Clicked(sender As Object, e As LeafCreations.EventArgs.BeforeResizeEventArgs) Handles refEditEst.Clicked
        MsgBox("Affichons qqch")
    End Sub
End Class