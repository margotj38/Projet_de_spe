Imports System.Diagnostics
Imports System.Windows.Forms.DataVisualization.Charting
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices

''' <summary>
''' Windows Form permettant la gestion du lancement de l'étude d'événement à partir des rentabilités centrées.
''' </summary>
''' <remarks>L'interface graphique de cette Windows Form contient deux objets de type refEdit (un pour sélectionner la 
''' plage correspondant à la période d'estimation, un second pour la sélection de la période d'événement) et un bouton 
''' pour lancer l'étude d'événement à partir des rentabilités choisies.</remarks>
Public Class SelectionFenetres

    ''' <summary>
    ''' Numéro du modèle à utiliser (0 => ModeleMoyenne; 1 => ModeleMarcheSimple; 2 => ModeleMarche).
    ''' </summary>
    ''' <remarks></remarks>
    Private model As Integer

    ''' <summary>
    ''' Numéro du test à utiliser (0 => test de Student; 1 => test de Patell; 2 => test de signe).
    ''' </summary>
    ''' <remarks></remarks>
    Private numTest As Integer

    ''' <summary>
    ''' Constructeur initialisant les numéro du modèle et du test.
    ''' </summary>
    ''' <param name="model">Numéro de modèle à utiliser pour l'étude d'événement.</param>
    ''' <param name="test">Numéro de test à utiliser pour l'étude d'événement.</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal model As Integer, ByVal test As Integer)
        InitializeComponent()
        Me.model = model
        Me.numTest = test
    End Sub

    ''' <summary>
    ''' Accesseur sur le numéro du modèle.
    ''' </summary>
    ''' <value>Numéro du modèle choisi</value>
    ''' <returns>Le numéro du modèle</returns>
    ''' <remarks></remarks>
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
            If value < 0 Or value > 2 Then
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
    Private Sub SelectionFenetres_Load(sender As Object, e As EventArgs) Handles Me.Load
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
        Me.refEditEv.Focus()
    End Sub

    ''' <summary>
    ''' Méthode associée au clic sur le bouton "Lancement". Elle calcule les AR selon le modèle choisi, les affiche dans une
    ''' nouvelle feuille, effectue le test statistique choisi, puis affiche les résultats finaux dans une autre feuille
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

        'On construit les 4 tableaux des rentabilités (entreprises et marché, période d'estimation et d'événement)
        Dim currentSheet As Excel.Worksheet = CType(Globals.ThisAddIn.Application.Worksheets(feuille), Excel.Worksheet)
        Dim tabRentaEst(,) As Double = Nothing
        Dim tabRentaEv(,) As Double = Nothing
        Dim tabRentaMarcheEst(,) As Double = Nothing
        Dim tabRentaMarcheEv(,) As Double = Nothing
        Dim tabRentaClassiquesMarcheEst(,) As Double = Nothing
        Dim tabRentaClassiquesMarcheEv(,) As Double = Nothing
        UtilitaireRentabilites.constructionTabRenta(plageEst, plageEv, _
                                                    UtilitaireRentabilites.tabRentaMarche, UtilitaireRentabilites.tabRenta, _
                                                    UtilitaireRentabilites.tabRentaClassiquesMarche, _
                                                    tabRentaMarcheEst, tabRentaMarcheEv, tabRentaEst, tabRentaEv, _
                                                    tabRentaClassiquesMarcheEst, tabRentaClassiquesMarcheEv)
        'Calcul des AR
        Dim tabAREst(,) As Double = Nothing
        Dim tabAREv(,) As Double = Nothing
        Dim tabDateEst() As Integer = Nothing
        Dim tabDateEv() As Integer = Nothing
        RentaAnormales.calculAR(tabRentaMarcheEst, tabRentaMarcheEv, tabRentaEst, tabRentaEv, _
                                tabAREst, tabAREv, tabDateEst, tabDateEv)

        'Affichage des AR dans une nouvelle feuille excel
        ExcelDialogue.affichageAR(tabAREst, tabAREv, tabDateEst, tabDateEv)

        'On appelle les différents tests
        Select Case test
            Case 0
                'test simple
                'Calcule et affiches les résultats des test AR et CAR
                RentaAnormales.traitementTabAR(tabAREv, tabAREst, tabDateEv)
            Case 1
                'test de Patell
                'Calcul du nombre de AR non manquants pour chaque entreprise sur la période d'estimation
                Dim nbNonMissingReturn() As Integer = TestsStatistiques.calculNbNonMissingReturn(tabAREst)
                Dim testHypAAR() As Double = Nothing
                Dim testHypCAAR() As Double = Nothing
                TestsStatistiques.patellTest(tabAREst, tabAREv, tabDateEst, tabDateEv, tabRentaClassiquesMarcheEst, _
                                             tabRentaClassiquesMarcheEv, nbNonMissingReturn, testHypAAR, testHypCAAR)
                ExcelDialogue.affichagePatell(tabDateEv, testHypAAR, testHypCAAR)
            Case 2
                'test de signe
                ExcelDialogue.affichageSigne(tabDateEv, tabAREst, tabAREv)
        End Select
    End Sub

End Class