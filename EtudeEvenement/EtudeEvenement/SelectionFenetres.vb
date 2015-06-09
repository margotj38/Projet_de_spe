﻿Imports System.Diagnostics
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
        'Les plages ont pour premiere colonne les dates
        Dim plageEst As String = ""
        Dim plageEv As String = ""
        Dim feuille As String = ""
        Utilitaires.recupererFeuillePlage(Me.refEditEst.Address, feuille, plageEst)
        Utilitaires.recupererFeuillePlage(Me.refEditEv.Address, feuille, plageEv)

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