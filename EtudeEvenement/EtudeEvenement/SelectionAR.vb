Imports System.Diagnostics
Imports System.Windows.Forms.DataVisualization.Charting
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices

Public Class SelectionAR

    Private numTest As Integer ' 0 => TestSimple; 1 => TestPatell; 2 => TestSigne

    'constructeur
    Public Sub New(ByVal valTest As Integer)
        InitializeComponent()
        test = valTest
    End Sub

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

    Private Sub SelectionAR_Load(sender As Object, e As EventArgs) Handles Me.Load
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
    End Sub


    Private Sub LancementEtEv_Click(sender As Object, e As EventArgs) Handles LancementEtEv.Click

        'On récupère les plages des périodes d'estimation et d'événement + la feuille sur laquelle elles sont
        'Les plages ont pour premiere colonne les dates
        Dim plageEst As String = ""
        Dim plageEv As String = ""
        Dim feuille As String = ""
        Utilitaires.recupererFeuillePlage(Me.refEditEst.Address, feuille, plageEst)
        Utilitaires.recupererFeuillePlage(Me.refEditEv.Address, feuille, plageEv)

        Select Case test
            Case 0
                'test simple
                Dim tabEstAR(,) As Double = Nothing
                Dim tabEvAR(,) As Double = Nothing
                Dim tabDateEst() As Integer = Nothing
                Dim tabDateEv() As Integer = Nothing
                ExcelDialogue.convertPlageTab(plageEst, feuille, tabEstAR, tabDateEst)
                ExcelDialogue.convertPlageTab(plageEv, feuille, tabEvAR, tabDateEv)
                ExcelDialogue.traitementPlageAR(plageEst, plageEv, feuille)
            Case 2
                'test de signe
                ExcelDialogue.affichageSigne(tabDateEv:=, tabEstAR:=, tabEvAR)
        End Select

    End Sub



End Class