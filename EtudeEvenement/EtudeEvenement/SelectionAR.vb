Imports System.Diagnostics
Imports System.Windows.Forms.DataVisualization.Charting
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices

Public Class SelectionAR

    'constructeur
    Public Sub New()
        InitializeComponent()
    End Sub

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
            'Me.refEditEv.ExcelConnector = excelApp
        End Try

        Me.refEditEst.Focus()
    End Sub


    Private Sub LancementEtEv_Click(sender As Object, e As EventArgs) Handles LancementEtEv.Click

        'On récupère les plages des périodes d'estimation et d'événement + la feuille sur laquelle elles sont
        'Les plages ont pour premiere colonne les dates
        Dim plageEst As String = ""
        Dim plageEv As String = "A20:F30"
        Dim feuille As String = ""
        Utilitaires.recupererFeuillePlage(Me.refEditEst.Address, feuille, plageEst)
        'Utilitaires.recupererFeuillePlage(Me.refEditEv.Address, feuille, plageEv)

        'Traitement des données AR fournies
        ExcelDialogue.traitementAR(plageEst, plageEv, feuille)

        '
        'Me.Visible = False

    End Sub



End Class