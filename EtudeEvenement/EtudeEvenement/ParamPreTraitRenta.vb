Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices

Public Class ParamPreTraitRenta

    Private nomFeuille As String

    'constructeur
    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub SelectionDatesEv_Load(sender As Object, e As EventArgs) Handles Me.Load
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

            Me.datesEvRefEdit.ExcelConnector = excelApp
        End Try

        Me.datesEvRefEdit.Focus()
    End Sub

    Private Sub lancementPreT_Click(sender As Object, e As EventArgs) Handles lancementPreT.Click
        'On récupère la plage des dates et la feuille sur laquelle elle est
        Dim plage As String = ""
        Dim feuilleDonnees As String = Me.nomFeuille
        Dim feuilleDates As String = ""
        Utilitaires.recupererFeuillePlage(Me.datesEvRefEdit.Address, feuilleDates, plage)

        'Traitement des rentabilités
        'On centre les rentabilités (2ème colonne : marché)
        Dim tabRentaCentrees(,) As Double = Nothing
        Dim tabMarcheCentre(,) As Double = Nothing

        UtilitaireRentabilites.donneesCentrees(plage, feuilleDates, feuilleDonnees, tabRentaCentrees, tabMarcheCentre, False)

        'On stocke le tableaux des rentabilités de marché dont on va avoir besoin
        'PB : où ? Dans nouveau module rentabilité ?
        UtilitaireRentabilites.tabRentaMarche = tabMarcheCentre
        UtilitaireRentabilites.tabRenta = tabRentaCentrees
        UtilitaireRentabilites.tabRentaClassiquesMarche = tabMarcheCentre

        'Calcul de maxPrixAbsent
        Dim maxPrixAbsent As Integer = UtilitaireRentabilites.calculMaxPrixAbs(tabRentaCentrees)
        UtilitaireRentabilites.maxPrixAbs = maxPrixAbsent

        'On affiche les rentabilités centrées
        ExcelDialogue.affichageRentaCentrees(tabRentaCentrees)

    End Sub

    Private Sub nomFeuilleBox_TextChanged(sender As Object, e As EventArgs) Handles nomFeuilleBox.TextChanged
        nomFeuille = nomFeuilleBox.Text
    End Sub
End Class