Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices

Public Class ParamPreTraitPrix

    Private nomFeuille As String
    Private rLog As Boolean
    Private cOuv As Boolean

    'constructeur
    Public Sub New()
        InitializeComponent()
        rLog = False
        cOuv = False
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

        'traitement des prix
        'On centre les cours des entreprises et du marché
        Dim tabPrixCentres(,) As Double = Nothing
        Dim tabMarcheCentre(,) As Double = Nothing
        UtilitaireRentabilites.donneesCentrees(plage, feuilleDates, feuilleDonnees, tabPrixCentres, tabMarcheCentre, cOuv)

        'On calcule les rentabilités
        Dim tabRenta(tabPrixCentres.GetUpperBound(0) - 1, tabPrixCentres.GetUpperBound(1)) As Double
        Dim tabRentaMarche(tabMarcheCentre.GetUpperBound(0) - 1, tabMarcheCentre.GetUpperBound(1)) As Double
        Dim tabRentaClassiquesMarche(tabMarcheCentre.GetUpperBound(0) - 1, tabMarcheCentre.GetUpperBound(1)) As Double
        Dim maxPrixAbsent As Integer
        UtilitaireRentabilites.calculTabRenta(tabPrixCentres, tabMarcheCentre, tabRenta, tabRentaMarche, tabRentaClassiquesMarche, _
                                              maxPrixAbsent, rLog)

        'On stocke le tableaux des rentabilités de marché et des entreprises dont on va avoir besoin
        'PB : où ? Dans nouveau module rentabilité ?
        UtilitaireRentabilites.tabRentaMarche = tabRentaMarche
        UtilitaireRentabilites.tabRenta = tabRenta
        UtilitaireRentabilites.tabRentaClassiquesMarche = tabRentaClassiquesMarche
        'Idem pour maxPrixAbsent
        UtilitaireRentabilites.maxPrixAbs = maxPrixAbsent

        'On affiche ces rentabilités centrées
        ExcelDialogue.affichageRentaCentrees(tabRenta)

    End Sub

    Private Sub nomFeuilleBox_TextChanged(sender As Object, e As EventArgs) Handles nomFeuilleBox.TextChanged
        nomFeuille = nomFeuilleBox.Text
    End Sub

    Private Sub rentaLog_CheckedChanged(sender As Object, e As EventArgs) Handles rentaLog.CheckedChanged
        rLog = rentaLog.Checked
    End Sub

    Private Sub coursOuverture_CheckedChanged(sender As Object, e As EventArgs) Handles coursOuverture.CheckedChanged
        cOuv = coursOuverture.Checked
    End Sub
End Class