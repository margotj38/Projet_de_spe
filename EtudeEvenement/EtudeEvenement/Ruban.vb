Imports Microsoft.Office.Tools.Ribbon
Imports System.Windows.Forms.DataVisualization.Charting
Imports System.Runtime.InteropServices
Imports System.Net.Mime.MediaTypeNames
Imports Microsoft.Office.Interop

Public Class Ruban

    Public selFenetres As SelectionFenetres
    Public selAR As SelectionAR
    Public paramRenta As ParamPreTraitRenta
    Public paramPrix As ParamPreTraitPrix

    Public graphPVal As GraphiquePValeur

    Private Sub Ruban_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub initialisationGraphPVal()
        graphPVal = New GraphiquePValeur()
        graphPVal.Visible = False

        graphPVal.GraphiqueChart.Series.Add("Series1")

        graphPVal.GraphiqueChart.Series("Series1").ChartType = SeriesChartType.Line
        graphPVal.GraphiqueChart.Series("Series1").IsValueShownAsLabel = True
        graphPVal.GraphiqueChart.ChartAreas(0).AxisX.MajorGrid.Enabled = False
        graphPVal.GraphiqueChart.ChartAreas(0).AxisY.MajorGrid.Enabled = False
    End Sub

    Private Sub initSelFenetres(selFenetres As SelectionFenetres)
        'Définition des options des refEdit
        selFenetres.refEditEst.IncludeSheetName = True
        selFenetres.refEditEst.ShowRowAbsoluteIndicator = False
        selFenetres.refEditEst.ShowColumnAbsoluteIndicator = False
        selFenetres.refEditEv.IncludeSheetName = True
        selFenetres.refEditEv.ShowRowAbsoluteIndicator = False
        selFenetres.refEditEv.ShowColumnAbsoluteIndicator = False

        'selFenetres.TopMost = True
    End Sub

    Private Sub testSimpleR_Click(sender As Object, e As RibbonControlEventArgs) Handles testSimpleR.Click
        initialisationGraphPVal()
        selFenetres = New SelectionFenetres(0, 0)
        initSelFenetres(selFenetres)
        'selFenetres.modele = 0
        'selFenetres.test = 0
        selFenetres.nomModele.Text = "Test simple - Modèle moyenne des rentabilités"
        'seuilFenetreTaskPane.Visible = True
        selFenetres.Visible = True
    End Sub

    Private Sub TestPatellR_Click(sender As Object, e As RibbonControlEventArgs) Handles TestPatellR.Click
        initialisationGraphPVal()
        selFenetres = New SelectionFenetres(0, 1)
        initSelFenetres(selFenetres)
        'selFenetres.modele = 0
        'selFenetres.test = 1
        selFenetres.nomModele.Text = "Test de Patell - Modèle moyenne des rentabilités"
        'seuilFenetreTaskPane.Visible = True
        selFenetres.Visible = True
    End Sub

    Private Sub testSigneR_Click(sender As Object, e As RibbonControlEventArgs) Handles testSigneR.Click
        initialisationGraphPVal()
        selFenetres = New SelectionFenetres(0, 2)
        initSelFenetres(selFenetres)
        'selFenetres.modele = 0
        'selFenetres.test = 2
        selFenetres.nomModele.Text = "Test de signe - Modèle moyenne des rentabilités"
        'seuilFenetreTaskPane.Visible = True
        selFenetres.Visible = True
    End Sub

    Private Sub testSimpleMS_Click(sender As Object, e As RibbonControlEventArgs) Handles testSimpleMS.Click
        initialisationGraphPVal()
        selFenetres = New SelectionFenetres(1, 0)
        initSelFenetres(selFenetres)
        'selFenetres.modele = 1
        'selFenetres.test = 0
        selFenetres.nomModele.Text = "Test simple - Modèle de marché simplifié"
        'seuilFenetreTaskPane.Visible = True
        selFenetres.Visible = True
    End Sub

    Private Sub testPatellMS_Click(sender As Object, e As RibbonControlEventArgs) Handles testPatellMS.Click
        initialisationGraphPVal()
        selFenetres = New SelectionFenetres(1, 1)
        initSelFenetres(selFenetres)
        'selFenetres.modele = 1
        'selFenetres.test = 1
        selFenetres.nomModele.Text = "Test de Patell - Modèle de marché simplifié"
        'seuilFenetreTaskPane.Visible = True
        selFenetres.Visible = True
    End Sub

    Private Sub testSigneMS_Click(sender As Object, e As RibbonControlEventArgs) Handles testSigneMS.Click
        initialisationGraphPVal()
        selFenetres = New SelectionFenetres(1, 2)
        initSelFenetres(selFenetres)
        'selFenetres.modele = 1
        'selFenetres.test = 2
        selFenetres.nomModele.Text = "Test de signe - Modèle de marché simplifié"
        'seuilFenetreTaskPane.Visible = True
        selFenetres.Visible = True
    End Sub

    Private Sub testSimpleM_Click(sender As Object, e As RibbonControlEventArgs) Handles testSimpleM.Click
        initialisationGraphPVal()
        selFenetres = New SelectionFenetres(2, 0)
        initSelFenetres(selFenetres)
        'selFenetres.modele = 2
        'selFenetres.test = 0
        selFenetres.nomModele.Text = "Test simple - Modèle de marché classique"
        'seuilFenetreTaskPane.Visible = True
        selFenetres.Visible = True
    End Sub

    Private Sub testPatellM_Click(sender As Object, e As RibbonControlEventArgs) Handles testPatellM.Click
        initialisationGraphPVal()
        selFenetres = New SelectionFenetres(2, 1)
        initSelFenetres(selFenetres)
        'selFenetres.modele = 2
        'selFenetres.test = 1
        selFenetres.nomModele.Text = "Test de Patell - Modèle de marché classique"
        'seuilFenetreTaskPane.Visible = True
        selFenetres.Visible = True
    End Sub

    Private Sub testSigneM_Click(sender As Object, e As RibbonControlEventArgs) Handles testSigneM.Click
        initialisationGraphPVal()
        selFenetres = New SelectionFenetres(2, 2)
        initSelFenetres(selFenetres)
        selFenetres.nomModele.Text = "Test de signe - Modèle de marché classique"
        selFenetres.Visible = True
    End Sub

    Private Sub preTraitPrix_Click(sender As Object, e As RibbonControlEventArgs) Handles preTraitPrix.Click

        paramPrix = New ParamPreTraitPrix()

        'Définition des options de refEdit
        paramPrix.datesEvRefEdit.IncludeSheetName = True
        paramPrix.datesEvRefEdit.ShowRowAbsoluteIndicator = False
        paramPrix.datesEvRefEdit.ShowColumnAbsoluteIndicator = False
        paramPrix.TopMost = True
        paramPrix.Visible = True
        paramPrix.nomFeuilleBox.Text = "Prix"
    End Sub

    Private Sub preTraitRenta_Click(sender As Object, e As RibbonControlEventArgs) Handles preTraitRenta.Click

        paramRenta = New ParamPreTraitRenta()

        'Définition des options de refEdit
        paramRenta.datesEvRefEdit.IncludeSheetName = True
        paramRenta.datesEvRefEdit.ShowRowAbsoluteIndicator = False
        paramRenta.datesEvRefEdit.ShowColumnAbsoluteIndicator = False
        paramRenta.TopMost = True
        paramRenta.Visible = True
        paramRenta.nomFeuilleBox.Text = "Rent"
    End Sub

    Private Sub testSimplAR_Click(sender As Object, e As RibbonControlEventArgs) Handles testSimplAR.Click
        selAR = New SelectionAR(0)

        'Définition des options des refEdit
        selAR.refEditEst.IncludeSheetName = True
        selAR.refEditEst.ShowRowAbsoluteIndicator = False
        selAR.refEditEst.ShowColumnAbsoluteIndicator = False
        selAR.refEditEv.IncludeSheetName = True
        selAR.refEditEv.ShowRowAbsoluteIndicator = False
        selAR.refEditEv.ShowColumnAbsoluteIndicator = False
        'selAR.TopMost = True

        selAR.Visible = True
    End Sub

    Private Sub testSignAR_Click(sender As Object, e As RibbonControlEventArgs) Handles testSignAR.Click
        selAR = New SelectionAR(1)

        'Définition des options des refEdit
        selAR.refEditEst.IncludeSheetName = True
        selAR.refEditEst.ShowRowAbsoluteIndicator = False
        selAR.refEditEst.ShowColumnAbsoluteIndicator = False
        selAR.refEditEv.IncludeSheetName = True
        selAR.refEditEv.ShowRowAbsoluteIndicator = False
        selAR.refEditEv.ShowColumnAbsoluteIndicator = False
        'selAR.TopMost = True

        selAR.Visible = True
    End Sub

End Class
