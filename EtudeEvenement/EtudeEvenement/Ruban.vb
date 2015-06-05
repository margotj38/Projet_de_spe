Imports Microsoft.Office.Tools.Ribbon
Imports System.Windows.Forms.DataVisualization.Charting
Imports System.Runtime.InteropServices
Imports System.Net.Mime.MediaTypeNames
Imports Microsoft.Office.Interop

Public Class Ruban

    'Public choixSeuilFenetre As ChoixSeuilFenetre
    Public selFenetres As SelectionFenetres
    Public paramAR As ParamAR
    'Public choixDatesEv As ChoixDatesEv
    Public selDatesEv As SelectionDatesEv
    'Public WithEvents seuilFenetreTaskPane As Microsoft.Office.Tools.CustomTaskPane
    Public WithEvents paramARTaskPane As Microsoft.Office.Tools.CustomTaskPane
    'Public WithEvents datesEvTaskPane As Microsoft.Office.Tools.CustomTaskPane


    Public graphPVal As GraphiquePValeur

    Private Sub Ruban_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

        'Initialisation du taskPane choixSeuilFenetre
        'choixSeuilFenetre = New ChoixSeuilFenetre(0, 0)
        'seuilFenetreTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(choixSeuilFenetre, "Choix des paramètres")
        'With seuilFenetreTaskPane
        '    .DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionFloating
        '    .Height = 500
        '    .DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight
        '    .Width = 300
        '    .Visible = False
        'End With

        'Initialisation du taskPane paramAR
        paramAR = New ParamAR()
        paramARTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(paramAR, "Choix des paramètres")
        With paramARTaskPane
            .DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionFloating
            .Height = 500
            .DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight
            .Width = 300
            .Visible = False
        End With
        'Initialisation du taskPane choixDatesEv
        'choixDatesEv = New ChoixDatesEv()
        'datesEvTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(choixDatesEv, "Choix des paramètres")
        'With datesEvTaskPane
        '    .DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionFloating
        '    .Height = 500
        '    .DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight
        '    .Width = 300
        '    .Visible = False
        'End With
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
        'selFenetres.modele = 2
        'selFenetres.test = 2
        selFenetres.nomModele.Text = "Test de signe - Modèle de marché classique"
        'seuilFenetreTaskPane.Visible = True
        selFenetres.Visible = True
    End Sub

    Private Sub ARparam_Click(sender As Object, e As RibbonControlEventArgs) Handles ARparam.Click

        'Définition des options de refEdit
        paramAR.estimation.IncludeSheetName = False
        paramAR.estimation.ShowRowAbsoluteIndicator = False
        paramAR.estimation.ShowColumnAbsoluteIndicator = False
        paramAR.evenement.IncludeSheetName = False
        paramAR.evenement.ShowRowAbsoluteIndicator = False
        paramAR.evenement.ShowColumnAbsoluteIndicator = False

        Dim excelApp As Excel.Application = Nothing

        ' Create an Excel App
        Try
            excelApp = Marshal.GetActiveObject("Excel.Application")
        Catch ex As COMException
            ' An exception is thrown if there is not an open excel instance.                    
        Finally
            If excelApp Is Nothing Then
                'excelApp = New Application
                excelApp.Workbooks.Add()
            End If
            excelApp.Visible = True

            paramAR.estimation.ExcelConnector = excelApp
            paramAR.evenement.ExcelConnector = excelApp

        End Try

        paramAR.estimation.Focus()
        paramAR.evenement.Focus()

        paramARTaskPane.Visible = True
    End Sub

    Private Sub preTraitPrix_Click(sender As Object, e As RibbonControlEventArgs) Handles preTraitPrix.Click

        selDatesEv = New SelectionDatesEv(0)

        'Définition des options de refEdit
        selDatesEv.datesEvRefEdit.IncludeSheetName = True
        selDatesEv.datesEvRefEdit.ShowRowAbsoluteIndicator = False
        selDatesEv.datesEvRefEdit.ShowColumnAbsoluteIndicator = False
        selDatesEv.TopMost = True
        selDatesEv.Visible = True
        selDatesEv.nomFeuilleBox.Text = "Prix"
    End Sub

    Private Sub preTraitRenta_Click(sender As Object, e As RibbonControlEventArgs) Handles preTraitRenta.Click
        selDatesEv = New SelectionDatesEv(1)

        'Définition des options de refEdit
        selDatesEv.datesEvRefEdit.IncludeSheetName = True
        selDatesEv.datesEvRefEdit.ShowRowAbsoluteIndicator = False
        selDatesEv.datesEvRefEdit.ShowColumnAbsoluteIndicator = False
        selDatesEv.TopMost = True
        selDatesEv.Visible = True
        selDatesEv.nomFeuilleBox.Text = "Rent"
    End Sub

    Private Sub BoutonTest_Click(sender As Object, e As RibbonControlEventArgs) Handles BoutonTest.Click

    End Sub
End Class
