Imports Microsoft.Office.Tools.Ribbon
Imports System.Windows.Forms.DataVisualization.Charting

Public Class Ruban

    'Public choixSeuil As ChoixSeuil
    'Public WithEvents myTaskPane As Microsoft.Office.Tools.CustomTaskPane

    Public choixSeuilFenetre As ChoixSeuilFenetre
    Public WithEvents seuilFenetreTaskPane As Microsoft.Office.Tools.CustomTaskPane

    Public graphPVal As GraphiquePValeur

    Private Sub Ruban_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

        'Initialisation du taskPane
        choixSeuilFenetre = New ChoixSeuilFenetre(0, 0)
        seuilFenetreTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(choixSeuilFenetre, "Choix des paramètres")
        With seuilFenetreTaskPane
            .DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionFloating
            .Height = 500
            .DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight
            .Width = 300
            .Visible = False
        End With
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

    Private Sub ModeleMoyenne_Click(ByVal sender As System.Object, _
    ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) _
        Handles ModeleMoyenne.Click
        initialisationGraphPVal()
        choixSeuilFenetre.modele = 0
        choixSeuilFenetre.nomModel.Text = "Modèle moyenne des rentabilités"
        seuilFenetreTaskPane.Visible = True
    End Sub

    Private Sub ModeleMarcheSimple_Click(ByVal sender As System.Object, _
        ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) _
            Handles ModeleMarcheSimple.Click
        initialisationGraphPVal()
        choixSeuilFenetre.modele = 1
        choixSeuilFenetre.nomModel.Text = "Modèle de marché simplifié"
        seuilFenetreTaskPane.Visible = True
    End Sub

    Private Sub ModeleMarche_Click(ByVal sender As System.Object, _
        ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) _
            Handles ModeleMarche.Click
        initialisationGraphPVal()
        choixSeuilFenetre.modele = 2
        choixSeuilFenetre.nomModel.Text = "Modèle de marché classique"
        seuilFenetreTaskPane.Visible = True
    End Sub

    Private Sub testSimpleR_Click(sender As Object, e As RibbonControlEventArgs) Handles testSimpleR.Click
        initialisationGraphPVal()
        choixSeuilFenetre.modele = 0
        choixSeuilFenetre.test = 0
        choixSeuilFenetre.nomModel.Text = "Test simple - Modèle moyenne des rentabilités"
        seuilFenetreTaskPane.Visible = True
    End Sub

    Private Sub TestPatellR_Click(sender As Object, e As RibbonControlEventArgs) Handles TestPatellR.Click
        initialisationGraphPVal()
        choixSeuilFenetre.modele = 0
        choixSeuilFenetre.test = 1
        choixSeuilFenetre.nomModel.Text = "Test de Patell - Modèle moyenne des rentabilités"
        seuilFenetreTaskPane.Visible = True
    End Sub

    Private Sub testSigneR_Click(sender As Object, e As RibbonControlEventArgs) Handles testSigneR.Click
        initialisationGraphPVal()
        choixSeuilFenetre.modele = 0
        choixSeuilFenetre.test = 2
        choixSeuilFenetre.nomModel.Text = "Test de signe - Modèle moyenne des rentabilités"
        seuilFenetreTaskPane.Visible = True
    End Sub

    Private Sub testSimpleMS_Click(sender As Object, e As RibbonControlEventArgs) Handles testSimpleMS.Click
        initialisationGraphPVal()
        choixSeuilFenetre.modele = 1
        choixSeuilFenetre.test = 0
        choixSeuilFenetre.nomModel.Text = "Test simple - Modèle de marché simplifié"
        seuilFenetreTaskPane.Visible = True
    End Sub

    Private Sub testPatellMS_Click(sender As Object, e As RibbonControlEventArgs) Handles testPatellMS.Click
        initialisationGraphPVal()
        choixSeuilFenetre.modele = 1
        choixSeuilFenetre.test = 1
        choixSeuilFenetre.nomModel.Text = "Test de Patell - Modèle de marché simplifié"
        seuilFenetreTaskPane.Visible = True
    End Sub

    Private Sub testSigneMS_Click(sender As Object, e As RibbonControlEventArgs) Handles testSigneMS.Click
        initialisationGraphPVal()
        choixSeuilFenetre.modele = 1
        choixSeuilFenetre.test = 2
        choixSeuilFenetre.nomModel.Text = "Test de signe - Modèle de marché simplifié"
        seuilFenetreTaskPane.Visible = True
    End Sub

    Private Sub testSimpleM_Click(sender As Object, e As RibbonControlEventArgs) Handles testSimpleM.Click
        initialisationGraphPVal()
        choixSeuilFenetre.modele = 2
        choixSeuilFenetre.test = 0
        choixSeuilFenetre.nomModel.Text = "Test simple - Modèle de marché classique"
        seuilFenetreTaskPane.Visible = True
    End Sub

    Private Sub testPatellM_Click(sender As Object, e As RibbonControlEventArgs) Handles testPatellM.Click
        initialisationGraphPVal()
        choixSeuilFenetre.modele = 2
        choixSeuilFenetre.test = 1
        choixSeuilFenetre.nomModel.Text = "Test de Patell - Modèle de marché classique"
        seuilFenetreTaskPane.Visible = True
    End Sub

    Private Sub testSigneM_Click(sender As Object, e As RibbonControlEventArgs) Handles testSigneM.Click
        initialisationGraphPVal()
        choixSeuilFenetre.modele = 2
        choixSeuilFenetre.test = 2
        choixSeuilFenetre.nomModel.Text = "Test de signe - Modèle de marché classique"
        seuilFenetreTaskPane.Visible = True
    End Sub
End Class
