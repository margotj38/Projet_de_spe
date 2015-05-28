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
        choixSeuilFenetre = New ChoixSeuilFenetre(0)
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
        seuilFenetreTaskPane.Visible = True
    End Sub

    Private Sub ModeleRentaMarche_Click(ByVal sender As System.Object, _
        ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) _
            Handles ModeleRentaMarche.Click
        initialisationGraphPVal()
        choixSeuilFenetre.modele = 1
        seuilFenetreTaskPane.Visible = True
    End Sub

    Private Sub ModeleMarche_Click(ByVal sender As System.Object, _
        ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) _
            Handles ModeleMarche.Click
        initialisationGraphPVal()
        choixSeuilFenetre.modele = 2
        seuilFenetreTaskPane.Visible = True
    End Sub

End Class
