Imports Microsoft.Office.Tools.Ribbon
Imports System.Windows.Forms.DataVisualization.Charting

Public Class Ruban

    Public choixSeuil As ChoixSeuil
    Public WithEvents myTaskPane As Microsoft.Office.Tools.CustomTaskPane

    Public choixSeuilFenetre As ChoixSeuilFenetre
    Public WithEvents seuilFenetreTaskPane As Microsoft.Office.Tools.CustomTaskPane

    Public graphPVal As GraphiquePValeur
    Public graphChart As New Chart
    Public graphChartArea As New ChartArea()
    Public valeurs As New Series()

    Private Sub Ruban_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
        choixSeuil = New ChoixSeuil()
        myTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(choixSeuil, "Choix du seuil")
        With myTaskPane
            .DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionFloating
            .Height = 500
            .Width = 500
            .DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight
            .Width = 300
            .Visible = False
        End With

        choixSeuilFenetre = New ChoixSeuilFenetre()
        seuilFenetreTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(choixSeuilFenetre, "Choix des paramètres")
        With seuilFenetreTaskPane
            .DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionFloating
            .Height = 500
            .Width = 500
            .DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight
            .Width = 300
            .Visible = False
        End With

        graphPVal = New GraphiquePValeur()
        graphPVal.Visible = False
        graphChart.ChartAreas.Add(graphChartArea)
        ' Positionner le controle Chart
        graphChart.Location = New System.Drawing.Point(15, 45)

        ' Dimensionner le Chart
        graphChart.Size = New System.Drawing.Size(350, 250)

        Globals.Ribbons.Ruban.valeurs.ChartArea = "ChartArea1"
        Globals.Ribbons.Ruban.graphChart.Series.Add(valeurs)
        ' Ajouter le chart à la form
        graphPVal.Controls.Add(graphChart)

        graphChart.Series("Series1").ChartType = SeriesChartType.Line

    End Sub

    Private Sub AR_Click(ByVal sender As System.Object, _
    ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) _
        Handles AR.Click
        myTaskPane.Visible = True
    End Sub

    Private Sub ModeleMoyenne_Click(ByVal sender As System.Object, _
    ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) _
        Handles ModeleMoyenne.Click
        choixSeuilFenetre.modele = 0
        seuilFenetreTaskPane.Visible = True
    End Sub

    Private Sub ModeleRentaMarche_Click(ByVal sender As System.Object, _
        ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) _
            Handles ModeleRentaMarche.Click
        choixSeuilFenetre.modele = 1
        seuilFenetreTaskPane.Visible = True
    End Sub

    Private Sub ModeleMarche_Click(ByVal sender As System.Object, _
        ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) _
            Handles ModeleMarche.Click
        choixSeuilFenetre.modele = 2
        seuilFenetreTaskPane.Visible = True
    End Sub

End Class
