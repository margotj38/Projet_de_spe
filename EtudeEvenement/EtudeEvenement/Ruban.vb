Imports Microsoft.Office.Tools.Ribbon

Public Class Ruban

    Private choixSeuil As ChoixSeuil
    Public WithEvents myTaskPane As Microsoft.Office.Tools.CustomTaskPane

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
    End Sub

    Private Sub AR_Click(ByVal sender As System.Object, _
    ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) _
        Handles AR.Click
        myTaskPane.Visible = True
    End Sub

    Private Sub ModeleMoyenne_Click(ByVal sender As System.Object, _
    ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) _
        Handles ModeleMoyenne.Click

    End Sub

    Private Sub ModeleRentaMarche_Click(ByVal sender As System.Object, _
        ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) _
            Handles ModeleRentaMarche.Click

    End Sub

    Private Sub ModeleMarche_Click(ByVal sender As System.Object, _
        ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) _
            Handles ModeleMarche.Click

    End Sub

End Class
