Imports Microsoft.Office.Tools.Ribbon

Public Class Ruban

    Dim actionsPane1 As New UserControl1
    Dim actionsPane2 As New UserControl2

    Private Sub MyRibbon_Load(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonUIEventArgs) Handles MyBase.Load
        Globals.ThisAddIn.CustomTaskPanes.Add(actionsPane1, "BBB1")
        Globals.ThisAddIn.CustomTaskPanes.Add(actionsPane2, "BBB2")
        'Globals.ThisWorkbook.ActionsPane.Controls.Add(actionsPane1)
        'Globals.ThisWorkbook.ActionsPane.Controls.Add(actionsPane2)
        actionsPane1.Hide()
        actionsPane2.Hide()
        '
        'Globals.ThisAddIn.Application.DisplayDocumentActionTaskPane() = False
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, _
    ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) _
        Handles Button1.Click
        'Globals.ThisWorkbook.Application.DisplayDocumentActionTaskPane = True
        'Globals.ThisAddIn.Application.DisplayDocumentActionTaskPane = True
        Globals.ThisAddIn.ThisAddIn_Methode_CAR()
        actionsPane2.Hide()
        actionsPane1.Show()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, _
        ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) _
            Handles Button2.Click
        '
        'Globals.ThisAddIn.Application.DisplayDocumentActionTaskPane = True
        MsgBox("Bouton 2 fonctionne")
        actionsPane1.Hide()
        actionsPane2.Show()
    End Sub


End Class
