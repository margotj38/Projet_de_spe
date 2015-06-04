Public Class ParamAR

    Private textPlageEst As String
    Private textPlageEv As String

    'constructeur
    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub plageEstBox_TextChanged(sender As Object, e As EventArgs) Handles plageEstBox.TextChanged
        textPlageEst = plageEstBox.Text
    End Sub

    Private Sub plageEvBox_TextChanged(sender As Object, e As EventArgs) Handles plageEvBox.TextChanged
        textPlageEv = plageEvBox.Text
    End Sub

    Private Sub LancementEtEv_Click(sender As Object, e As EventArgs) Handles LancementEtEv.Click

        Dim plageEst As String = textPlageEst
        Dim plageEv As String = textPlageEv

        Dim currentSheet As Excel.Worksheet = CType(Globals.ThisAddIn.Application.Worksheets("AR"), Excel.Worksheet)
        Dim premiereDate As Integer = currentSheet.Cells(2, 1).Value
        Dim tailleEchant As Integer = currentSheet.UsedRange.Columns.Count - 1

        'traitement des données AR fournies
        ExcelDialogue.traitementAR(plageEst, plageEv)
        Globals.Ribbons.Ruban.paramARTaskPane.Visible = False
    End Sub


End Class
