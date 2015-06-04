Public Class ParamAR

    Private textFenetreEstDebut As String
    Private textFenetreEstFin As String
    Private textFenetreEvDebut As String
    Private textFenetreEvFin As String

    'constructeur
    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub FenetreEstDebBox_TextChanged(sender As Object, e As EventArgs) Handles FenetreEstDebBox.TextChanged
        textFenetreEstDebut = FenetreEstDebBox.Text
    End Sub

    Private Sub FenetreEstFinBox_TextChanged(sender As Object, e As EventArgs) Handles FenetreEstFinBox.TextChanged
        textFenetreEstFin = FenetreEstFinBox.Text
    End Sub

    Private Sub FenetreDebBox_TextChanged(sender As Object, e As EventArgs) Handles FenetreDebBox.TextChanged
        textFenetreEvDebut = FenetreDebBox.Text
    End Sub

    Private Sub FenetreFinBox_TextChanged(sender As Object, e As EventArgs) Handles FenetreFinBox.TextChanged
        textFenetreEvFin = FenetreFinBox.Text
    End Sub

    Private Sub LancementEtEv_Click(sender As Object, e As EventArgs) Handles LancementEtEv.Click

        Dim fenetreEstDebut As Integer = CInt(textFenetreEstDebut)
        Dim fenetreEstFin As Integer = CInt(textFenetreEstFin)
        Dim fenetreEvDebut As Integer = CInt(textFenetreEvDebut)
        Dim fenetreEvFin As Integer = CInt(textFenetreEvFin)

        Dim currentSheet As Excel.Worksheet = CType(Globals.ThisAddIn.Application.Worksheets("prixCentres"), Excel.Worksheet)
        Dim premiereDate As Integer = currentSheet.Cells(2, 1).Value
        Dim tailleEchant As Integer = currentSheet.UsedRange.Columns.Count - 1

        If fenetreEvDebut > fenetreEvFin Or fenetreEvFin > premiereDate + currentSheet.UsedRange.Rows.Count - 1 _
            Or fenetreEstDebut > fenetreEstFin Or fenetreEstDebut < premiereDate + 1 Or fenetreEstFin >= fenetreEvDebut Then
            MsgBox("Erreur : La fenêtre de temps de l'événement doit être cohérente avec les données", 16)
        Else
            MsgBox("ok")
        End If
        Globals.Ribbons.Ruban.paramARTaskPane.Visible = False
    End Sub

End Class
