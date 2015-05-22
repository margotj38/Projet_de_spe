Public Class ChoixSeuil

    Private textSeuil As String

    Private Sub ChoixBox_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChoixBox.TextChanged
        textSeuil = ChoixBox.Text
    End Sub

    Private Sub Ok_Click(sender As Object, e As EventArgs) Handles Ok.Click
        Try
            Dim seuil As Double = CDbl(textSeuil)
            If seuil > 1 Or seuil < 0 Then
                MsgBox("Erreur : Le seuil doit être compris entre 0 et 1")
            Else
                Dim rejet As Boolean = Globals.ThisAddIn.ThisAddIn_MethodeCAR(seuil)
                Globals.Ribbons.Ruban.myTaskPane.Visible = False
                If rejet Then
                    MsgBox("Rejet de l'hypothèse")
                Else
                    MsgBox("Non rejet de l'hypothèse")
                End If
            End If
        Catch erreur As InvalidCastException
            MsgBox("Erreur : Vous devez entrer un nombre décimal (utiliser la virgule)")
        End Try
    End Sub
End Class
