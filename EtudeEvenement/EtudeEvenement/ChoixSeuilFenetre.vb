Public Class ChoixSeuilFenetre

    Private textSeuil As String
    Private textFenetre As String

    Private Sub ChoixBox_TextChanged(sender As Object, e As EventArgs) Handles ChoixBox.TextChanged
        textSeuil = ChoixBox.Text
    End Sub

    Private Sub FenetreBox_TextChanged(sender As Object, e As EventArgs) Handles FenetreBox.TextChanged
        textFenetre = FenetreBox.Text
    End Sub

    Private Sub Ok_Click(sender As Object, e As EventArgs) Handles Ok.Click
        Try
            Dim seuil As Double = CDbl(textSeuil)
            If seuil >= 1 Or seuil <= 0 Then
                MsgBox("Erreur : Le seuil doit être compris entre 0 et 1", 16)
            Else
                Dim fenetre As Integer = CInt(textFenetre)
                If fenetre <= 0 Or fenetre > Globals.ThisAddIn.Application.ActiveSheet.UsedRange.Rows.Count Then
                    MsgBox("Erreur : La fenêtre de temps de l'événement doit être positive et inférieure au nombre de données dont on dispose", 16)
                Else
                    Dim rejet As Boolean = Globals.ThisAddIn.ThisAddIn_CalcNormMoy(fenetre, seuil)
                    Globals.Ribbons.Ruban.seuilFenetreTaskPane.Visible = False
                    If rejet Then
                        MsgBox("Rejet de l'hypothèse")
                    Else
                        MsgBox("Non rejet de l'hypothèse")
                    End If
                End If
            End If
        Catch erreur As InvalidCastException
            MsgBox("Erreur : Vous devez entrer des données correctes (utiliser la virgule pour les nombres décimaux)", 16)
        End Try
    End Sub

End Class
