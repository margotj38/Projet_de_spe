Public Class ChoixSeuilFenetre

    Private textFenetreDebut As String
    Private textFenetreFin As String

    Public modele As Integer = -1 '-1 => probleme; 0 => ModeleMoyenne; 1 => ModeleRentaMarche; 2 => ModeleMarche

    Private Sub FenetreBox_TextChanged(sender As Object, e As EventArgs) Handles FenetreBox.TextChanged
        textFenetreDebut = FenetreBox.Text
    End Sub

    Private Sub FenetreFinBox_TextChanged(sender As Object, e As EventArgs) Handles FenetreFinBox.TextChanged
        textFenetreFin = FenetreFinBox.Text
    End Sub

    Private Sub PValeur_Click(sender As Object, e As EventArgs) Handles PValeur.Click
        Try
            Dim fenetreDebut As Integer = CInt(textFenetreDebut)
            Dim fenetreFin As Integer = CInt(textFenetreFin)
            Dim currentSheet As Excel.Worksheet = CType(Globals.ThisAddIn.Application.Worksheets("Rt"), Excel.Worksheet)
            Dim premiereDate As Integer = currentSheet.Cells(2, 1).Value
            If fenetreDebut > fenetreFin Or fenetreDebut <= premiereDate Or fenetreFin > premiereDate + currentSheet.UsedRange.Rows.Count - 1 Then
                MsgBox("Erreur : La fenêtre de temps de l'événement doit être cohérente avec les données", 16)
            Else
                If modele < 0 Or modele > 3 Then
                    MsgBox("Erreur interne : Provient de ChoixSeuilFenetre.vb", 16)
                End If
                Dim pValeur As Double = Globals.ThisAddIn.ThisAddIn_PValeur(modele, fenetreDebut, fenetreFin) * 100
                MsgBox("P-Valeur : " & pValeur.ToString("0.0000") & "%")
                Globals.Ribbons.Ruban.seuilFenetreTaskPane.Visible = False
            End If
        Catch erreur As InvalidCastException
            MsgBox("Erreur : Vous devez entrer des données correctes (utiliser la virgule pour les nombres décimaux)", 16)
        End Try
    End Sub
End Class
