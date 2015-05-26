Public Class ChoixSeuilFenetre

    Private textSeuil As String
    Private textFenetreDebut As String
    Private textFenetreFin As String

    Public modele As Integer = -1 '-1 => probleme; 0 => ModeleMoyenne; 1 => ModeleRentaMarche; 2 => ModeleMarche

    Private Sub ChoixBox_TextChanged(sender As Object, e As EventArgs) Handles ChoixBox.TextChanged
        textSeuil = ChoixBox.Text
    End Sub

    Private Sub FenetreBox_TextChanged(sender As Object, e As EventArgs) Handles FenetreBox.TextChanged
        textFenetreDebut = FenetreBox.Text
    End Sub

    Private Sub FenetreFinBox_TextChanged(sender As Object, e As EventArgs) Handles FenetreFinBox.TextChanged
        textFenetreFin = FenetreFinBox.Text
    End Sub

    Private Sub Ok_Click(sender As Object, e As EventArgs) Handles Ok.Click
        Try
            Dim seuil As Double = CDbl(textSeuil)
            If seuil >= 1 Or seuil <= 0 Then
                MsgBox("Erreur : Le seuil doit être compris entre 0 et 1", 16)
            Else
                Dim fenetreDebut As Integer = CInt(textFenetreDebut)
                Dim fenetreFin As Integer = CInt(textFenetreFin)
                Dim currentSheet As Excel.Worksheet = CType(Globals.ThisAddIn.Application.Worksheets("Rt"), Excel.Worksheet)
                Dim premiereDate As Integer = currentSheet.Cells(2, 1).Value
                If fenetreDebut <= premiereDate Or fenetreFin > premiereDate + currentSheet.UsedRange.Rows.Count - 1 Then
                    MsgBox("Erreur : La fenêtre de temps de l'événement doit être cohérente avec les données", 16)
                Else
                    Dim rejet As Boolean
                    Select Case modele
                        Case 0
                            rejet = Globals.ThisAddIn.ThisAddIn_CalcNormMoy(fenetreDebut, fenetreFin, seuil)
                        Case 1
                            rejet = Globals.ThisAddIn.ThisAddIn_ModeleRentaMarche(fenetreDebut, seuil)
                        Case 2
                            rejet = Globals.ThisAddIn.ThisAddIn_ModeleMarche(fenetreDebut, seuil)
                        Case Else
                            MsgBox("Erreur interne : Provient de ChoixSeuilFenetre.vb", 16)
                    End Select
                    Dim pValeur As Double = Globals.ThisAddIn.ThisAddIn_PValeur(modele, fenetreDebut)
                    MsgBox("P-Valeur : " & pValeur)
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
