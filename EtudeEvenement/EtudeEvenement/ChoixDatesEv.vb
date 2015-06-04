Public Class ChoixDatesEv

    'constructeur
    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub lancementPreT_Click(sender As Object, e As EventArgs) Handles lancementPreT.Click
        MsgBox(Me.datesEv.Address)
    End Sub

End Class
