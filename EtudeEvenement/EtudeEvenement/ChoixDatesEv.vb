Imports System.Runtime.InteropServices
Imports System.Net.Mime.MediaTypeNames
Imports Microsoft.Office.Interop

Public Class ChoixDatesEv

    'constructeur
    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub lancementPreT_Click(sender As Object, e As EventArgs) Handles lancementPreT.Click
        MsgBox(Me.datesEv.Address)
    End Sub

End Class
