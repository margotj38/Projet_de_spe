Imports System.Drawing.Imaging.ImageFormat

Public Class GraphiquePValeur

    Private Sub SaveGraph_Click_1(sender As Object, e As EventArgs) Handles SaveGraph.Click
        SaveFileDialog1.ShowDialog()
    End Sub

    Private Sub SaveFileDialog1_FileOk(sender As Object, e As ComponentModel.CancelEventArgs) Handles SaveFileDialog1.FileOk
        Dim fileToSave As String = SaveFileDialog1.FileName
        Me.GraphiqueChart.SaveImage(fileToSave, Png)
    End Sub

End Class