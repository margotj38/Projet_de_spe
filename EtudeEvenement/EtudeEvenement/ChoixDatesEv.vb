Imports System.Runtime.InteropServices

Public Class ChoixDatesEv

    Private plageDates As String
    Private nomFeuille As String
    Private ligne As Integer
    Private colonne As Integer

    Private Sub lancementPreT_Click(sender As Object, e As EventArgs) Handles lancementPreT.Click
        MsgBox(plageDates)
        MsgBox(nomFeuille)
        MsgBox(ligne)
        MsgBox(colonne)
    End Sub

    Private Sub datesEv_Click(sender As Object, e As EventArgs) Handles datesEv.Click
        Dim excelApp As Excel.Application = Nothing

        ' Create an Excel App
        Try
            excelApp = Marshal.GetActiveObject("Excel.Application")
        Catch ex As COMException
            ' An exception is thrown if there is not an open excel instance.                    
        Finally
            If excelApp Is Nothing Then
                'excelApp = New Application
                excelApp.Workbooks.Add()
            End If
            excelApp.Visible = True

            Me.datesEv.ExcelConnector = excelApp
            Me.datesEv.ExcelConnector = excelApp
            Me.datesEv.ExcelConnector = excelApp
        End Try

        Me.datesEv.Focus()
    End Sub

End Class
