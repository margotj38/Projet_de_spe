<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class GraphiquePValeur
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim ChartArea1 As System.Windows.Forms.DataVisualization.Charting.ChartArea = New System.Windows.Forms.DataVisualization.Charting.ChartArea()
        Dim Title1 As System.Windows.Forms.DataVisualization.Charting.Title = New System.Windows.Forms.DataVisualization.Charting.Title()
        Me.GraphiqueChart = New System.Windows.Forms.DataVisualization.Charting.Chart()
        CType(Me.GraphiqueChart, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GraphiqueChart
        '
        ChartArea1.Name = "ChartArea1"
        Me.GraphiqueChart.ChartAreas.Add(ChartArea1)
        Me.GraphiqueChart.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GraphiqueChart.Location = New System.Drawing.Point(0, 0)
        Me.GraphiqueChart.Name = "GraphiqueChart"
        Me.GraphiqueChart.Size = New System.Drawing.Size(391, 300)
        Me.GraphiqueChart.TabIndex = 1
        Me.GraphiqueChart.Text = "P-Valeur"
        Title1.Name = "Title1"
        Title1.Text = "P-Valeurs en fonction de la fenêtre d'événement"
        Me.GraphiqueChart.Titles.Add(Title1)
        '
        'GraphiquePValeur
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(391, 300)
        Me.Controls.Add(Me.GraphiqueChart)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow
        Me.Name = "GraphiquePValeur"
        Me.Text = "Graphique P-Valeur"
        CType(Me.GraphiqueChart, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GraphiqueChart As System.Windows.Forms.DataVisualization.Charting.Chart
End Class
