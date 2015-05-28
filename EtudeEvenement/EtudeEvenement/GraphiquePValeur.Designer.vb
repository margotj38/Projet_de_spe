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
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.GraphiqueChart = New System.Windows.Forms.DataVisualization.Charting.Chart()
        Me.SaveGraph = New System.Windows.Forms.Button()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.GraphiqueChart, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'SaveFileDialog1
        '
        Me.SaveFileDialog1.DefaultExt = "png"
        Me.SaveFileDialog1.Filter = "files|*.png"
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.GraphiqueChart)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.SaveGraph)
        Me.SplitContainer1.Size = New System.Drawing.Size(484, 361)
        Me.SplitContainer1.SplitterDistance = 316
        Me.SplitContainer1.TabIndex = 3
        '
        'GraphiqueChart
        '
        ChartArea1.Name = "ChartArea1"
        Me.GraphiqueChart.ChartAreas.Add(ChartArea1)
        Me.GraphiqueChart.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GraphiqueChart.Location = New System.Drawing.Point(0, 0)
        Me.GraphiqueChart.Name = "GraphiqueChart"
        Me.GraphiqueChart.Size = New System.Drawing.Size(484, 316)
        Me.GraphiqueChart.TabIndex = 3
        Me.GraphiqueChart.Text = "P-Valeur"
        Title1.Name = "Title1"
        Title1.Text = "P-Valeurs en fonction de la fenêtre d'événement"
        Me.GraphiqueChart.Titles.Add(Title1)
        '
        'SaveGraph
        '
        Me.SaveGraph.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SaveGraph.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SaveGraph.Location = New System.Drawing.Point(0, 0)
        Me.SaveGraph.Name = "SaveGraph"
        Me.SaveGraph.Size = New System.Drawing.Size(484, 41)
        Me.SaveGraph.TabIndex = 4
        Me.SaveGraph.Text = "Save"
        Me.SaveGraph.UseVisualStyleBackColor = True
        '
        'GraphiquePValeur
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(484, 361)
        Me.Controls.Add(Me.SplitContainer1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow
        Me.MaximumSize = New System.Drawing.Size(1000, 800)
        Me.MinimumSize = New System.Drawing.Size(400, 320)
        Me.Name = "GraphiquePValeur"
        Me.Text = "Graphique P-Valeur"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.GraphiqueChart, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents GraphiqueChart As System.Windows.Forms.DataVisualization.Charting.Chart
    Friend WithEvents SaveGraph As System.Windows.Forms.Button
End Class
