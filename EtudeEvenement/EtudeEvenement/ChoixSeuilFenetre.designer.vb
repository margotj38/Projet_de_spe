<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ChoixSeuilFenetre
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
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
        Me.FenetreText = New System.Windows.Forms.Label()
        Me.FenetreBox = New System.Windows.Forms.TextBox()
        Me.FenetreFinBox = New System.Windows.Forms.TextBox()
        Me.PValeur = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'FenetreText
        '
        Me.FenetreText.AutoSize = True
        Me.FenetreText.Location = New System.Drawing.Point(9, 22)
        Me.FenetreText.Name = "FenetreText"
        Me.FenetreText.Size = New System.Drawing.Size(151, 13)
        Me.FenetreText.TabIndex = 6
        Me.FenetreText.Text = "Fenêtre autour de l'événement"
        '
        'FenetreBox
        '
        Me.FenetreBox.Location = New System.Drawing.Point(12, 56)
        Me.FenetreBox.Name = "FenetreBox"
        Me.FenetreBox.Size = New System.Drawing.Size(100, 20)
        Me.FenetreBox.TabIndex = 7
        '
        'FenetreFinBox
        '
        Me.FenetreFinBox.Location = New System.Drawing.Point(156, 56)
        Me.FenetreFinBox.Name = "FenetreFinBox"
        Me.FenetreFinBox.Size = New System.Drawing.Size(100, 20)
        Me.FenetreFinBox.TabIndex = 8
        '
        'PValeur
        '
        Me.PValeur.Location = New System.Drawing.Point(79, 108)
        Me.PValeur.Name = "PValeur"
        Me.PValeur.Size = New System.Drawing.Size(111, 23)
        Me.PValeur.TabIndex = 10
        Me.PValeur.Text = "Calculer la P-Valeur"
        Me.PValeur.UseVisualStyleBackColor = True
        '
        'ChoixSeuilFenetre
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.PValeur)
        Me.Controls.Add(Me.FenetreFinBox)
        Me.Controls.Add(Me.FenetreBox)
        Me.Controls.Add(Me.FenetreText)
        Me.Name = "ChoixSeuilFenetre"
        Me.Size = New System.Drawing.Size(473, 319)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents FenetreText As System.Windows.Forms.Label
    Friend WithEvents FenetreBox As System.Windows.Forms.TextBox
    Friend WithEvents FenetreFinBox As System.Windows.Forms.TextBox
    Friend WithEvents PValeur As System.Windows.Forms.Button

End Class
