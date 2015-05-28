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
        Me.PValeurFenetre = New System.Windows.Forms.Button()
        Me.nomModel = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'FenetreText
        '
        Me.FenetreText.AutoSize = True
        Me.FenetreText.Location = New System.Drawing.Point(9, 61)
        Me.FenetreText.Name = "FenetreText"
        Me.FenetreText.Size = New System.Drawing.Size(151, 13)
        Me.FenetreText.TabIndex = 6
        Me.FenetreText.Text = "Fenêtre autour de l'événement"
        '
        'FenetreBox
        '
        Me.FenetreBox.Location = New System.Drawing.Point(12, 103)
        Me.FenetreBox.Name = "FenetreBox"
        Me.FenetreBox.Size = New System.Drawing.Size(100, 20)
        Me.FenetreBox.TabIndex = 7
        '
        'FenetreFinBox
        '
        Me.FenetreFinBox.Location = New System.Drawing.Point(160, 103)
        Me.FenetreFinBox.Name = "FenetreFinBox"
        Me.FenetreFinBox.Size = New System.Drawing.Size(100, 20)
        Me.FenetreFinBox.TabIndex = 8
        '
        'PValeur
        '
        Me.PValeur.Location = New System.Drawing.Point(81, 164)
        Me.PValeur.Name = "PValeur"
        Me.PValeur.Size = New System.Drawing.Size(111, 23)
        Me.PValeur.TabIndex = 10
        Me.PValeur.Text = "Calculer la P-Valeur"
        Me.PValeur.UseVisualStyleBackColor = True
        '
        'PValeurFenetre
        '
        Me.PValeurFenetre.Location = New System.Drawing.Point(70, 226)
        Me.PValeurFenetre.Name = "PValeurFenetre"
        Me.PValeurFenetre.Size = New System.Drawing.Size(133, 49)
        Me.PValeurFenetre.TabIndex = 11
        Me.PValeurFenetre.Text = "Tracer la P-Valeur en fonction de la fenêtre d'événement"
        Me.PValeurFenetre.UseVisualStyleBackColor = True
        '
        'nomModel
        '
        Me.nomModel.AutoSize = True
        Me.nomModel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.nomModel.Location = New System.Drawing.Point(24, 17)
        Me.nomModel.Name = "nomModel"
        Me.nomModel.Size = New System.Drawing.Size(0, 16)
        Me.nomModel.TabIndex = 12
        '
        'ChoixSeuilFenetre
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.nomModel)
        Me.Controls.Add(Me.PValeurFenetre)
        Me.Controls.Add(Me.PValeur)
        Me.Controls.Add(Me.FenetreFinBox)
        Me.Controls.Add(Me.FenetreBox)
        Me.Controls.Add(Me.FenetreText)
        Me.Name = "ChoixSeuilFenetre"
        Me.Size = New System.Drawing.Size(293, 318)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents FenetreText As System.Windows.Forms.Label
    Friend WithEvents FenetreBox As System.Windows.Forms.TextBox
    Friend WithEvents FenetreFinBox As System.Windows.Forms.TextBox
    Friend WithEvents PValeur As System.Windows.Forms.Button
    Friend WithEvents PValeurFenetre As System.Windows.Forms.Button
    Friend WithEvents nomModel As System.Windows.Forms.Label

End Class
