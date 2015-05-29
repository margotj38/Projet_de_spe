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
        Me.FenetreDebBox = New System.Windows.Forms.TextBox()
        Me.FenetreFinBox = New System.Windows.Forms.TextBox()
        Me.PValeur = New System.Windows.Forms.Button()
        Me.PValeurFenetre = New System.Windows.Forms.Button()
        Me.nomModel = New System.Windows.Forms.Label()
        Me.BoutonPatell = New System.Windows.Forms.Button()
        Me.FenetreEstFinBox = New System.Windows.Forms.TextBox()
        Me.FenetreEstDebBox = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'FenetreText
        '
        Me.FenetreText.AutoSize = True
        Me.FenetreText.Location = New System.Drawing.Point(9, 88)
        Me.FenetreText.Name = "FenetreText"
        Me.FenetreText.Size = New System.Drawing.Size(151, 13)
        Me.FenetreText.TabIndex = 6
        Me.FenetreText.Text = "Fenêtre autour de l'événement"
        '
        'FenetreDebBox
        '
        Me.FenetreDebBox.Location = New System.Drawing.Point(12, 125)
        Me.FenetreDebBox.Name = "FenetreDebBox"
        Me.FenetreDebBox.Size = New System.Drawing.Size(100, 20)
        Me.FenetreDebBox.TabIndex = 7
        '
        'FenetreFinBox
        '
        Me.FenetreFinBox.Location = New System.Drawing.Point(163, 125)
        Me.FenetreFinBox.Name = "FenetreFinBox"
        Me.FenetreFinBox.Size = New System.Drawing.Size(100, 20)
        Me.FenetreFinBox.TabIndex = 8
        '
        'PValeur
        '
        Me.PValeur.Location = New System.Drawing.Point(83, 190)
        Me.PValeur.Name = "PValeur"
        Me.PValeur.Size = New System.Drawing.Size(111, 23)
        Me.PValeur.TabIndex = 10
        Me.PValeur.Text = "Calculer la P-Valeur"
        Me.PValeur.UseVisualStyleBackColor = True
        '
        'PValeurFenetre
        '
        Me.PValeurFenetre.Location = New System.Drawing.Point(71, 245)
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
        'BoutonPatell
        '
        Me.BoutonPatell.Location = New System.Drawing.Point(99, 318)
        Me.BoutonPatell.Name = "BoutonPatell"
        Me.BoutonPatell.Size = New System.Drawing.Size(75, 23)
        Me.BoutonPatell.TabIndex = 13
        Me.BoutonPatell.Text = "Test Patell"
        Me.BoutonPatell.UseVisualStyleBackColor = True
        '
        'FenetreEstFinBox
        '
        Me.FenetreEstFinBox.Location = New System.Drawing.Point(163, 54)
        Me.FenetreEstFinBox.Name = "FenetreEstFinBox"
        Me.FenetreEstFinBox.Size = New System.Drawing.Size(100, 20)
        Me.FenetreEstFinBox.TabIndex = 16
        '
        'FenetreEstDebBox
        '
        Me.FenetreEstDebBox.Location = New System.Drawing.Point(12, 54)
        Me.FenetreEstDebBox.Name = "FenetreEstDebBox"
        Me.FenetreEstDebBox.Size = New System.Drawing.Size(100, 20)
        Me.FenetreEstDebBox.TabIndex = 15
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(9, 17)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(101, 13)
        Me.Label1.TabIndex = 14
        Me.Label1.Text = "Fenêtre d'estimation"
        '
        'ChoixSeuilFenetre
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.FenetreEstFinBox)
        Me.Controls.Add(Me.FenetreEstDebBox)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.BoutonPatell)
        Me.Controls.Add(Me.nomModel)
        Me.Controls.Add(Me.PValeurFenetre)
        Me.Controls.Add(Me.PValeur)
        Me.Controls.Add(Me.FenetreFinBox)
        Me.Controls.Add(Me.FenetreDebBox)
        Me.Controls.Add(Me.FenetreText)
        Me.Name = "ChoixSeuilFenetre"
        Me.Size = New System.Drawing.Size(294, 392)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents FenetreText As System.Windows.Forms.Label
    Friend WithEvents FenetreDebBox As System.Windows.Forms.TextBox
    Friend WithEvents FenetreFinBox As System.Windows.Forms.TextBox
    Friend WithEvents PValeur As System.Windows.Forms.Button
    Friend WithEvents PValeurFenetre As System.Windows.Forms.Button
    Friend WithEvents nomModel As System.Windows.Forms.Label
    Friend WithEvents BoutonPatell As System.Windows.Forms.Button
    Friend WithEvents FenetreEstFinBox As System.Windows.Forms.TextBox
    Friend WithEvents FenetreEstDebBox As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label

End Class
