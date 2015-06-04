<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ParamAR
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
        Me.plageEstBox = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.nomModel = New System.Windows.Forms.Label()
        Me.LancementEtEv = New System.Windows.Forms.Button()
        Me.plageEvBox = New System.Windows.Forms.TextBox()
        Me.FenetreText = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'plageEstBox
        '
        Me.plageEstBox.Location = New System.Drawing.Point(26, 108)
        Me.plageEstBox.Name = "plageEstBox"
        Me.plageEstBox.Size = New System.Drawing.Size(100, 20)
        Me.plageEstBox.TabIndex = 23
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(25, 79)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(95, 13)
        Me.Label1.TabIndex = 22
        Me.Label1.Text = "Plage d'estimation "
        '
        'nomModel
        '
        Me.nomModel.AutoSize = True
        Me.nomModel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.nomModel.Location = New System.Drawing.Point(38, 33)
        Me.nomModel.Name = "nomModel"
        Me.nomModel.Size = New System.Drawing.Size(234, 16)
        Me.nomModel.TabIndex = 21
        Me.nomModel.Text = "Paramétrage de l'étude avec AR"
        '
        'LancementEtEv
        '
        Me.LancementEtEv.Location = New System.Drawing.Point(96, 168)
        Me.LancementEtEv.Name = "LancementEtEv"
        Me.LancementEtEv.Size = New System.Drawing.Size(111, 23)
        Me.LancementEtEv.TabIndex = 20
        Me.LancementEtEv.Text = "Lancement"
        Me.LancementEtEv.UseVisualStyleBackColor = True
        '
        'plageEvBox
        '
        Me.plageEvBox.Location = New System.Drawing.Point(177, 111)
        Me.plageEvBox.Name = "plageEvBox"
        Me.plageEvBox.Size = New System.Drawing.Size(100, 20)
        Me.plageEvBox.TabIndex = 19
        '
        'FenetreText
        '
        Me.FenetreText.AutoSize = True
        Me.FenetreText.Location = New System.Drawing.Point(174, 79)
        Me.FenetreText.Name = "FenetreText"
        Me.FenetreText.Size = New System.Drawing.Size(98, 13)
        Me.FenetreText.TabIndex = 17
        Me.FenetreText.Text = "Plage d'événement"
        '
        'ParamAR
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.plageEstBox)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.nomModel)
        Me.Controls.Add(Me.LancementEtEv)
        Me.Controls.Add(Me.plageEvBox)
        Me.Controls.Add(Me.FenetreText)
        Me.Name = "ParamAR"
        Me.Size = New System.Drawing.Size(322, 341)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents plageEstBox As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents nomModel As System.Windows.Forms.Label
    Friend WithEvents LancementEtEv As System.Windows.Forms.Button
    Friend WithEvents plageEvBox As System.Windows.Forms.TextBox
    Friend WithEvents FenetreText As System.Windows.Forms.Label

End Class
