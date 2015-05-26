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
        Me.Ok = New System.Windows.Forms.Button()
        Me.ChoixBox = New System.Windows.Forms.TextBox()
        Me.ChoixText = New System.Windows.Forms.Label()
        Me.FenetreText = New System.Windows.Forms.Label()
        Me.FenetreBox = New System.Windows.Forms.TextBox()
        Me.FenetreFinBox = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'Ok
        '
        Me.Ok.Location = New System.Drawing.Point(130, 175)
        Me.Ok.Name = "Ok"
        Me.Ok.Size = New System.Drawing.Size(58, 21)
        Me.Ok.TabIndex = 5
        Me.Ok.Text = "Ok"
        Me.Ok.UseVisualStyleBackColor = True
        '
        'ChoixBox
        '
        Me.ChoixBox.Location = New System.Drawing.Point(54, 55)
        Me.ChoixBox.Name = "ChoixBox"
        Me.ChoixBox.Size = New System.Drawing.Size(100, 20)
        Me.ChoixBox.TabIndex = 4
        '
        'ChoixText
        '
        Me.ChoixText.AutoSize = True
        Me.ChoixText.Location = New System.Drawing.Point(3, 25)
        Me.ChoixText.Name = "ChoixText"
        Me.ChoixText.Size = New System.Drawing.Size(185, 13)
        Me.ChoixText.TabIndex = 3
        Me.ChoixText.Text = "Seuil à utiliser pour le test statistique : "
        '
        'FenetreText
        '
        Me.FenetreText.AutoSize = True
        Me.FenetreText.Location = New System.Drawing.Point(3, 96)
        Me.FenetreText.Name = "FenetreText"
        Me.FenetreText.Size = New System.Drawing.Size(160, 13)
        Me.FenetreText.TabIndex = 6
        Me.FenetreText.Text = "Fenêtre autour de l'événement : "
        '
        'FenetreBox
        '
        Me.FenetreBox.Location = New System.Drawing.Point(54, 127)
        Me.FenetreBox.Name = "FenetreBox"
        Me.FenetreBox.Size = New System.Drawing.Size(100, 20)
        Me.FenetreBox.TabIndex = 7
        '
        'FenetreFinBox
        '
        Me.FenetreFinBox.Location = New System.Drawing.Point(197, 127)
        Me.FenetreFinBox.Name = "FenetreFinBox"
        Me.FenetreFinBox.Size = New System.Drawing.Size(100, 20)
        Me.FenetreFinBox.TabIndex = 8
        '
        'ChoixSeuilFenetre
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.FenetreFinBox)
        Me.Controls.Add(Me.FenetreBox)
        Me.Controls.Add(Me.FenetreText)
        Me.Controls.Add(Me.Ok)
        Me.Controls.Add(Me.ChoixBox)
        Me.Controls.Add(Me.ChoixText)
        Me.Name = "ChoixSeuilFenetre"
        Me.Size = New System.Drawing.Size(337, 219)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Ok As System.Windows.Forms.Button
    Friend WithEvents ChoixBox As System.Windows.Forms.TextBox
    Friend WithEvents ChoixText As System.Windows.Forms.Label
    Friend WithEvents FenetreText As System.Windows.Forms.Label
    Friend WithEvents FenetreBox As System.Windows.Forms.TextBox
    Friend WithEvents FenetreFinBox As System.Windows.Forms.TextBox

End Class
