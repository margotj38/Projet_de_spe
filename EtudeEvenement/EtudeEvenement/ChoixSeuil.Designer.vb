<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ChoixSeuil
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
        Me.SuspendLayout()
        '
        'Ok
        '
        Me.Ok.Location = New System.Drawing.Point(110, 93)
        Me.Ok.Name = "Ok"
        Me.Ok.Size = New System.Drawing.Size(58, 21)
        Me.Ok.TabIndex = 2
        Me.Ok.Text = "Ok"
        Me.Ok.UseVisualStyleBackColor = True
        '
        'ChoixBox
        '
        Me.ChoixBox.Location = New System.Drawing.Point(88, 58)
        Me.ChoixBox.Name = "ChoixBox"
        Me.ChoixBox.Size = New System.Drawing.Size(100, 20)
        Me.ChoixBox.TabIndex = 1
        '
        'ChoixText
        '
        Me.ChoixText.AutoSize = True
        Me.ChoixText.Location = New System.Drawing.Point(3, 25)
        Me.ChoixText.Name = "ChoixText"
        Me.ChoixText.Size = New System.Drawing.Size(185, 13)
        Me.ChoixText.TabIndex = 0
        Me.ChoixText.Text = "Seuil à utiliser pour le test statistique : "
        '
        'ChoixSeuil
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.Ok)
        Me.Controls.Add(Me.ChoixBox)
        Me.Controls.Add(Me.ChoixText)
        Me.Name = "ChoixSeuil"
        Me.Size = New System.Drawing.Size(249, 151)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Ok As System.Windows.Forms.Button
    Friend WithEvents ChoixBox As System.Windows.Forms.TextBox
    Friend WithEvents ChoixText As System.Windows.Forms.Label

End Class
