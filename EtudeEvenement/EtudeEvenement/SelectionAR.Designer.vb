<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SelectionAR
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(SelectionAR))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.LancementEtEv = New System.Windows.Forms.Button()
        Me.FenetreText = New System.Windows.Forms.Label()
        Me.refEditEst = New LeafCreations.Excel2007RefEdit()
        Me.refEditEv = New LeafCreations.Excel2007RefEdit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(27, 36)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(104, 13)
        Me.Label1.TabIndex = 27
        Me.Label1.Text = "Fenêtre d'estimation "
        '
        'LancementEtEv
        '
        Me.LancementEtEv.Location = New System.Drawing.Point(93, 196)
        Me.LancementEtEv.Name = "LancementEtEv"
        Me.LancementEtEv.Size = New System.Drawing.Size(111, 23)
        Me.LancementEtEv.TabIndex = 26
        Me.LancementEtEv.Text = "Lancement"
        Me.LancementEtEv.UseVisualStyleBackColor = True
        '
        'FenetreText
        '
        Me.FenetreText.AutoSize = True
        Me.FenetreText.Location = New System.Drawing.Point(27, 115)
        Me.FenetreText.Name = "FenetreText"
        Me.FenetreText.Size = New System.Drawing.Size(151, 13)
        Me.FenetreText.TabIndex = 25
        Me.FenetreText.Text = "Fenêtre autour de l'événement"
        '
        'refEditEst
        '
        Me.refEditEst.Address = Nothing
        Me.refEditEst.BackColor = System.Drawing.Color.Transparent
        Me.refEditEst.ExcelConnector = Nothing
        Me.refEditEst.ImageMaximized = CType(resources.GetObject("refEditEst.ImageMaximized"), System.Drawing.Image)
        Me.refEditEst.ImageMinimized = CType(resources.GetObject("refEditEst.ImageMinimized"), System.Drawing.Image)
        Me.refEditEst.Location = New System.Drawing.Point(43, 65)
        Me.refEditEst.Name = "refEditEst"
        Me.refEditEst.RefEditFont = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.refEditEst.Size = New System.Drawing.Size(195, 26)
        Me.refEditEst.TabIndex = 28
        '
        'refEditEv
        '
        Me.refEditEv.Address = Nothing
        Me.refEditEv.BackColor = System.Drawing.Color.Transparent
        Me.refEditEv.ExcelConnector = Nothing
        Me.refEditEv.ImageMaximized = CType(resources.GetObject("refEditEv.ImageMaximized"), System.Drawing.Image)
        Me.refEditEv.ImageMinimized = CType(resources.GetObject("refEditEv.ImageMinimized"), System.Drawing.Image)
        Me.refEditEv.Location = New System.Drawing.Point(43, 147)
        Me.refEditEv.Name = "refEditEv"
        Me.refEditEv.RefEditFont = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.refEditEv.Size = New System.Drawing.Size(195, 26)
        Me.refEditEv.TabIndex = 30
        '
        'SelectionAR
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(284, 261)
        Me.Controls.Add(Me.refEditEv)
        Me.Controls.Add(Me.refEditEst)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.LancementEtEv)
        Me.Controls.Add(Me.FenetreText)
        Me.Name = "SelectionAR"
        Me.Text = "SelectionAR"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents refEditEst As LeafCreations.Excel2007RefEdit
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents LancementEtEv As System.Windows.Forms.Button
    Friend WithEvents FenetreText As System.Windows.Forms.Label
    Friend WithEvents refEditEv As LeafCreations.Excel2007RefEdit
End Class
