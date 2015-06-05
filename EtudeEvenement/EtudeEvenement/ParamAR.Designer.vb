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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ParamAR))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.nomModel = New System.Windows.Forms.Label()
        Me.LancementEtEv = New System.Windows.Forms.Button()
        Me.FenetreText = New System.Windows.Forms.Label()
        Me.estimation = New LeafCreations.Excel2007RefEdit()
        Me.evenement = New LeafCreations.Excel2007RefEdit()
        Me.SuspendLayout()
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
        Me.LancementEtEv.Location = New System.Drawing.Point(63, 227)
        Me.LancementEtEv.Name = "LancementEtEv"
        Me.LancementEtEv.Size = New System.Drawing.Size(111, 23)
        Me.LancementEtEv.TabIndex = 20
        Me.LancementEtEv.Text = "Lancement"
        Me.LancementEtEv.UseVisualStyleBackColor = True
        '
        'FenetreText
        '
        Me.FenetreText.AutoSize = True
        Me.FenetreText.Location = New System.Drawing.Point(25, 147)
        Me.FenetreText.Name = "FenetreText"
        Me.FenetreText.Size = New System.Drawing.Size(98, 13)
        Me.FenetreText.TabIndex = 17
        Me.FenetreText.Text = "Plage d'événement"
        '
        'estimation
        '
        Me.estimation.Address = Nothing
        Me.estimation.BackColor = System.Drawing.Color.Transparent
        Me.estimation.ExcelConnector = Nothing
        Me.estimation.ImageMaximized = CType(resources.GetObject("estimation.ImageMaximized"), System.Drawing.Image)
        Me.estimation.ImageMinimized = CType(resources.GetObject("estimation.ImageMinimized"), System.Drawing.Image)
        Me.estimation.Location = New System.Drawing.Point(26, 95)
        Me.estimation.Name = "estimation"
        Me.estimation.RefEditFont = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.estimation.Size = New System.Drawing.Size(195, 26)
        Me.estimation.TabIndex = 24
        '
        'evenement
        '
        Me.evenement.Address = Nothing
        Me.evenement.BackColor = System.Drawing.Color.Transparent
        Me.evenement.ExcelConnector = Nothing
        Me.evenement.ImageMaximized = CType(resources.GetObject("evenement.ImageMaximized"), System.Drawing.Image)
        Me.evenement.ImageMinimized = CType(resources.GetObject("evenement.ImageMinimized"), System.Drawing.Image)
        Me.evenement.Location = New System.Drawing.Point(26, 163)
        Me.evenement.Name = "evenement"
        Me.evenement.RefEditFont = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.evenement.Size = New System.Drawing.Size(195, 26)
        Me.evenement.TabIndex = 25
        '
        'ParamAR
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.evenement)
        Me.Controls.Add(Me.estimation)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.nomModel)
        Me.Controls.Add(Me.LancementEtEv)
        Me.Controls.Add(Me.FenetreText)
        Me.Name = "ParamAR"
        Me.Size = New System.Drawing.Size(309, 341)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents nomModel As System.Windows.Forms.Label
    Friend WithEvents LancementEtEv As System.Windows.Forms.Button
    Friend WithEvents FenetreText As System.Windows.Forms.Label
    Friend WithEvents estimation As LeafCreations.Excel2007RefEdit
    Friend WithEvents evenement As LeafCreations.Excel2007RefEdit

End Class
