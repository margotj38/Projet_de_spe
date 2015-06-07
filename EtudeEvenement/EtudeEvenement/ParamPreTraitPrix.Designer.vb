<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ParamPreTraitPrix
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ParamPreTraitPrix))
        Me.LabelFeuille = New System.Windows.Forms.Label()
        Me.nomFeuilleBox = New System.Windows.Forms.TextBox()
        Me.lancementPreT = New System.Windows.Forms.Button()
        Me.LabelDatesEv = New System.Windows.Forms.Label()
        Me.datesEvRefEdit = New LeafCreations.Excel2007RefEdit()
        Me.rentaLog = New System.Windows.Forms.CheckBox()
        Me.coursOuverture = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'LabelFeuille
        '
        Me.LabelFeuille.AutoSize = True
        Me.LabelFeuille.Location = New System.Drawing.Point(29, 124)
        Me.LabelFeuille.Name = "LabelFeuille"
        Me.LabelFeuille.Size = New System.Drawing.Size(230, 13)
        Me.LabelFeuille.TabIndex = 10
        Me.LabelFeuille.Text = "Nom de la feuille où trouver les données de prix"
        '
        'nomFeuilleBox
        '
        Me.nomFeuilleBox.Location = New System.Drawing.Point(45, 163)
        Me.nomFeuilleBox.Name = "nomFeuilleBox"
        Me.nomFeuilleBox.Size = New System.Drawing.Size(137, 20)
        Me.nomFeuilleBox.TabIndex = 9
        '
        'lancementPreT
        '
        Me.lancementPreT.Location = New System.Drawing.Point(137, 301)
        Me.lancementPreT.Name = "lancementPreT"
        Me.lancementPreT.Size = New System.Drawing.Size(88, 58)
        Me.lancementPreT.TabIndex = 8
        Me.lancementPreT.Text = "Lancer prétraitement"
        Me.lancementPreT.UseVisualStyleBackColor = True
        '
        'LabelDatesEv
        '
        Me.LabelDatesEv.AutoSize = True
        Me.LabelDatesEv.Location = New System.Drawing.Point(29, 29)
        Me.LabelDatesEv.Name = "LabelDatesEv"
        Me.LabelDatesEv.Size = New System.Drawing.Size(211, 13)
        Me.LabelDatesEv.TabIndex = 7
        Me.LabelDatesEv.Text = "Plage de données des dates d'événements"
        '
        'datesEvRefEdit
        '
        Me.datesEvRefEdit.Address = Nothing
        Me.datesEvRefEdit.BackColor = System.Drawing.Color.Transparent
        Me.datesEvRefEdit.ExcelConnector = Nothing
        Me.datesEvRefEdit.ImageMaximized = CType(resources.GetObject("datesEvRefEdit.ImageMaximized"), System.Drawing.Image)
        Me.datesEvRefEdit.ImageMinimized = CType(resources.GetObject("datesEvRefEdit.ImageMinimized"), System.Drawing.Image)
        Me.datesEvRefEdit.Location = New System.Drawing.Point(45, 65)
        Me.datesEvRefEdit.Name = "datesEvRefEdit"
        Me.datesEvRefEdit.RefEditFont = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.datesEvRefEdit.Size = New System.Drawing.Size(195, 26)
        Me.datesEvRefEdit.TabIndex = 6
        '
        'rentaLog
        '
        Me.rentaLog.AutoSize = True
        Me.rentaLog.Location = New System.Drawing.Point(45, 228)
        Me.rentaLog.Name = "rentaLog"
        Me.rentaLog.Size = New System.Drawing.Size(151, 17)
        Me.rentaLog.TabIndex = 11
        Me.rentaLog.Text = "Rentabilités logarithmiques"
        Me.rentaLog.UseVisualStyleBackColor = True
        '
        'coursOuverture
        '
        Me.coursOuverture.AutoSize = True
        Me.coursOuverture.Location = New System.Drawing.Point(235, 228)
        Me.coursOuverture.Name = "coursOuverture"
        Me.coursOuverture.Size = New System.Drawing.Size(109, 17)
        Me.coursOuverture.TabIndex = 12
        Me.coursOuverture.Text = "Cours d'ouverture"
        Me.coursOuverture.UseVisualStyleBackColor = True
        '
        'ParamPreTraitPrix
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(382, 406)
        Me.Controls.Add(Me.coursOuverture)
        Me.Controls.Add(Me.rentaLog)
        Me.Controls.Add(Me.LabelFeuille)
        Me.Controls.Add(Me.nomFeuilleBox)
        Me.Controls.Add(Me.lancementPreT)
        Me.Controls.Add(Me.LabelDatesEv)
        Me.Controls.Add(Me.datesEvRefEdit)
        Me.Name = "ParamPreTraitPrix"
        Me.Text = "ParamPreTraitPrix"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents LabelFeuille As System.Windows.Forms.Label
    Friend WithEvents nomFeuilleBox As System.Windows.Forms.TextBox
    Friend WithEvents lancementPreT As System.Windows.Forms.Button
    Friend WithEvents LabelDatesEv As System.Windows.Forms.Label
    Friend WithEvents datesEvRefEdit As LeafCreations.Excel2007RefEdit
    Friend WithEvents rentaLog As System.Windows.Forms.CheckBox
    Friend WithEvents coursOuverture As System.Windows.Forms.CheckBox
End Class
