﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SelectionDatesEv
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(SelectionDatesEv))
        Me.LabelDatesEv = New System.Windows.Forms.Label()
        Me.lancementPreT = New System.Windows.Forms.Button()
        Me.datesEvRefEdit = New LeafCreations.Excel2007RefEdit()
        Me.nomFeuilleBox = New System.Windows.Forms.TextBox()
        Me.LabelFeuille = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'LabelDatesEv
        '
        Me.LabelDatesEv.AutoSize = True
        Me.LabelDatesEv.Location = New System.Drawing.Point(23, 28)
        Me.LabelDatesEv.Name = "LabelDatesEv"
        Me.LabelDatesEv.Size = New System.Drawing.Size(211, 13)
        Me.LabelDatesEv.TabIndex = 1
        Me.LabelDatesEv.Text = "Plage de données des dates d'événements"
        '
        'lancementPreT
        '
        Me.lancementPreT.Location = New System.Drawing.Point(107, 216)
        Me.lancementPreT.Name = "lancementPreT"
        Me.lancementPreT.Size = New System.Drawing.Size(88, 58)
        Me.lancementPreT.TabIndex = 3
        Me.lancementPreT.Text = "Lancer prétraitement"
        Me.lancementPreT.UseVisualStyleBackColor = True
        '
        'datesEvRefEdit
        '
        Me.datesEvRefEdit.Address = Nothing
        Me.datesEvRefEdit.BackColor = System.Drawing.Color.Transparent
        Me.datesEvRefEdit.ExcelConnector = Nothing
        Me.datesEvRefEdit.ImageMaximized = CType(resources.GetObject("datesEvRefEdit.ImageMaximized"), System.Drawing.Image)
        Me.datesEvRefEdit.ImageMinimized = CType(resources.GetObject("datesEvRefEdit.ImageMinimized"), System.Drawing.Image)
        Me.datesEvRefEdit.Location = New System.Drawing.Point(39, 64)
        Me.datesEvRefEdit.Name = "datesEvRefEdit"
        Me.datesEvRefEdit.RefEditFont = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.datesEvRefEdit.Size = New System.Drawing.Size(195, 26)
        Me.datesEvRefEdit.TabIndex = 0
        '
        'nomFeuilleBox
        '
        Me.nomFeuilleBox.Location = New System.Drawing.Point(39, 162)
        Me.nomFeuilleBox.Name = "nomFeuilleBox"
        Me.nomFeuilleBox.Size = New System.Drawing.Size(137, 20)
        Me.nomFeuilleBox.TabIndex = 4
        '
        'LabelFeuille
        '
        Me.LabelFeuille.AutoSize = True
        Me.LabelFeuille.Location = New System.Drawing.Point(23, 123)
        Me.LabelFeuille.Name = "LabelFeuille"
        Me.LabelFeuille.Size = New System.Drawing.Size(196, 13)
        Me.LabelFeuille.TabIndex = 5
        Me.LabelFeuille.Text = "Nom de la feuille où trouver les données"
        '
        'SelectionDatesEv
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(308, 295)
        Me.Controls.Add(Me.LabelFeuille)
        Me.Controls.Add(Me.nomFeuilleBox)
        Me.Controls.Add(Me.lancementPreT)
        Me.Controls.Add(Me.LabelDatesEv)
        Me.Controls.Add(Me.datesEvRefEdit)
        Me.Name = "SelectionDatesEv"
        Me.Text = "SelectionDatesEv"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents datesEvRefEdit As LeafCreations.Excel2007RefEdit
    Friend WithEvents LabelDatesEv As System.Windows.Forms.Label
    Friend WithEvents lancementPreT As System.Windows.Forms.Button
    Friend WithEvents nomFeuilleBox As System.Windows.Forms.TextBox
    Friend WithEvents LabelFeuille As System.Windows.Forms.Label
End Class
