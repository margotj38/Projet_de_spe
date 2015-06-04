<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ChoixDatesEv
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ChoixDatesEv))
        Me.datesEv = New LeafCreations.Excel2007RefEdit()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lancementPreT = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'datesEv
        '
        Me.datesEv.Address = Nothing
        Me.datesEv.BackColor = System.Drawing.Color.Transparent
        Me.datesEv.ExcelConnector = Nothing
        Me.datesEv.ImageMaximized = CType(resources.GetObject("datesEv.ImageMaximized"), System.Drawing.Image)
        Me.datesEv.ImageMinimized = CType(resources.GetObject("datesEv.ImageMinimized"), System.Drawing.Image)
        Me.datesEv.Location = New System.Drawing.Point(24, 81)
        Me.datesEv.Name = "datesEv"
        Me.datesEv.RefEditFont = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.datesEv.Size = New System.Drawing.Size(195, 26)
        Me.datesEv.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(21, 42)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(206, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Plage de données des dates d'événement"
        '
        'lancementPreT
        '
        Me.lancementPreT.Location = New System.Drawing.Point(74, 141)
        Me.lancementPreT.Name = "lancementPreT"
        Me.lancementPreT.Size = New System.Drawing.Size(88, 58)
        Me.lancementPreT.TabIndex = 2
        Me.lancementPreT.Text = "Lancer prétraitement"
        Me.lancementPreT.UseVisualStyleBackColor = True
        '
        'ChoixDatesEv
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.lancementPreT)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.datesEv)
        Me.Name = "ChoixDatesEv"
        Me.Size = New System.Drawing.Size(262, 302)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents datesEv As LeafCreations.Excel2007RefEdit
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lancementPreT As System.Windows.Forms.Button

End Class
