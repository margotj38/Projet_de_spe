Partial Class Ruban
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
   Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Requis pour la prise en charge du Concepteur de composition de classes Windows.Forms
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'Cet appel est requis par le Concepteur de composants.
        InitializeComponent()

    End Sub

    'Component remplace la méthode Dispose pour nettoyer la liste des composants.
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

    'Requise par le Concepteur de composants
    Private components As System.ComponentModel.IContainer

    'REMARQUE : la procédure suivante est requise par le Concepteur de composants
    'Elle peut être modifiée à l'aide du Concepteur de composants.
    'Ne la modifiez pas à l'aide de l'éditeur de code.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Ruban))
        Me.EtudeEvenement = Me.Factory.CreateRibbonTab
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.Separator1 = Me.Factory.CreateRibbonSeparator
        Me.EtudeAR = Me.Factory.CreateRibbonMenu
        Me.testSimplAR = Me.Factory.CreateRibbonButton
        Me.testSignAR = Me.Factory.CreateRibbonButton
        Me.preTraitPrix = Me.Factory.CreateRibbonButton
        Me.preTraitRenta = Me.Factory.CreateRibbonButton
        Me.ModeleMoy = Me.Factory.CreateRibbonMenu
        Me.testSimpleR = Me.Factory.CreateRibbonButton
        Me.TestPatellR = Me.Factory.CreateRibbonButton
        Me.testSigneR = Me.Factory.CreateRibbonButton
        Me.ModeleMS = Me.Factory.CreateRibbonMenu
        Me.testSimpleMS = Me.Factory.CreateRibbonButton
        Me.testPatellMS = Me.Factory.CreateRibbonButton
        Me.testSigneMS = Me.Factory.CreateRibbonButton
        Me.ModeleM = Me.Factory.CreateRibbonMenu
        Me.testSimpleM = Me.Factory.CreateRibbonButton
        Me.testPatellM = Me.Factory.CreateRibbonButton
        Me.testSigneM = Me.Factory.CreateRibbonButton
        Me.EtudeEvenement.SuspendLayout()
        Me.Group2.SuspendLayout()
        Me.Group1.SuspendLayout()
        '
        'EtudeEvenement
        '
        Me.EtudeEvenement.Groups.Add(Me.Group2)
        Me.EtudeEvenement.Groups.Add(Me.Group1)
        Me.EtudeEvenement.Label = "Etude d'événements"
        Me.EtudeEvenement.Name = "EtudeEvenement"
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.EtudeAR)
        Me.Group2.Label = "Etude à partir des AR"
        Me.Group2.Name = "Group2"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.preTraitPrix)
        Me.Group1.Items.Add(Me.preTraitRenta)
        Me.Group1.Items.Add(Me.Separator1)
        Me.Group1.Items.Add(Me.ModeleMoy)
        Me.Group1.Items.Add(Me.ModeleMS)
        Me.Group1.Items.Add(Me.ModeleM)
        Me.Group1.Label = "Etude complète"
        Me.Group1.Name = "Group1"
        '
        'Separator1
        '
        Me.Separator1.Name = "Separator1"
        '
        'EtudeAR
        '
        Me.EtudeAR.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.EtudeAR.Image = CType(resources.GetObject("EtudeAR.Image"), System.Drawing.Image)
        Me.EtudeAR.Items.Add(Me.testSimplAR)
        Me.EtudeAR.Items.Add(Me.testSignAR)
        Me.EtudeAR.Label = "Choix du test"
        Me.EtudeAR.Name = "EtudeAR"
        Me.EtudeAR.ShowImage = True
        '
        'testSimplAR
        '
        Me.testSimplAR.Label = "Test simple"
        Me.testSimplAR.Name = "testSimplAR"
        Me.testSimplAR.ShowImage = True
        '
        'testSignAR
        '
        Me.testSignAR.Label = "Test de signe"
        Me.testSignAR.Name = "testSignAR"
        Me.testSignAR.ShowImage = True
        '
        'preTraitPrix
        '
        Me.preTraitPrix.Label = "Prétraitement prix"
        Me.preTraitPrix.Name = "preTraitPrix"
        '
        'preTraitRenta
        '
        Me.preTraitRenta.Label = "Prétraitement rentabilités"
        Me.preTraitRenta.Name = "preTraitRenta"
        '
        'ModeleMoy
        '
        Me.ModeleMoy.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ModeleMoy.Image = CType(resources.GetObject("ModeleMoy.Image"), System.Drawing.Image)
        Me.ModeleMoy.Items.Add(Me.testSimpleR)
        Me.ModeleMoy.Items.Add(Me.TestPatellR)
        Me.ModeleMoy.Items.Add(Me.testSigneR)
        Me.ModeleMoy.Label = "Modèle moyenne des rentabilités"
        Me.ModeleMoy.Name = "ModeleMoy"
        Me.ModeleMoy.ShowImage = True
        '
        'testSimpleR
        '
        Me.testSimpleR.Label = "Test simple"
        Me.testSimpleR.Name = "testSimpleR"
        Me.testSimpleR.ShowImage = True
        '
        'TestPatellR
        '
        Me.TestPatellR.Label = "Test de Patell"
        Me.TestPatellR.Name = "TestPatellR"
        Me.TestPatellR.ShowImage = True
        '
        'testSigneR
        '
        Me.testSigneR.Label = "Test de signe"
        Me.testSigneR.Name = "testSigneR"
        Me.testSigneR.ShowImage = True
        '
        'ModeleMS
        '
        Me.ModeleMS.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ModeleMS.Image = CType(resources.GetObject("ModeleMS.Image"), System.Drawing.Image)
        Me.ModeleMS.Items.Add(Me.testSimpleMS)
        Me.ModeleMS.Items.Add(Me.testPatellMS)
        Me.ModeleMS.Items.Add(Me.testSigneMS)
        Me.ModeleMS.Label = "Modèle de marché simplifié"
        Me.ModeleMS.Name = "ModeleMS"
        Me.ModeleMS.ShowImage = True
        '
        'testSimpleMS
        '
        Me.testSimpleMS.Label = "Test simple"
        Me.testSimpleMS.Name = "testSimpleMS"
        Me.testSimpleMS.ShowImage = True
        '
        'testPatellMS
        '
        Me.testPatellMS.Label = "Test de Patell"
        Me.testPatellMS.Name = "testPatellMS"
        Me.testPatellMS.ShowImage = True
        '
        'testSigneMS
        '
        Me.testSigneMS.Label = "Test de signe"
        Me.testSigneMS.Name = "testSigneMS"
        Me.testSigneMS.ShowImage = True
        '
        'ModeleM
        '
        Me.ModeleM.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ModeleM.Image = CType(resources.GetObject("ModeleM.Image"), System.Drawing.Image)
        Me.ModeleM.Items.Add(Me.testSimpleM)
        Me.ModeleM.Items.Add(Me.testPatellM)
        Me.ModeleM.Items.Add(Me.testSigneM)
        Me.ModeleM.Label = "Modèle de marché classique"
        Me.ModeleM.Name = "ModeleM"
        Me.ModeleM.ShowImage = True
        '
        'testSimpleM
        '
        Me.testSimpleM.Label = "Test simple"
        Me.testSimpleM.Name = "testSimpleM"
        Me.testSimpleM.ShowImage = True
        '
        'testPatellM
        '
        Me.testPatellM.Label = "Test de Patell"
        Me.testPatellM.Name = "testPatellM"
        Me.testPatellM.ShowImage = True
        '
        'testSigneM
        '
        Me.testSigneM.Label = "Test de signe"
        Me.testSigneM.Name = "testSigneM"
        Me.testSigneM.ShowImage = True
        '
        'Ruban
        '
        Me.Name = "Ruban"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.EtudeEvenement)
        Me.EtudeEvenement.ResumeLayout(False)
        Me.EtudeEvenement.PerformLayout()
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()

    End Sub

    Friend WithEvents EtudeEvenement As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ModeleMoy As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents ModeleMS As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents ModeleM As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents testSimpleR As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents TestPatellR As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents testSigneR As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents testSimpleMS As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents testPatellMS As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents testSigneMS As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents testSimpleM As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents testPatellM As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents testSigneM As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents preTraitRenta As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents preTraitPrix As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents EtudeAR As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents testSimplAR As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents testSignAR As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator1 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ruban() As Ruban
        Get
            Return Me.GetRibbon(Of Ruban)()
        End Get
    End Property
End Class
