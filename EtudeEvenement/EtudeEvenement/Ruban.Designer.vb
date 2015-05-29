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
        Me.ModelesClassiques = Me.Factory.CreateRibbonGroup
        Me.TestsTemp = Me.Factory.CreateRibbonGroup
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.ModeleMoyenne = Me.Factory.CreateRibbonButton
        Me.ModeleMarcheSimple = Me.Factory.CreateRibbonButton
        Me.ModeleMarche = Me.Factory.CreateRibbonButton
        Me.AR = Me.Factory.CreateRibbonButton
        Me.Menu1 = Me.Factory.CreateRibbonMenu
        Me.testSimpleR = Me.Factory.CreateRibbonButton
        Me.TestPatellR = Me.Factory.CreateRibbonButton
        Me.testSigneR = Me.Factory.CreateRibbonButton
        Me.Menu2 = Me.Factory.CreateRibbonMenu
        Me.testSimpleMS = Me.Factory.CreateRibbonButton
        Me.testPatellMS = Me.Factory.CreateRibbonButton
        Me.testSigneMS = Me.Factory.CreateRibbonButton
        Me.Menu3 = Me.Factory.CreateRibbonMenu
        Me.testSimpleM = Me.Factory.CreateRibbonButton
        Me.testPatellM = Me.Factory.CreateRibbonButton
        Me.testSigneM = Me.Factory.CreateRibbonButton
        Me.EtudeEvenement.SuspendLayout()
        Me.ModelesClassiques.SuspendLayout()
        Me.TestsTemp.SuspendLayout()
        Me.Group1.SuspendLayout()
        '
        'EtudeEvenement
        '
        Me.EtudeEvenement.Groups.Add(Me.ModelesClassiques)
        Me.EtudeEvenement.Groups.Add(Me.TestsTemp)
        Me.EtudeEvenement.Groups.Add(Me.Group1)
        Me.EtudeEvenement.Label = "Etude d'événements"
        Me.EtudeEvenement.Name = "EtudeEvenement"
        '
        'ModelesClassiques
        '
        Me.ModelesClassiques.Items.Add(Me.ModeleMoyenne)
        Me.ModelesClassiques.Items.Add(Me.ModeleMarcheSimple)
        Me.ModelesClassiques.Items.Add(Me.ModeleMarche)
        Me.ModelesClassiques.Label = "Modèles classiques"
        Me.ModelesClassiques.Name = "ModelesClassiques"
        '
        'TestsTemp
        '
        Me.TestsTemp.Items.Add(Me.AR)
        Me.TestsTemp.Label = "Tests temporaires"
        Me.TestsTemp.Name = "TestsTemp"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.Menu1)
        Me.Group1.Items.Add(Me.Menu2)
        Me.Group1.Items.Add(Me.Menu3)
        Me.Group1.Label = "Modèles"
        Me.Group1.Name = "Group1"
        '
        'ModeleMoyenne
        '
        Me.ModeleMoyenne.Label = "Modèle moyenne des rentabilités"
        Me.ModeleMoyenne.Name = "ModeleMoyenne"
        '
        'ModeleMarcheSimple
        '
        Me.ModeleMarcheSimple.Label = "Modèle de marché simplifié"
        Me.ModeleMarcheSimple.Name = "ModeleMarcheSimple"
        '
        'ModeleMarche
        '
        Me.ModeleMarche.Label = "Modèle de marché classique"
        Me.ModeleMarche.Name = "ModeleMarche"
        '
        'AR
        '
        Me.AR.Label = "Etude depuis AR"
        Me.AR.Name = "AR"
        '
        'Menu1
        '
        Me.Menu1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Menu1.Image = CType(resources.GetObject("Menu1.Image"), System.Drawing.Image)
        Me.Menu1.Items.Add(Me.testSimpleR)
        Me.Menu1.Items.Add(Me.TestPatellR)
        Me.Menu1.Items.Add(Me.testSigneR)
        Me.Menu1.Label = "Modèle moyenne des rentabilités"
        Me.Menu1.Name = "Menu1"
        Me.Menu1.ShowImage = True
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
        'Menu2
        '
        Me.Menu2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Menu2.Image = CType(resources.GetObject("Menu2.Image"), System.Drawing.Image)
        Me.Menu2.Items.Add(Me.testSimpleMS)
        Me.Menu2.Items.Add(Me.testPatellMS)
        Me.Menu2.Items.Add(Me.testSigneMS)
        Me.Menu2.Label = "Modèle de marché simplifié"
        Me.Menu2.Name = "Menu2"
        Me.Menu2.ShowImage = True
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
        'Menu3
        '
        Me.Menu3.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Menu3.Image = CType(resources.GetObject("Menu3.Image"), System.Drawing.Image)
        Me.Menu3.Items.Add(Me.testSimpleM)
        Me.Menu3.Items.Add(Me.testPatellM)
        Me.Menu3.Items.Add(Me.testSigneM)
        Me.Menu3.Label = "Modèle de marché classique"
        Me.Menu3.Name = "Menu3"
        Me.Menu3.ShowImage = True
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
        Me.ModelesClassiques.ResumeLayout(False)
        Me.ModelesClassiques.PerformLayout()
        Me.TestsTemp.ResumeLayout(False)
        Me.TestsTemp.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()

    End Sub

    Friend WithEvents EtudeEvenement As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents ModelesClassiques As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ModeleMarche As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ModeleMarcheSimple As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents TestsTemp As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents AR As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ModeleMoyenne As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Menu1 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents Menu2 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents Menu3 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents testSimpleR As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents TestPatellR As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents testSigneR As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents testSimpleMS As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents testPatellMS As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents testSigneMS As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents testSimpleM As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents testPatellM As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents testSigneM As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ruban() As Ruban
        Get
            Return Me.GetRibbon(Of Ruban)()
        End Get
    End Property
End Class
