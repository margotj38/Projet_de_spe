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
        Me.EtudeEvenement = Me.Factory.CreateRibbonTab
        Me.ModelesClassiques = Me.Factory.CreateRibbonGroup
        Me.ModeleMoyenne = Me.Factory.CreateRibbonButton
        Me.ModeleMarcheSimple = Me.Factory.CreateRibbonButton
        Me.ModeleMarche = Me.Factory.CreateRibbonButton
        Me.TestsTemp = Me.Factory.CreateRibbonGroup
        Me.test = Me.Factory.CreateRibbonButton
        Me.Group1 = Me.Factory.CreateRibbonGroup
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
        'TestsTemp
        '
        Me.TestsTemp.Items.Add(Me.test)
        Me.TestsTemp.Label = "Tests temporaires"
        Me.TestsTemp.Name = "TestsTemp"
        '
        'test
        '
        Me.test.Label = "test"
        Me.test.Name = "test"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.ModeleMoy)
        Me.Group1.Items.Add(Me.ModeleMS)
        Me.Group1.Items.Add(Me.ModeleM)
        Me.Group1.Label = "Modèles"
        Me.Group1.Name = "Group1"
        '
        'ModeleMoy
        '
        Me.ModeleMoy.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
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
    Friend WithEvents test As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ModeleMoyenne As Microsoft.Office.Tools.Ribbon.RibbonButton
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
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ruban() As Ruban
        Get
            Return Me.GetRibbon(Of Ruban)()
        End Get
    End Property
End Class
