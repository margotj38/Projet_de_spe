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
        Me.ModeleRentaMarche = Me.Factory.CreateRibbonButton
        Me.ModeleMarche = Me.Factory.CreateRibbonButton
        Me.TestsTemp = Me.Factory.CreateRibbonGroup
        Me.AR = Me.Factory.CreateRibbonButton
        Me.EtudeEvenement.SuspendLayout()
        Me.ModelesClassiques.SuspendLayout()
        Me.TestsTemp.SuspendLayout()
        '
        'EtudeEvenement
        '
        Me.EtudeEvenement.Groups.Add(Me.ModelesClassiques)
        Me.EtudeEvenement.Groups.Add(Me.TestsTemp)
        Me.EtudeEvenement.Label = "Etude d'événements"
        Me.EtudeEvenement.Name = "EtudeEvenement"
        '
        'ModelesClassiques
        '
        Me.ModelesClassiques.Items.Add(Me.ModeleMoyenne)
        Me.ModelesClassiques.Items.Add(Me.ModeleRentaMarche)
        Me.ModelesClassiques.Items.Add(Me.ModeleMarche)
        Me.ModelesClassiques.Label = "Modèles classiques"
        Me.ModelesClassiques.Name = "ModelesClassiques"
        '
        'ModeleMoyenne
        '
        Me.ModeleMoyenne.Label = "Modèle moyenne des rentabilités"
        Me.ModeleMoyenne.Name = "ModeleMoyenne"
        '
        'ModeleRentaMarche
        '
        Me.ModeleRentaMarche.Label = "Modèle rentabilité de marché"
        Me.ModeleRentaMarche.Name = "ModeleRentaMarche"
        '
        'ModeleMarche
        '
        Me.ModeleMarche.Label = "Modèle de marché"
        Me.ModeleMarche.Name = "ModeleMarche"
        '
        'TestsTemp
        '
        Me.TestsTemp.Items.Add(Me.AR)
        Me.TestsTemp.Label = "Tests temporaires"
        Me.TestsTemp.Name = "TestsTemp"
        '
        'AR
        '
        Me.AR.Label = "Etude depuis AR"
        Me.AR.Name = "AR"
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

    End Sub

    Friend WithEvents EtudeEvenement As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents ModelesClassiques As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ModeleMarche As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ModeleRentaMarche As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ModeleMoyenne As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents TestsTemp As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents AR As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ruban() As Ruban
        Get
            Return Me.GetRibbon(Of Ruban)()
        End Get
    End Property
End Class
