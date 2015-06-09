Imports Microsoft.Office.Tools.Ribbon
Imports System.Windows.Forms.DataVisualization.Charting
Imports System.Runtime.InteropServices
Imports System.Net.Mime.MediaTypeNames
Imports Microsoft.Office.Interop

''' <summary>
''' Ruban de notre complément. Permet de sélectionner différentes méthodes et différents tests pour effectuer l'étude 
''' d'événements.
''' </summary>
''' <remarks></remarks>
Public Class Ruban

    ''' <summary>
    ''' Windows Form gérant le pré-traitement des prix.
    ''' </summary>
    ''' <remarks></remarks>
    Public paramPrix As ParamPreTraitPrix

    ''' <summary>
    ''' Windows Form gérant le pré-traitement des prix.
    ''' </summary>
    ''' <remarks></remarks>
    Public paramRenta As ParamPreTraitRenta

    ''' <summary>
    ''' Windows Form gérant la sélection des fenêtres d'estimation et d'événement sur 
    ''' les rentabilités centrées.
    ''' </summary>
    ''' <remarks></remarks>
    Public selFenetres As SelectionFenetres

    ''' <summary>
    ''' Widows Form gérant la sélection des fenêtres d'estimation et d'événement sur 
    ''' les rentabilités anormales.
    ''' </summary>
    ''' <remarks></remarks>
    Public selAR As SelectionAR

    ''' <summary>
    ''' Méthode permettant d'initialiser une windows form de sélection de fenêtres d'estimation et d'événement.
    ''' </summary>
    ''' <param name="selFenetres">Windows Form à initialiser.</param>
    ''' <remarks></remarks>
    Private Sub initSelFenetres(selFenetres As SelectionFenetres)
        'Définition des options des refEdit
        selFenetres.refEditEst.IncludeSheetName = True
        selFenetres.refEditEst.ShowRowAbsoluteIndicator = False
        selFenetres.refEditEst.ShowColumnAbsoluteIndicator = False
        selFenetres.refEditEv.IncludeSheetName = True
        selFenetres.refEditEv.ShowRowAbsoluteIndicator = False
        selFenetres.refEditEv.ShowColumnAbsoluteIndicator = False
        selFenetres.TopMost = True
    End Sub

    ''' <summary>
    ''' Méthode exécutée lors du choix d'une étude d'événement avec le modèle "moyenne des rentabilités", et le test 
    ''' classique de Student
    ''' </summary>
    ''' <param name="sender">Non utilisé</param>
    ''' <param name="e">Non utilisé</param>
    ''' <remarks></remarks>
    Private Sub testSimpleR_Click(sender As Object, e As RibbonControlEventArgs) Handles testSimpleR.Click
        selFenetres = New SelectionFenetres(0, 0)
        initSelFenetres(selFenetres)
        selFenetres.nomModele.Text = "Test simple - Modèle moyenne des rentabilités"
        selFenetres.Visible = True
    End Sub

    ''' <summary>
    ''' Méthode exécutée lors du choix d'une étude d'événement avec le modèle "moyenne des rentabilités", et le test 
    ''' de Patell
    ''' </summary>
    ''' <param name="sender">Non utilisé</param>
    ''' <param name="e">Non utilisé</param>
    ''' <remarks></remarks>
    Private Sub TestPatellR_Click(sender As Object, e As RibbonControlEventArgs) Handles TestPatellR.Click
        selFenetres = New SelectionFenetres(0, 1)
        initSelFenetres(selFenetres)
        selFenetres.nomModele.Text = "Test de Patell - Modèle moyenne des rentabilités"
        selFenetres.Visible = True
    End Sub

    ''' <summary>
    ''' Méthode exécutée lors du choix d'une étude d'événement avec le modèle "moyenne des rentabilités", et le test 
    ''' de signe
    ''' </summary>
    ''' <param name="sender">Non utilisé</param>
    ''' <param name="e">Non utilisé</param>
    ''' <remarks></remarks>
    Private Sub testSigneR_Click(sender As Object, e As RibbonControlEventArgs) Handles testSigneR.Click
        selFenetres = New SelectionFenetres(0, 2)
        initSelFenetres(selFenetres)
        selFenetres.nomModele.Text = "Test de signe - Modèle moyenne des rentabilités"
        selFenetres.Visible = True
    End Sub

    ''' <summary>
    ''' Méthode exécutée lors du choix d'une étude d'événement avec le modèle de marché simplifié, et le test 
    ''' classique de Student
    ''' </summary>
    ''' <param name="sender">Non utilisé</param>
    ''' <param name="e">Non utilisé</param>
    ''' <remarks></remarks>
    Private Sub testSimpleMS_Click(sender As Object, e As RibbonControlEventArgs) Handles testSimpleMS.Click
        selFenetres = New SelectionFenetres(1, 0)
        initSelFenetres(selFenetres)
        selFenetres.nomModele.Text = "Test simple - Modèle de marché simplifié"
        selFenetres.Visible = True
    End Sub

    ''' <summary>
    ''' Méthode exécutée lors du choix d'une étude d'événement avec le modèle de marché simplifié, et le test 
    ''' de Patell
    ''' </summary>
    ''' <param name="sender">Non utilisé</param>
    ''' <param name="e">Non utilisé</param>
    ''' <remarks></remarks>
    Private Sub testPatellMS_Click(sender As Object, e As RibbonControlEventArgs) Handles testPatellMS.Click
        selFenetres = New SelectionFenetres(1, 1)
        initSelFenetres(selFenetres)
        selFenetres.nomModele.Text = "Test de Patell - Modèle de marché simplifié"
        selFenetres.Visible = True
    End Sub

    ''' <summary>
    ''' Méthode exécutée lors du choix d'une étude d'événement avec le modèle de marché simplifié, et le test 
    ''' de signe
    ''' </summary>
    ''' <param name="sender">Non utilisé</param>
    ''' <param name="e">Non utilisé</param>
    ''' <remarks></remarks>
    Private Sub testSigneMS_Click(sender As Object, e As RibbonControlEventArgs) Handles testSigneMS.Click
        selFenetres = New SelectionFenetres(1, 2)
        initSelFenetres(selFenetres)
        selFenetres.nomModele.Text = "Test de signe - Modèle de marché simplifié"
        selFenetres.Visible = True
    End Sub

    ''' <summary>
    ''' Méthode exécutée lors du choix d'une étude d'événement avec le modèle marché, et le test 
    ''' classique de Student
    ''' </summary>
    ''' <param name="sender">Non utilisé</param>
    ''' <param name="e">Non utilisé</param>
    ''' <remarks></remarks>
    Private Sub testSimpleM_Click(sender As Object, e As RibbonControlEventArgs) Handles testSimpleM.Click
        selFenetres = New SelectionFenetres(2, 0)
        initSelFenetres(selFenetres)
        selFenetres.nomModele.Text = "Test simple - Modèle de marché classique"
        'seuilFenetreTaskPane.Visible = True
        selFenetres.Visible = True
    End Sub

    ''' <summary>
    ''' Méthode exécutée lors du choix d'une étude d'événement avec le modèle de marché, et le test 
    ''' de Patell
    ''' </summary>
    ''' <param name="sender">Non utilisé</param>
    ''' <param name="e">Non utilisé</param>
    ''' <remarks></remarks>
    Private Sub testPatellM_Click(sender As Object, e As RibbonControlEventArgs) Handles testPatellM.Click
        selFenetres = New SelectionFenetres(2, 1)
        initSelFenetres(selFenetres)
        selFenetres.nomModele.Text = "Test de Patell - Modèle de marché classique"
        'seuilFenetreTaskPane.Visible = True
        selFenetres.Visible = True
    End Sub

    ''' <summary>
    ''' Méthode exécutée lors du choix d'une étude d'événement avec le modèle de marché, et le test 
    ''' de signe
    ''' </summary>
    ''' <param name="sender">Non utilisé</param>
    ''' <param name="e">Non utilisé</param>
    ''' <remarks></remarks>
    Private Sub testSigneM_Click(sender As Object, e As RibbonControlEventArgs) Handles testSigneM.Click
        selFenetres = New SelectionFenetres(2, 2)
        initSelFenetres(selFenetres)
        selFenetres.nomModele.Text = "Test de signe - Modèle de marché classique"
        selFenetres.Visible = True
    End Sub

    ''' <summary>
    ''' Méthode appelée lors du choix d'un pré-traitement sur les prix.
    ''' </summary>
    ''' <param name="sender">Non utilisé</param>
    ''' <param name="e">Non utilisé</param>
    ''' <remarks></remarks>
    Private Sub preTraitPrix_Click(sender As Object, e As RibbonControlEventArgs) Handles preTraitPrix.Click

        paramPrix = New ParamPreTraitPrix()

        'Définition des options de refEdit
        paramPrix.datesEvRefEdit.IncludeSheetName = True
        paramPrix.datesEvRefEdit.ShowRowAbsoluteIndicator = False
        paramPrix.datesEvRefEdit.ShowColumnAbsoluteIndicator = False
        paramPrix.TopMost = True
        paramPrix.Visible = True
        paramPrix.nomFeuilleBox.Text = "Prix"
    End Sub

    ''' <summary>
    ''' Méthode appelée lors du choix d'un pré-traitement sur les rentabilités.
    ''' </summary>
    ''' <param name="sender">Non utilisé</param>
    ''' <param name="e">Non utilisé</param>
    ''' <remarks></remarks>
    Private Sub preTraitRenta_Click(sender As Object, e As RibbonControlEventArgs) Handles preTraitRenta.Click

        paramRenta = New ParamPreTraitRenta()

        'Définition des options de refEdit
        paramRenta.datesEvRefEdit.IncludeSheetName = True
        paramRenta.datesEvRefEdit.ShowRowAbsoluteIndicator = False
        paramRenta.datesEvRefEdit.ShowColumnAbsoluteIndicator = False
        paramRenta.TopMost = True
        paramRenta.Visible = True
        paramRenta.nomFeuilleBox.Text = "Rent"
    End Sub

    ''' <summary>
    ''' Méthode exécutée lors du choix d'une étude d'événement à partir des AR avec et le test classique de Student
    ''' </summary>
    ''' <param name="sender">Non utilisé</param>
    ''' <param name="e">Non utilisé</param>
    ''' <remarks></remarks>
    Private Sub testSimplAR_Click(sender As Object, e As RibbonControlEventArgs) Handles testSimplAR.Click
        selAR = New SelectionAR(0)

        'Définition des options des refEdit
        selAR.refEditEst.IncludeSheetName = True
        selAR.refEditEst.ShowRowAbsoluteIndicator = False
        selAR.refEditEst.ShowColumnAbsoluteIndicator = False
        selAR.refEditEv.IncludeSheetName = True
        selAR.refEditEv.ShowRowAbsoluteIndicator = False
        selAR.refEditEv.ShowColumnAbsoluteIndicator = False
        selAR.TopMost = True

        selAR.Visible = True
    End Sub

    ''' <summary>
    ''' Méthode exécutée lors du choix d'une étude d'événement à partir des AR avec et le test de signe
    ''' </summary>
    ''' <param name="sender">Non utilisé</param>
    ''' <param name="e">Non utilisé</param>
    ''' <remarks></remarks>
    Private Sub testSignAR_Click(sender As Object, e As RibbonControlEventArgs) Handles testSignAR.Click
        selAR = New SelectionAR(1)

        'Définition des options des refEdit
        selAR.refEditEst.IncludeSheetName = True
        selAR.refEditEst.ShowRowAbsoluteIndicator = False
        selAR.refEditEst.ShowColumnAbsoluteIndicator = False
        selAR.refEditEv.IncludeSheetName = True
        selAR.refEditEv.ShowRowAbsoluteIndicator = False
        selAR.refEditEv.ShowColumnAbsoluteIndicator = False
        selAR.TopMost = True

        selAR.Visible = True
    End Sub

End Class
