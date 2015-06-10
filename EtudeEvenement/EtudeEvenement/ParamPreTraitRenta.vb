Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices

''' <summary>
''' Windows Form permettant la gestion du pré-traitement des rentabilités.
''' </summary>
''' <remarks>L'interface graphique de cette Windows Form contient un objet de type refEdit (pour sélectionner les dates 
''' d'événements), une textBox (contenant la feuille sur laquelle on doit récupérer les rentabilités) et un bouton 
''' pour lancer le calcul des rentabilités centrées autour de la date d'événement.</remarks>
Public Class ParamPreTraitRenta

    ''' <summary>
    ''' Nom de la feuille sur laquelle on va récupérer les cours.
    ''' </summary>
    ''' <remarks></remarks>
    Private nomFeuille As String

    ''' <summary>
    ''' Constructeur initialisant les composants.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()
        InitializeComponent()
    End Sub

    ''' <summary>
    ''' Méthode appelée au chargement de cette classe. Elle connecte l'objet refEdit à l'application Excel en cours
    ''' d'exécution.
    ''' </summary>
    ''' <param name="sender">Non utilisé</param>
    ''' <param name="e">Non utilisé</param>
    ''' <remarks></remarks>
    Private Sub SelectionDatesEv_Load(sender As Object, e As EventArgs) Handles Me.Load
        Dim excelApp As Excel.Application = Nothing

        'Create an Excel App
        Try
            excelApp = Marshal.GetActiveObject("Excel.Application")
        Catch ex As COMException
            'An exception is thrown if there is not an open excel instance.                    
        Finally
            If excelApp Is Nothing Then
                excelApp = New Microsoft.Office.Interop.Excel.Application
                excelApp.Workbooks.Add()
            End If
            excelApp.Visible = True

            Me.datesEvRefEdit.ExcelConnector = excelApp
        End Try

        Me.datesEvRefEdit.Focus()
    End Sub

    ''' <summary>
    ''' Méthode associée au clic sur le bouton "Lancer prétraitement". Elle centre les rentabilités et les affiche
    ''' </summary>
    ''' <param name="sender">Non utilisé</param>
    ''' <param name="e">Non utilisé</param>
    ''' <remarks></remarks>
    Private Sub lancementPreT_Click(sender As Object, e As EventArgs) Handles lancementPreT.Click
        'On récupère la plage des dates et la feuille sur laquelle elle est
        Dim plage As String = ""
        Dim feuilleDonnees As String = Me.nomFeuille
        Dim feuilleDates As String = ""
        'Si aucune plage n'a été sélectionnée, on lève une erreur
        If Me.datesEvRefEdit.Address = "" Then
            MsgBox("Erreur : Aucune plage de dates n'a été sélectionnée", 16)
            Return
        End If
        Utilitaires.recupererFeuillePlage(Me.datesEvRefEdit.Address, feuilleDates, plage)

        'On compte le nom de colonnes de la plage fournie
        Dim premiereColonne As Integer = 0
        Dim derniereColonne As Integer = 0
        Utilitaires.parserPlageColonnes(plage, premiereColonne, derniereColonne)
        'S'il n'y a pas exactement une colonne, on lève une erreur
        If Not premiereColonne = derniereColonne Then
            MsgBox("Erreur : Vous devez sélectionner uniquement la colonne des dates", 16)
            Return
        End If

        'Traitement des rentabilités
        'On centre les rentabilités (2ème colonne : marché)
        Dim tabRentaCentrees(,) As Double = Nothing
        Dim tabMarcheCentre(,) As Double = Nothing

        OpPrixRenta.donneesCentrees(plage, feuilleDates, feuilleDonnees, tabRentaCentrees, tabMarcheCentre, False)

        'On stocke le tableaux des rentabilités de marché dont on va avoir besoin
        'PB : où ? Dans nouveau module rentabilité ?
        OpPrixRenta.tabRentaMarche = tabMarcheCentre
        OpPrixRenta.tabRenta = tabRentaCentrees
        OpPrixRenta.tabRentaClassiquesMarche = tabMarcheCentre

        'Calcul de maxPrixAbsent
        Dim maxPrixAbsent As Integer = OpPrixRenta.calculMaxRentAbs(tabRentaCentrees)
        OpPrixRenta.maxRentAbs = maxPrixAbsent

        'On affiche les rentabilités centrées
        ExcelDialogue.affichageRentaCentrees(tabRentaCentrees)

    End Sub

    ''' <summary>
    ''' Méthode qui récupère le nom de la feuille contenant les rentabilités
    ''' </summary>
    ''' <param name="sender">Non utilisé</param>
    ''' <param name="e">Non utilisé</param>
    ''' <remarks></remarks>
    Private Sub nomFeuilleBox_TextChanged(sender As Object, e As EventArgs) Handles nomFeuilleBox.TextChanged
        nomFeuille = nomFeuilleBox.Text
    End Sub
End Class