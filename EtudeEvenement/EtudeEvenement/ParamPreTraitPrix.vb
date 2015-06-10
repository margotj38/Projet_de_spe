Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices

''' <summary>
''' Windows Form permettant la gestion du pré-traitement des prix.
''' </summary>
''' <remarks>L'interface graphique de cette Windows Form contient un objet de type refEdit (pour sélectionner les dates 
''' d'événements), une textBox (contenant la feuille sur laquelle on doit récupérer les cours), deux checkBox (une pour savoir 
''' comment on calcule les rentabilités, une seconde pour savoir si ce sont des cours d'ouverture ou de clôture) et un bouton 
''' pour lancer le calcul des rentabilités centrées autour de la date d'événement.</remarks>
Public Class ParamPreTraitPrix

    ''' <summary>
    ''' Nom de la feuille sur laquelle on va récupérer les cours.
    ''' </summary>
    ''' <remarks></remarks>
    Private nomFeuille As String

    ''' <summary>
    ''' Variable permettant de savoir quel mode de calcul sera utilisé pour les rentabilités 
    ''' (False pour un calcul arithmétique (valeur par défaut), True pour un calcul logarithmique).
    ''' </summary>
    ''' <remarks></remarks>
    Private rLog As Boolean

    ''' <summary>
    ''' Variable permettant de savoir si l'on dispose de cours d'ouverture ou de clôture 
    ''' (False pour des cours de clôture (valeur par défaut), True pour des cours d'ouverture).
    ''' </summary>
    ''' <remarks></remarks>
    Private cOuv As Boolean

    ''' <summary>
    ''' Constructeur initialisant les variables de la classe à leur valeur par défaut.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()
        InitializeComponent()
        rLog = False
        cOuv = False
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
    ''' Méthode associée au clic sur le bouton "Lancer prétraitement". Elle centre les prix, calcule les rentabilités pour 
    ''' le marché et pour les entreprises et affiche ces rentabilités centrées
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

        'traitement des prix
        'On centre les cours des entreprises et du marché
        Dim tabPrixCentres(,) As Double = Nothing
        Dim tabMarcheCentre(,) As Double = Nothing
        OpPrixRenta.donneesCentrees(plage, feuilleDates, feuilleDonnees, tabPrixCentres, tabMarcheCentre, cOuv)

        'On calcule les rentabilités
        Dim tabRenta(tabPrixCentres.GetUpperBound(0) - 1, tabPrixCentres.GetUpperBound(1)) As Double
        Dim tabRentaMarche(tabMarcheCentre.GetUpperBound(0) - 1, tabMarcheCentre.GetUpperBound(1)) As Double
        Dim tabRentaClassiquesMarche(tabMarcheCentre.GetUpperBound(0) - 1, tabMarcheCentre.GetUpperBound(1)) As Double
        Dim maxPrixAbsent As Integer
        OpPrixRenta.calculTabRenta(tabPrixCentres, tabMarcheCentre, tabRenta, tabRentaMarche, tabRentaClassiquesMarche, _
                                              maxPrixAbsent, rLog)

        'On stocke le tableaux des rentabilités de marché et des entreprises dont on va avoir besoin
        OpPrixRenta.tabRentaMarche = tabRentaMarche
        OpPrixRenta.tabRenta = tabRenta
        OpPrixRenta.tabRentaClassiquesMarche = tabRentaClassiquesMarche
        'Idem pour maxPrixAbsent
        OpPrixRenta.maxRentAbs = maxPrixAbsent

        'On affiche ces rentabilités centrées
        ExcelDialogue.affichageRentaCentrees(tabRenta)

    End Sub

    ''' <summary>
    ''' Méthode qui récupère le nom de la feuille contenant les cours
    ''' </summary>
    ''' <param name="sender">Non utilisé</param>
    ''' <param name="e">Non utilisé</param>
    ''' <remarks></remarks>
    Private Sub nomFeuilleBox_TextChanged(sender As Object, e As EventArgs) Handles nomFeuilleBox.TextChanged
        nomFeuille = nomFeuilleBox.Text
    End Sub

    ''' <summary>
    ''' Méthode qui récupère la valeur de la checkBox correspondant à la sélection du mode de calcul des rentabilités
    ''' </summary>
    ''' <param name="sender">Non utilisé</param>
    ''' <param name="e">Non utilisé</param>
    ''' <remarks></remarks>
    Private Sub rentaLog_CheckedChanged(sender As Object, e As EventArgs) Handles rentaLog.CheckedChanged
        rLog = rentaLog.Checked
    End Sub

    ''' <summary>
    ''' Méthode qui récupère la valeur de la checkBox correspondant à la sélection du type de cours fournis 
    ''' (ouverture ou clôture)
    ''' </summary>
    ''' <param name="sender">Non utilisé</param>
    ''' <param name="e">Non utilisé</param>
    ''' <remarks></remarks>
    Private Sub coursOuverture_CheckedChanged(sender As Object, e As EventArgs) Handles coursOuverture.CheckedChanged
        cOuv = coursOuverture.Checked
    End Sub
End Class