Imports System.Windows.Forms.DataVisualization.Charting
Imports System.Runtime.InteropServices
Imports System.Net.Mime.MediaTypeNames
Imports Microsoft.Office.Interop

Public Class ChoixDatesEv

    Private donneesPreT As Integer ' 0 => prix; 1 => rentabilités 

    'accesseur sur model
    Public Property donneesPreTraitement As Integer
        Get
            Return donneesPreT
        End Get
        Set(value As Integer)
            If value < 0 Or value > 1 Then
                MsgBox("Erreur interne : numéro de données à analyser incorrect", 16)
            End If
            donneesPreT = value
        End Set
    End Property

    'constructeur
    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub lancementPreT_Click(sender As Object, e As EventArgs) Handles lancementPreT.Click
        'On récupère la plage des dates et la feuille sur laquelle elle est
        Dim plage As String = ""
        Dim feuille As String = ""
        Utilitaires.recupererFeuillePlage(Me.datesEv.Address, feuille, plage)

        Select Case donneesPreTraitement
            Case 0
                'Si on doit traiter des prix
                'On centre les cours des entreprises et du marché
                Dim tabPrixCentres(,) As Double = Nothing
                Dim tabMarcheCentre(,) As Double = Nothing
                UtilitaireRentabilites.donneesCentrees(plage, feuille, tabPrixCentres, tabMarcheCentre)

                'On calcule les rentabilités
                Dim tabRenta(tabPrixCentres.GetUpperBound(0) - 1, tabPrixCentres.GetUpperBound(1)) As Double
                Dim tabRentaMarche(tabMarcheCentre.GetUpperBound(0) - 1, tabMarcheCentre.GetUpperBound(1)) As Double
                'UtilitaireRentabilites.calculTabRenta(tabPrixCentres, tabMarcheCentre, tabRenta, tabRentaMarche)

                'On affiche ces rentabilités centrées
                ExcelDialogue.affichageRentaCentrees(tabRenta)
            Case 1
                'Si on doit traiter des rentabilités
                'On centre les cours des entreprises et du marché
                Dim tabRentaCentres(,) As Double = Nothing
                Dim tabMarcheCentre(,) As Double = Nothing
                UtilitaireRentabilites.donneesCentrees(plage, feuille, tabRentaCentres, tabMarcheCentre)

                'On affiche ces rentabilités centrées
                ExcelDialogue.affichageRentaCentrees(tabRentaCentres)
                'MsgBox("Fonctionnalité pas encore implémentée", 16)
        End Select

    End Sub

End Class
