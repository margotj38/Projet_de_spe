﻿Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices

Public Class SelectionDatesEv

    Private donneesPreT As Integer ' 0 => prix; 1 => rentabilités 
    Private nomFeuille As String

    'accesseur sur donneesPreTraitement
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
    Public Sub New(donnees As Integer)
        InitializeComponent()
        donneesPreTraitement = donnees
    End Sub

    Private Sub SelectionDatesEv_Load(sender As Object, e As EventArgs) Handles Me.Load
        Dim excelApp As Excel.Application = Nothing

        ' Create an Excel App
        Try
            excelApp = Marshal.GetActiveObject("Excel.Application")
        Catch ex As COMException
            ' An exception is thrown if there is not an open excel instance.                    
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

    Private Sub lancementPreT_Click(sender As Object, e As EventArgs) Handles lancementPreT.Click
        'On récupère la plage des dates et la feuille sur laquelle elle est
        Dim plage As String = ""
        Dim feuilleDonnees As String = Me.nomFeuille
        Dim feuilleDates As String = ""
        Utilitaires.recupererFeuillePlage(Me.datesEvRefEdit.Address, feuilleDates, plage)

        Select Case donneesPreTraitement
            Case 0
                'Si on doit traiter des prix
                'On centre les cours des entreprises et du marché
                Dim tabPrixCentres(,) As Double = Nothing
                Dim tabMarcheCentre(,) As Double = Nothing
                UtilitaireRentabilites.donneesCentrees(plage, feuilleDates, feuilleDonnees, tabPrixCentres, tabMarcheCentre)

                'On calcule les rentabilités
                Dim tabRenta(tabPrixCentres.GetUpperBound(0) - 1, tabPrixCentres.GetUpperBound(1)) As Double
                Dim tabRentaMarche(tabMarcheCentre.GetUpperBound(0) - 1, tabMarcheCentre.GetUpperBound(1)) As Double
                Dim maxPrixAbsent As Integer
                UtilitaireRentabilites.calculTabRenta(tabPrixCentres, tabMarcheCentre, tabRenta, tabRentaMarche, maxPrixAbsent)

                'On stocke le tableaux des rentabilités de marché et des entreprises dont on va avoir besoin
                'PB : où ? Dans nouveau module rentabilité ?
                UtilitaireRentabilites.tabRentaMarche = tabRentaMarche
                UtilitaireRentabilites.tabRenta = tabRenta
                'Idem pour maxPrixAbsent
                UtilitaireRentabilites.maxPrixAbs = maxPrixAbsent

                'On affiche ces rentabilités centrées
                ExcelDialogue.affichageRentaCentrees(tabRenta)
            Case 1
                'Si on doit traiter des rentabilités
                'On centre les rentabilités (2ème colonne : marché)
                Dim tabRentaCentrees(,) As Double = Nothing
                Dim tabMarcheCentre(,) As Double = Nothing

                UtilitaireRentabilites.donneesCentrees(plage, feuilleDates, feuilleDonnees, tabRentaCentrees, tabMarcheCentre)

                'On stocke le tableaux des rentabilités de marché dont on va avoir besoin
                'PB : où ? Dans nouveau module rentabilité ?
                UtilitaireRentabilites.tabRentaMarche = tabMarcheCentre
                UtilitaireRentabilites.tabRenta = tabRentaCentrees

                'Calcul de maxPrixAbsent
                Dim maxPrixAbsent As Integer = UtilitaireRentabilites.calculMaxPrixAbs(tabRentaCentrees)
                UtilitaireRentabilites.maxPrixAbs = maxPrixAbsent

                'On affiche les rentabilités centrées
                ExcelDialogue.affichageRentaCentrees(tabRentaCentrees)


                'MsgBox("Fonctionnalité pas encore implémentée", 16)
        End Select

    End Sub

    Private Sub nomFeuilleBox_TextChanged(sender As Object, e As EventArgs) Handles nomFeuilleBox.TextChanged
        nomFeuille = nomFeuilleBox.Text
    End Sub
End Class