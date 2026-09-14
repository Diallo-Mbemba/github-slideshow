Option Strict On
Option Explicit On

Imports System.Data.SqlClient

''' <summary>
''' Écran d'administration de la connexion à la base.
'''
''' La banque change souvent de serveur. Cet écran permet de le faire sans réinstaller, sans
''' recompiler et sans éditer de fichier à la main — et surtout, en une seule fois pour tous les
''' postes lorsque le réglage est propagé sur le partage.
'''
''' Deux garde-fous y sont posés. Le test de connexion est offert AVANT l'enregistrement, parce
''' qu'un serveur mal orthographié propagé à toute la banque arrête tout le monde. Et
''' l'enregistrement sur le partage est tenté avant l'enregistrement local : mieux vaut n'avoir
''' rien changé que d'avoir un poste réglé sur un serveur que les autres ignorent.
''' </summary>
Public Class FrmParametresConnexion

    Public Sub New()

        InitializeComponent()
    End Sub

#Region "Ouverture"

    Private Sub FrmParametresConnexion_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        AfficherLaConfigurationEnService()
    End Sub

    ''' <summary>
    ''' Remplit l'écran avec ce que l'application emploie réellement en ce moment — et dit d'où
    ''' cela vient. Un administrateur qui croit lire le partage alors qu'il lit une copie locale
    ''' périmée changerait un fichier sans effet.
    ''' </summary>
    Private Sub AfficherLaConfigurationEnService()

        Dim constructeur As SqlConnectionStringBuilder = Decomposer(WURepository.ObtenirChaineConnexion())

        txtServeur.Text = constructeur.DataSource
        txtBase.Text = If(String.IsNullOrWhiteSpace(constructeur.InitialCatalog),
                          ConfigurationWU.BASE_PAR_DEFAUT, constructeur.InitialCatalog)

        nudDelai.Value = BornerLeDelai(constructeur.ConnectTimeout)

        txtPartage.Text = ConfigurationWU.CheminDuPartage
        chkPropager.Checked = txtPartage.Text.Length > 0

        lblOrigine.Text = $"Chaîne en service : {ConfigurationWU.Origine()}." & Environment.NewLine &
                          $"Configuration de ce poste : {ConfigurationWU.CheminLocal}"
    End Sub

    ''' <summary>
    ''' Décompose une chaîne de connexion. Une chaîne illisible — écrite à la main dans le
    ''' fichier partagé, par exemple — rend un constructeur vide plutôt que de faire échouer
    ''' l'ouverture de l'écran : c'est justement ici qu'on vient la corriger.
    ''' </summary>
    Private Shared Function Decomposer(chaine As String) As SqlConnectionStringBuilder

        Try
            Return New SqlConnectionStringBuilder(chaine)
        Catch
            Return New SqlConnectionStringBuilder()
        End Try
    End Function

    ''' <summary>Ramène un délai dans les bornes du compteur, qui refuserait une valeur hors limites.</summary>
    Private Function BornerLeDelai(secondes As Integer) As Decimal

        Dim valeur As Decimal = CDec(secondes)

        If valeur < nudDelai.Minimum Then Return nudDelai.Minimum
        If valeur > nudDelai.Maximum Then Return nudDelai.Maximum

        Return valeur
    End Function

#End Region

#Region "Test de la connexion"

    Private Sub btnTester_Click(sender As Object, e As EventArgs) Handles btnTester.Click

        Dim message As String = String.Empty

        Cursor = Cursors.WaitCursor
        lblResultat.ForeColor = Drawing.SystemColors.ControlText
        lblResultat.Text = "Connexion en cours…"
        lblResultat.Refresh()

        Dim reussi As Boolean = Tester(message)

        Cursor = Cursors.Default
        lblResultat.ForeColor = If(reussi, Drawing.Color.DarkGreen, Drawing.Color.Firebrick)
        lblResultat.Text = message
    End Sub

    ''' <summary>
    ''' Ouvre une connexion sur les valeurs SAISIES — et non sur celles en service — puis vérifie
    ''' que la base répond. Ouvrir ne suffit pas : SQL Server accepte la connexion avant de dire
    ''' que la base n'existe pas, et un serveur joignable avec une mauvaise base passerait pour
    ''' un succès.
    ''' </summary>
    Private Function Tester(ByRef message As String) As Boolean

        If txtServeur.Text.Trim().Length = 0 Then
            message = "Indiquez d'abord le serveur."
            Return False
        End If

        Dim chaine As String = ConfigurationWU.ChaineDepuis(txtServeur.Text, txtBase.Text, CInt(nudDelai.Value))

        Try
            Using connexion As New SqlConnection(chaine)
                connexion.Open()

                Using commande As New SqlCommand("SELECT DB_NAME(), SUSER_SNAME()", connexion)
                    Using lecteur As SqlDataReader = commande.ExecuteReader()

                        If Not lecteur.Read() Then
                            message = "Le serveur répond, mais n'a rien renvoyé : base inattendue."
                            Return False
                        End If

                        message = $"Connexion réussie — base {lecteur.GetString(0)}, " &
                                  $"ouverte au nom de {lecteur.GetString(1)}."
                        Return True
                    End Using
                End Using
            End Using

        Catch ex As SqlException
            message = "Connexion refusée : " & ex.Message
            Return False

        Catch ex As Exception
            message = "Connexion impossible : " & ex.Message
            Return False
        End Try
    End Function

#End Region

#Region "Enregistrement"

    Private Sub btnEnregistrer_Click(sender As Object, e As EventArgs) Handles btnEnregistrer.Click

        If Not ConfirmerLaPropagation() Then Return

        Dim messageErreur As String = String.Empty

        Cursor = Cursors.WaitCursor

        Dim enregistre As Boolean = ConfigurationWU.Enregistrer(
            txtServeur.Text, txtBase.Text, CInt(nudDelai.Value),
            txtPartage.Text.Trim(), chkPropager.Checked, messageErreur)

        Cursor = Cursors.Default

        If Not enregistre Then
            MessageBox.Show(messageErreur, "Enregistrement impossible",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        ' Les jours fériés ont été lus sur l'ancien serveur : les garder ferait travailler
        ' l'application sur le calendrier d'une autre base.
        CalendrierWU.Oublier()

        MessageBox.Show(
            If(chkPropager.Checked,
               "Le réglage est enregistré sur ce poste et sur le fichier partagé." & Environment.NewLine &
               "Les autres postes le prendront à leur prochain démarrage.",
               "Le réglage est enregistré sur ce poste seulement." & Environment.NewLine &
               "Les autres postes continuent de lire l'ancien serveur.") & Environment.NewLine & Environment.NewLine &
            "Fermez et rouvrez l'application pour que les écrans déjà ouverts en tiennent compte.",
            "Connexion enregistrée", MessageBoxButtons.OK, MessageBoxIcon.Information)

        AfficherLaConfigurationEnService()
    End Sub

    ''' <summary>
    ''' Fait confirmer un changement qui engage toute la banque, et rappelle qu'il vaut mieux
    ''' avoir testé. Un serveur mal orthographié propagé à tous arrête tout le monde, et se
    ''' corrige alors depuis un poste qui ne se connecte plus.
    ''' </summary>
    Private Function ConfirmerLaPropagation() As Boolean

        If Not chkPropager.Checked Then Return True

        Return MessageBox.Show(
            $"Ce réglage sera appliqué à TOUS les postes de la banque, par le fichier :" &
            Environment.NewLine & Environment.NewLine &
            $"    {txtPartage.Text.Trim()}" & Environment.NewLine & Environment.NewLine &
            "Un serveur mal saisi arrêterait tout le monde. Avez-vous testé la connexion ?" &
            Environment.NewLine & Environment.NewLine &
            "Enregistrer pour tous les postes ?",
            "Changement pour toute la banque", MessageBoxButtons.YesNo, MessageBoxIcon.Warning,
            MessageBoxDefaultButton.Button2) = DialogResult.Yes
    End Function

#End Region

End Class
