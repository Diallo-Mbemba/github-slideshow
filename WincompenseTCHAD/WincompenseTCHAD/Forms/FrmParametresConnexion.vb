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

        ' Une chaîne complète déjà posée doit se voir : sans cela, l'écran afficherait un
        ' serveur et une base décomposés, et les enregistrer écraserait la chaîne — avec les
        ' mots-clés qu'elle porte et que trois champs ne savent pas exprimer.
        Dim complete As String = ConfigurationWU.ChaineComplete
        chkChaineComplete.Checked = complete.Length > 0
        txtChaine.Text = complete
        AppliquerLeMode()

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

    Private Sub chkChaineComplete_CheckedChanged(sender As Object, e As EventArgs) Handles chkChaineComplete.CheckedChanged
        AppliquerLeMode()
    End Sub

    ''' <summary>
    ''' Rend actif l'un des deux modes et grise l'autre.
    '''
    ''' Les deux ne coexistent jamais : la clé CHAINE prime sur SERVEUR, BASE et DELAI, et
    ''' laisser les deux saisies ouvertes ferait croire que la seconde compte encore.
    ''' </summary>
    Private Sub AppliquerLeMode()

        Dim complete As Boolean = chkChaineComplete.Checked

        txtChaine.Enabled = complete

        txtServeur.Enabled = Not complete
        txtBase.Enabled = Not complete
        nudDelai.Enabled = Not complete

        If complete AndAlso txtChaine.Text.Trim().Length = 0 Then
            ' Départ de saisie : la chaîne en service évite de tout retaper quand la banque n'a
            ' changé qu'un mot-clé.
            txtChaine.Text = WURepository.ObtenirChaineConnexion()
        End If
    End Sub

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

        Dim chaine As String

        If chkChaineComplete.Checked Then

            If txtChaine.Text.Trim().Length = 0 Then
                message = "Collez d'abord la chaîne de connexion fournie."
                Return False
            End If

            ' Sur une ligne, comme elle sera écrite : tester autre chose que ce qu'on
            ' enregistre n'aurait aucune valeur.
            chaine = txtChaine.Text.Replace(vbCr, " ").Replace(vbLf, " ").Trim()
        Else

            If txtServeur.Text.Trim().Length = 0 Then
                message = "Indiquez d'abord le serveur."
                Return False
            End If

            chaine = ConfigurationWU.ChaineDepuis(txtServeur.Text, txtBase.Text, CInt(nudDelai.Value))
        End If

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

        If Not ConfirmerLAbsenceDeMotDePasse() Then Return

        Dim messageErreur As String = String.Empty

        Cursor = Cursors.WaitCursor

        Dim enregistre As Boolean

        If chkChaineComplete.Checked Then
            enregistre = ConfigurationWU.EnregistrerChaineComplete(
                txtChaine.Text, txtPartage.Text.Trim(), chkPropager.Checked, messageErreur)
        Else
            enregistre = ConfigurationWU.Enregistrer(
                txtServeur.Text, txtBase.Text, CInt(nudDelai.Value),
                txtPartage.Text.Trim(), chkPropager.Checked, messageErreur)
        End If

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
    ''' Refuse de poser un mot de passe SQL sur le partage, et le fait confirmer ailleurs.
    '''
    ''' Le fichier partagé est lisible par tous les utilisateurs de l'application : c'est ce qui
    ''' permet à un changement de serveur de valoir pour tout le monde. Un mot de passe écrit là
    ''' serait donc lisible par tous, en clair. La banque est en authentification Windows, où la
    ''' question ne se pose pas ; une chaîne fournie par un tiers peut néanmoins en contenir un.
    ''' </summary>
    Private Function ConfirmerLAbsenceDeMotDePasse() As Boolean

        If Not chkChaineComplete.Checked Then Return True

        Dim chaine As String = txtChaine.Text
        Dim porteUnMotDePasse As Boolean =
            chaine.IndexOf("Password", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
            chaine.IndexOf("Pwd", StringComparison.OrdinalIgnoreCase) >= 0

        If Not porteUnMotDePasse Then Return True

        If chkPropager.Checked Then
            MessageBox.Show(
                "Cette chaîne contient un mot de passe, et le fichier partagé est lisible par tous" &
                Environment.NewLine &
                "les utilisateurs de l'application : il y serait en clair." & Environment.NewLine & Environment.NewLine &
                "Demandez à la banque une chaîne en authentification Windows " &
                "(Integrated Security=True)," & Environment.NewLine &
                "ou décochez la propagation pour ne régler que ce poste.",
                "Mot de passe sur le partage", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End If

        Return MessageBox.Show(
            "Cette chaîne contient un mot de passe. Il sera écrit en clair dans le fichier de" &
            Environment.NewLine &
            "configuration de ce poste." & Environment.NewLine & Environment.NewLine &
            "L'authentification Windows évite ce risque, et c'est le mode retenu par la banque." &
            Environment.NewLine & Environment.NewLine &
            "Enregistrer tout de même sur ce poste ?",
            "Mot de passe en clair", MessageBoxButtons.YesNo, MessageBoxIcon.Warning,
            MessageBoxDefaultButton.Button2) = DialogResult.Yes
    End Function

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
