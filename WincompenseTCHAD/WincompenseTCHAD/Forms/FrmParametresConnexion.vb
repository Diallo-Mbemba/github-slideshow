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
        IconesWU.Habiller(Me)
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

        ' Décochée par défaut, à dessein : le geste ordinaire est un dépannage, et la chaîne
        ' publiée avec l'application doit rester la référence. Une case déjà cochée ferait de
        ' l'exception la règle.
        chkConserver.Checked = False
        chkPropager.Checked = False
        AppliquerLaPortee()

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

    Private Sub chkConserver_CheckedChanged(sender As Object, e As EventArgs) Handles chkConserver.CheckedChanged
        AppliquerLaPortee()
    End Sub

    ''' <summary>
    ''' Rend actifs ou inactifs les réglages qui n'ont de sens que si l'on conserve.
    '''
    ''' Propager à tous les postes suppose d'écrire quelque part : sans conservation, le
    ''' fichier partagé n'a pas lieu d'être touché, et laisser la case accessible ferait
    ''' croire qu'un réglage temporaire peut engager la banque entière.
    ''' </summary>
    Private Sub AppliquerLaPortee()

        Dim conserve As Boolean = chkConserver.Checked

        txtPartage.Enabled = conserve
        chkPropager.Enabled = conserve

        If Not conserve Then chkPropager.Checked = False
    End Sub

    ''' <summary>
    ''' Enregistre pour cette session seulement : rien n'est écrit, et le poste reprendra sa
    ''' chaîne habituelle au prochain démarrage.
    ''' </summary>
    Private Function RetenirPourLaSession() As Boolean

        Dim chaine As String

        If chkChaineComplete.Checked Then
            chaine = txtChaine.Text.Replace(vbCr, " ").Replace(vbLf, " ").Trim()
            If chaine.Length = 0 Then
                MessageBox.Show("Collez d'abord la chaîne de connexion.", "Réglage incomplet",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return False
            End If
        Else
            If txtServeur.Text.Trim().Length = 0 Then
                MessageBox.Show("Indiquez d'abord le serveur.", "Réglage incomplet",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return False
            End If
            chaine = ConfigurationWU.ChaineDepuis(txtServeur.Text, txtBase.Text, CInt(nudDelai.Value))
        End If

        ConfigurationWU.ForcerPourCetteSession(chaine)
        CalendrierWU.Oublier()

        MessageBox.Show(
            "Le réglage vaut pour cette session seulement." & Environment.NewLine & Environment.NewLine &
            "Rien n'a été écrit sur ce poste : au prochain démarrage, l'application reprendra la" &
            Environment.NewLine &
            "chaîne publiée avec elle. Pour un changement durable, c'est cette chaîne qu'il faut" &
            Environment.NewLine &
            "corriger, puis republier l'application.",
            "Réglage temporaire", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Return True
    End Function

    ''' <summary>
    ''' Fait confirmer un réglage conservé sur le poste.
    '''
    ''' Un réglage écrit ici l'emporte sur la chaîne publiée avec l'application — et pour
    ''' toujours. Le poste cesse alors de suivre les republications, sans que rien ne le
    ''' signale ailleurs que sur cet écran. Cela peut être voulu ; cela ne doit pas être subi.
    ''' </summary>
    Private Function ConfirmerLaConservation() As Boolean

        Return MessageBox.Show(
            "Ce réglage sera CONSERVÉ sur ce poste." & Environment.NewLine & Environment.NewLine &
            "Il l'emportera désormais sur la chaîne publiée avec l'application : ce poste ne" &
            Environment.NewLine &
            "suivra plus les republications, jusqu'à ce que quelqu'un revienne le défaire ici." &
            Environment.NewLine & Environment.NewLine &
            "Pour un simple dépannage, décochez « Conserver » : le réglage vaudra le temps de" &
            Environment.NewLine &
            "la session, et le poste restera aligné sur la version publiée." & Environment.NewLine & Environment.NewLine &
            "Conserver tout de même ?",
            "Réglage durable", MessageBoxButtons.YesNo, MessageBoxIcon.Warning,
            MessageBoxDefaultButton.Button2) = DialogResult.Yes
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
    ''' <summary>
    ''' La chaîne décrite par l'écran en ce moment, prête à être essayée.
    '''
    ''' Le test et la préparation de la base doivent porter sur EXACTEMENT la même chaîne :
    ''' éprouver l'une et agir sur l'autre reviendrait à ne rien avoir éprouvé.
    ''' </summary>
    ''' <returns>Chaîne vide si l'écran n'est pas en état d'en décrire une.</returns>
    Private Function ChaineDeLEcran(ByRef message As String) As String

        If chkChaineComplete.Checked Then

            If txtChaine.Text.Trim().Length = 0 Then
                message = "Collez d'abord la chaîne de connexion fournie."
                Return String.Empty
            End If

            ' Sur une ligne, comme elle sera écrite : tester autre chose que ce qu'on
            ' enregistre n'aurait aucune valeur.
            Dim surUneLigne As String = txtChaine.Text.Replace(vbCr, " ").Replace(vbLf, " ").Trim()

            ' Une chaîne déjà enregistrée n'affiche plus son mot de passe : il est chiffré à
            ' part. On le remet pour l'essai, sinon le test échouerait là où l'application
            ' réussit — le pire des verdicts, celui qui envoie chercher une panne ailleurs.
            Return ConfigurationWU.ChaineEssayable(surUneLigne)
        End If

        If txtServeur.Text.Trim().Length = 0 Then
            message = "Indiquez d'abord le serveur."
            Return String.Empty
        End If

        Return ConfigurationWU.ChaineDepuis(txtServeur.Text, txtBase.Text, CInt(nudDelai.Value))
    End Function

    Private Function Tester(ByRef message As String) As Boolean

        Dim chaine As String = ChaineDeLEcran(message)
        If chaine.Length = 0 Then Return False

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

            ' Un message brut de SQL Server est exact mais inexploitable : en anglais, et
            ' muet sur ce qu'il reste à faire. Quand le diagnostic sait répondre, c'est lui
            ' qui parle, et le libellé de l'écran ne garde que la première ligne.
            Dim diagnostic As String = DiagnosticSqlWU.Expliquer(ex, chaine)

            If diagnostic.Length = 0 Then
                message = "Connexion refusée : " & ex.Message
                Return False
            End If

            message = PremiereLigne(diagnostic)
            FrmDiagnostic.Afficher(Me, "Connexion refusée par le serveur", diagnostic)
            Return False

        Catch ex As Exception
            message = "Connexion impossible : " & ex.Message
            Return False
        End Try
    End Function

    ''' <summary>
    ''' La première ligne d'un diagnostic, pour le libellé de l'écran : il fait deux lignes,
    ''' et le détail s'affiche dans sa propre fenêtre.
    ''' </summary>
    Private Shared Function PremiereLigne(texte As String) As String

        Dim fin As Integer = texte.IndexOfAny(New Char() {ControlChars.Cr, ControlChars.Lf})
        If fin < 0 Then Return texte
        Return texte.Substring(0, fin)
    End Function

#End Region

#Region "Préparation de la base"

    ''' <summary>
    ''' Constate l'état du compte sur le serveur, répare ce qui peut l'être, et rédige le reste.
    ''' 
    ''' Ce bouton répond à une demande simple : coller la chaîne de la banque et que tout suive.
    ''' Il y suit ce qu'il peut, et dit clairement ce qu'il ne peut pas — créer un accès à
    ''' SQL Server demande des droits d'administration du serveur, que le compte applicatif n'a
    ''' presque jamais. Le script produit alors n'est pas un pis-aller : c'est le livrable.
    ''' </summary>
    Private Sub btnPreparer_Click(sender As Object, e As EventArgs) Handles btnPreparer.Click

        Dim message As String = String.Empty
        Dim chaine As String = ChaineDeLEcran(message)

        If chaine.Length = 0 Then
            MessageBox.Show(Me, message, "Préparation de la base",
                            MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Cursor = Cursors.WaitCursor
        lblResultat.ForeColor = Drawing.SystemColors.ControlText
        lblResultat.Text = "Examen du serveur…"
        lblResultat.Refresh()

        Dim rapport As PreparationBaseWU.Rapport = PreparationBaseWU.Analyser(chaine)

        Cursor = Cursors.Default

        If rapport.RienAFaire Then
            lblResultat.ForeColor = Drawing.Color.DarkGreen
            lblResultat.Text = "Rien à préparer : ce compte a déjà accès à la base."
            MessageBox.Show(Me, Constat(rapport), "Préparation de la base",
                            MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        If rapport.PeutAgir AndAlso ConfirmerLaPreparation(rapport) Then
            AppliquerLaPreparation(chaine, rapport)
            Return
        End If

        RemettreLeScript(rapport)
    End Sub

    ''' <summary>Exécute ce qui manque, puis rend compte de chaque ordre.</summary>
    Private Sub AppliquerLaPreparation(chaine As String, rapport As PreparationBaseWU.Rapport)

        Dim journal As String = String.Empty

        Cursor = Cursors.WaitCursor
        Dim tout As Boolean = PreparationBaseWU.Appliquer(chaine, rapport, journal)
        Cursor = Cursors.Default

        lblResultat.ForeColor = If(tout, Drawing.Color.DarkGreen, Drawing.Color.Firebrick)
        lblResultat.Text = If(tout, "Base préparée.", "Préparation incomplète — voir le détail.")

        Dim texte As String = Constat(rapport) & Environment.NewLine & Environment.NewLine &
                              "--- Ce qui a été exécuté ---" & Environment.NewLine & Environment.NewLine &
                              journal

        ' Ce qui n'a pas pu être fait reste à faire par quelqu'un : le script accompagne donc
        ' le journal, plutôt que d'obliger à rouvrir l'écran pour l'obtenir.
        If Not tout Then
            texte &= Environment.NewLine & Environment.NewLine &
                     "--- À remettre à la banque pour le reste ---" & Environment.NewLine & Environment.NewLine &
                     PreparationBaseWU.ScriptPourLaBanque(rapport)
        End If

        FrmDiagnostic.Afficher(Me, "Préparation de la base", texte)
    End Sub

    ''' <summary>Affiche le script, précédé de ce qui l'explique.</summary>
    Private Sub RemettreLeScript(rapport As PreparationBaseWU.Rapport)

        lblResultat.ForeColor = Drawing.Color.Firebrick
        lblResultat.Text = "Droits insuffisants : script à remettre à la banque."

        Dim texte As String =
            Constat(rapport) & Environment.NewLine & Environment.NewLine &
            "Ce compte n'a pas les droits d'ouvrir lui-même la base — et c'est normal :" & Environment.NewLine &
            "une application comptable qui pourrait se donner des droits à elle-même n'aurait" & Environment.NewLine &
            "plus de contrôle d'accès du tout." & Environment.NewLine & Environment.NewLine &
            "Le script ci-dessous fait exactement ce qui manque. « Copier », puis courriel à" & Environment.NewLine &
            "l'informatique de la banque." & Environment.NewLine & Environment.NewLine &
            PreparationBaseWU.ScriptPourLaBanque(rapport)

        FrmDiagnostic.Afficher(Me, "Script à remettre à la banque", texte)
    End Sub

    ''' <summary>Ce que le serveur a répondu, en clair.</summary>
    Private Shared Function Constat(rapport As PreparationBaseWU.Rapport) As String

        If Not rapport.Joignable Then
            Return "Le serveur a refusé la connexion." & Environment.NewLine & Environment.NewLine &
                   rapport.MessageDeRefus
        End If

        Dim lignes As New System.Text.StringBuilder()

        lignes.AppendLine("Constat sur le serveur :")
        lignes.AppendLine()
        lignes.AppendLine("    compte          : " & rapport.Compte)
        lignes.AppendLine("    base visée      : " & rapport.Base)
        lignes.AppendLine("    base existe     : " & OuiNon(rapport.BaseExiste))
        lignes.AppendLine("    accès à la base : " & OuiNon(rapport.AccesALaBase))
        lignes.AppendLine("    tables en place : " & OuiNon(rapport.TablesPresentes))
        lignes.AppendLine("    rôles présents  : " & rapport.RolesPresents.ToString() & " sur 3")
        lignes.AppendLine("    rôle du compte  : " &
                          If(rapport.RoleDejaAccorde.Length > 0, rapport.RoleDejaAccorde, "aucun"))

        If rapport.EstAdministrateurDuServeur Then
            lignes.AppendLine()
            lignes.AppendLine("Ce compte administre le serveur : l'application peut tout faire elle-même.")
        End If

        Return lignes.ToString().TrimEnd()
    End Function

    Private Shared Function OuiNon(valeur As Boolean) As String
        Return If(valeur, "oui", "non")
    End Function

    ''' <summary>
    ''' Fait confirmer une modification des droits sur le serveur de la banque.
    ''' 
    ''' Ce n'est pas un réglage d'application : cela change qui a accès à quoi sur une base de
    ''' production. On ne le fait pas sans l'avoir demandé.
    ''' </summary>
    Private Function ConfirmerLaPreparation(rapport As PreparationBaseWU.Rapport) As Boolean

        Dim aFaire As New System.Text.StringBuilder()

        If Not rapport.AccesALaBase Then
            aFaire.AppendLine("    - ouvrir la base " & rapport.Base & " au compte " & rapport.Compte)
        End If

        If rapport.RoleDejaAccorde.Length = 0 Then
            aFaire.AppendLine("    - lui accorder le rôle " & PreparationBaseWU.ROLE_PROPOSE)
        End If

        If aFaire.Length = 0 Then Return False

        Return MessageBox.Show(Me,
            "L'application va modifier les droits sur le serveur :" & Environment.NewLine & Environment.NewLine &
            aFaire.ToString() & Environment.NewLine &
            "Cela engage la base de production. Si vous préférez que l'informatique de la" & Environment.NewLine &
            "banque le fasse, répondez Non : le script exact vous sera remis." & Environment.NewLine & Environment.NewLine &
            "Exécuter maintenant ?",
            "Préparation de la base", MessageBoxButtons.YesNo, MessageBoxIcon.Warning,
            MessageBoxDefaultButton.Button2) = DialogResult.Yes
    End Function

#End Region

#Region "Enregistrement"

    Private Sub btnEnregistrer_Click(sender As Object, e As EventArgs) Handles btnEnregistrer.Click

        ' Sans conservation, rien n'est écrit : la chaîne vaut pour cette session seulement.
        If Not chkConserver.Checked Then
            If RetenirPourLaSession() Then AfficherLaConfigurationEnService()
            Return
        End If

        If Not ConfirmerLaConservation() Then Return
        If Not ConfirmerLaPropagation() Then Return
        If Not ConfirmerLeMotDePasse() Then Return

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
    ''' Explique ce qu'il advient d'un mot de passe SQL Server, et le fait confirmer.
    '''
    ''' Il n'est plus refusé — la banque en fournit un — mais il ne s'écrit pas n'importe où.
    ''' Retiré de la chaîne avant toute écriture, il est chiffré par Windows pour CE poste
    ''' seulement (voir SecretWU), et le fichier partagé ne reçoit que le serveur, la base et
    ''' le nom du compte.
    '''
    ''' D'où l'avertissement sur la propagation : le serveur changera bien pour toute la
    ''' banque, mais un poste qui n'a jamais reçu le mot de passe ne se connectera pas pour
    ''' autant. Le taire ferait croire à un déploiement terminé qui ne l'est pas.
    ''' </summary>
    Private Function ConfirmerLeMotDePasse() As Boolean

        If Not chkChaineComplete.Checked Then Return True
        If Not ConfigurationWU.PorteUnMotDePasse(txtChaine.Text) Then Return True

        Dim texte As String =
            "Cette chaîne contient un mot de passe." & Environment.NewLine & Environment.NewLine &
            "Il sera retiré de la chaîne, puis chiffré par Windows pour ce poste : il" & Environment.NewLine &
            "n'apparaîtra en clair dans aucun fichier, et ne partira jamais sur le partage." & Environment.NewLine & Environment.NewLine &
            "Le chiffrement est lié à cette machine. Un fichier recopié ailleurs ne donne" & Environment.NewLine &
            "rien — mais il ne protège pas d'un programme lancé sur ce poste même."

        If chkPropager.Checked Then
            texte &= Environment.NewLine & Environment.NewLine &
                     "PROPAGATION : le partage recevra le serveur, la base et le nom du compte," & Environment.NewLine &
                     "jamais le mot de passe. Chaque autre poste devra le saisir une fois, ici." & Environment.NewLine &
                     "Sans quoi il verra « Login failed for user »."
        End If

        texte &= Environment.NewLine & Environment.NewLine & "Enregistrer ?"

        Return MessageBox.Show(texte, "Mot de passe du compte SQL Server",
                               MessageBoxButtons.YesNo, MessageBoxIcon.Information,
                               MessageBoxDefaultButton.Button1) = DialogResult.Yes
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
