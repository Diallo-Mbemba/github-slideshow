Option Strict On
Option Explicit On

Imports System.Configuration
Imports System.IO
Imports System.Text

''' <summary>
''' Trouve la chaîne de connexion à SQL Server, et sait la changer.
'''
''' La banque change souvent de serveur. Tant que la chaîne vivait dans le fichier de
''' configuration posé à côté de l'exécutable, en changer imposait une tournée dans les bureaux :
''' ce fichier est dans Program Files, donc protégé, il est propre à chaque poste, et une
''' réinstallation l'écrase.
'''
''' La chaîne est donc cherchée ailleurs, dans un ordre où le premier trouvé l'emporte :
'''
'''   1. la variable d'environnement WINCOMPENSE_CONNEXION — dépannage, poste de test ;
'''   2. le FICHIER PARTAGÉ, dont le chemin est inscrit sur le poste à l'installation :
'''      c'est la source de vérité, et le seul fichier à changer pour toute la banque ;
'''   3. le fichier local %PROGRAMDATA%\Wincompense\wincompense.config, rafraîchi à chaque
'''      lecture réussie du partage ;
'''   4. App.config, pour un poste de développement qui n'a rien installé ;
'''   5. la valeur compilée, dernier recours.
'''
''' Le rang 3 n'est pas un doublon : c'est ce qui permet de travailler le matin où le partage
''' est injoignable. Sans lui, une coupure réseau arrêterait la compense de toute la banque.
'''
''' Format des deux fichiers : une ligne CLE=VALEUR, le dièse ouvre un commentaire. Ni XML ni
''' registre — l'informatique de la banque doit pouvoir changer de serveur dans le Bloc-notes,
''' sans outil et sans risque de casser une balise.
''' </summary>
Public NotInheritable Class ConfigurationWU

    Private Sub New()
    End Sub

#Region "Emplacements et clés"

    ''' <summary>Nom du dossier de l'application sous %PROGRAMDATA%.</summary>
    Private Const DOSSIER As String = "Wincompense"

    ''' <summary>Fichier local : chemin du partage, et copie de secours de la connexion.</summary>
    Private Const FICHIER_LOCAL As String = "wincompense.config"

    Private Const VARIABLE_ENVIRONNEMENT As String = "WINCOMPENSE_CONNEXION"

    ''' <summary>Clé du fichier local portant le chemin du fichier partagé.</summary>
    Public Const CLE_PARTAGE As String = "PARTAGE"

    Public Const CLE_SERVEUR As String = "SERVEUR"
    Public Const CLE_BASE As String = "BASE"
    Public Const CLE_DELAI As String = "DELAI"

    ''' <summary>Chaîne complète, pour les cas exceptionnels : elle prime sur SERVEUR/BASE/DELAI.</summary>
    Public Const CLE_CHAINE As String = "CHAINE"

    Public Const BASE_PAR_DEFAUT As String = "GWC_WINCOMPENSE_ETD"
    Public Const DELAI_PAR_DEFAUT As Integer = 10

    ''' <summary>Nom de la section App.config conservée pour les postes de développement.</summary>
    Private Const SECTION_APP_CONFIG As String = BASE_PAR_DEFAUT

    ''' <summary>Dossier local de configuration : %PROGRAMDATA%\Wincompense.</summary>
    Public Shared ReadOnly Property DossierLocal As String
        Get
            Return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), DOSSIER)
        End Get
    End Property

    ''' <summary>Fichier local de configuration, qu'il existe ou non.</summary>
    Public Shared ReadOnly Property CheminLocal As String
        Get
            Return Path.Combine(DossierLocal, FICHIER_LOCAL)
        End Get
    End Property

    ''' <summary>
    ''' Chemin du fichier partagé, tel qu'il est inscrit sur ce poste. Chaîne vide si le poste
    ''' n'a pas été installé, ou si l'installation n'a pas désigné de partage.
    ''' </summary>
    Public Shared ReadOnly Property CheminDuPartage As String
        Get
            Return LireCle(LireFichier(CheminLocal), CLE_PARTAGE)
        End Get
    End Property

#End Region

#Region "Résolution de la chaîne de connexion"

    ''' <summary>Chaîne retenue, calculée une fois puis conservée le temps de l'exécution.</summary>
    Private Shared _chaine As String = String.Empty

    ''' <summary>Provenance de la chaîne retenue, en clair, pour l'écran d'administration.</summary>
    Private Shared _origine As String = String.Empty

    ''' <summary>
    ''' Chaîne de connexion à employer. Ne lève jamais : un poste mal configuré doit atteindre
    ''' l'écran de connexion et y lire un message, non se fermer sur une exception.
    ''' </summary>
    Public Shared Function ChaineDeConnexion() As String

        If _chaine.Length > 0 Then Return _chaine

        Resoudre()
        Return _chaine
    End Function

    ''' <summary>
    ''' D'où vient la chaîne en service : « fichier partagé », « copie locale », etc.
    ''' Affiché à l'administrateur, qui doit pouvoir constater que le partage est bien lu.
    ''' </summary>
    Public Shared Function Origine() As String

        If _chaine.Length = 0 Then Resoudre()
        Return _origine
    End Function

    ''' <summary>
    ''' Oblige à tout relire au prochain appel. À employer après un enregistrement, pour que le
    ''' changement prenne effet sans quitter l'application.
    ''' </summary>
    Public Shared Sub Oublier()
        _chaine = String.Empty
        _origine = String.Empty
    End Sub

    Private Shared Sub Resoudre()

        ' 1. Variable d'environnement : elle passe avant tout, et n'engage que ce poste.
        Dim environnement As String = Nothing
        Try
            environnement = Environment.GetEnvironmentVariable(VARIABLE_ENVIRONNEMENT)
        Catch
            ' Un accès refusé aux variables d'environnement ne doit pas arrêter la résolution.
        End Try

        If Not String.IsNullOrWhiteSpace(environnement) Then
            _chaine = environnement.Trim()
            _origine = $"variable d'environnement {VARIABLE_ENVIRONNEMENT}"
            Return
        End If

        Dim local As Dictionary(Of String, String) = LireFichier(CheminLocal)
        Dim partage As String = LireCle(local, CLE_PARTAGE)

        ' Distingue un partage qui ne répond pas d'un partage qui répond mal : les deux
        ' renvoient à la copie locale, mais ne se corrigent pas au même endroit.
        Dim partageLu As Boolean = False

        ' 2. Fichier partagé : la source de vérité.
        If partage.Length > 0 Then

            Dim valeurs As Dictionary(Of String, String) = LireFichier(partage)
            partageLu = valeurs.Count > 0

            If partageLu Then
                Dim chaine As String = Construire(valeurs)

                If chaine.Length > 0 Then
                    _chaine = chaine
                    _origine = $"fichier partagé {partage}"

                    ' La copie locale est rafraîchie maintenant, pendant que le partage répond :
                    ' c'est elle qui fera tourner le poste le jour où il ne répondra plus.
                    RafraichirLaCopieLocale(local, valeurs)
                    Return
                End If
            End If
        End If

        ' 3. Copie locale.
        Dim chaineLocale As String = Construire(local)

        If chaineLocale.Length > 0 Then
            _chaine = chaineLocale
            If partage.Length = 0 Then
                _origine = "fichier local de ce poste"
            ElseIf partageLu Then
                _origine = $"copie locale — {partage} est lisible mais n'indique aucun serveur"
            Else
                _origine = $"copie locale — {partage} injoignable"
            End If

            Return
        End If

        ' 4. App.config, pour un poste de développement.
        Try
            Dim section As ConnectionStringSettings = ConfigurationManager.ConnectionStrings(SECTION_APP_CONFIG)

            If section IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(section.ConnectionString) Then
                _chaine = section.ConnectionString
                _origine = "App.config de l'application"
                Return
            End If
        Catch
            ' Un fichier de configuration illisible retombe simplement sur la valeur par défaut.
        End Try

        ' 5. Dernier recours.
        _chaine = ChaineDepuis(".\SQLEXPRESS", BASE_PAR_DEFAUT, DELAI_PAR_DEFAUT)
        _origine = "valeur par défaut (aucune configuration trouvée)"
    End Sub

#End Region

#Region "Construction de la chaîne"

    ''' <summary>
    ''' Chaîne de connexion formée à partir d'un serveur et d'une base, en authentification
    ''' Windows intégrée — la banque n'emploie pas d'identifiant SQL. Aucun mot de passe ne
    ''' circule donc, ce qui est précisément ce qui permet de poser la configuration sur un
    ''' partage lisible par tous.
    ''' </summary>
    Public Shared Function ChaineDepuis(serveur As String, base As String, delai As Integer) As String

        Dim constructeur As New System.Data.SqlClient.SqlConnectionStringBuilder() With {
            .DataSource = If(serveur, String.Empty).Trim(),
            .InitialCatalog = If(String.IsNullOrWhiteSpace(base), BASE_PAR_DEFAUT, base.Trim()),
            .IntegratedSecurity = True,
            .ConnectTimeout = If(delai > 0, delai, DELAI_PAR_DEFAUT)
        }

        Return constructeur.ConnectionString
    End Function

    ''' <summary>
    ''' Chaîne décrite par un jeu de clés. CHAINE prime, s'il est renseigné ; sinon la chaîne
    ''' est bâtie sur SERVEUR, BASE et DELAI. Chaîne vide si le serveur n'est pas renseigné :
    ''' sans serveur il n'y a rien à tenter, et mieux vaut passer au rang suivant.
    ''' </summary>
    Private Shared Function Construire(valeurs As Dictionary(Of String, String)) As String

        Dim complete As String = LireCle(valeurs, CLE_CHAINE)
        If complete.Length > 0 Then Return complete

        Dim serveur As String = LireCle(valeurs, CLE_SERVEUR)
        If serveur.Length = 0 Then Return String.Empty

        Dim base As String = LireCle(valeurs, CLE_BASE)

        Dim delai As Integer
        If Not Integer.TryParse(LireCle(valeurs, CLE_DELAI), delai) Then delai = DELAI_PAR_DEFAUT

        Return ChaineDepuis(serveur, base, delai)
    End Function

#End Region

#Region "Lecture et écriture des fichiers"

    ''' <summary>
    ''' Lit un fichier CLE=VALEUR. Retourne un dictionnaire vide si le fichier est absent,
    ''' illisible ou sur un partage injoignable — jamais d'exception : chacun des cinq rangs
    ''' doit pouvoir échouer sans emporter les suivants.
    ''' </summary>
    Private Shared Function LireFichier(chemin As String) As Dictionary(Of String, String)

        Dim valeurs As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

        If String.IsNullOrWhiteSpace(chemin) Then Return valeurs

        Try
            If Not File.Exists(chemin) Then Return valeurs

            For Each ligne As String In File.ReadAllLines(chemin, Encoding.UTF8)

                Dim nette As String = ligne.Trim()
                If nette.Length = 0 OrElse nette.StartsWith("#", StringComparison.Ordinal) Then Continue For

                Dim separateur As Integer = nette.IndexOf("="c)
                If separateur <= 0 Then Continue For

                Dim cle As String = nette.Substring(0, separateur).Trim()
                Dim valeur As String = nette.Substring(separateur + 1).Trim()

                If cle.Length > 0 Then valeurs(cle) = valeur
            Next

        Catch
            ' Partage injoignable, droit refusé, fichier verrouillé : le rang suivant prendra
            ' la main. Signaler ici n'aurait pas de sens — la résolution n'a pas encore échoué.
        End Try

        Return valeurs
    End Function

    ''' <summary>Valeur d'une clé, ou chaîne vide.</summary>
    Private Shared Function LireCle(valeurs As Dictionary(Of String, String), cle As String) As String

        Dim valeur As String = Nothing
        If valeurs Is Nothing OrElse Not valeurs.TryGetValue(cle, valeur) Then Return String.Empty

        Return If(valeur, String.Empty).Trim()
    End Function

    ''' <summary>
    ''' Réécrit la copie locale avec ce que le partage vient de donner, en conservant le chemin
    ''' du partage lui-même. Un échec est ignoré : le poste tourne déjà sur la chaîne lue, et
    ''' refuser de démarrer parce qu'un cache n'a pas pu s'écrire serait absurde.
    ''' </summary>
    Private Shared Sub RafraichirLaCopieLocale(local As Dictionary(Of String, String),
                                               valeursDuPartage As Dictionary(Of String, String))

        Try
            Dim fusion As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

            For Each paire As KeyValuePair(Of String, String) In valeursDuPartage
                fusion(paire.Key) = paire.Value
            Next

            ' Le chemin du partage n'appartient qu'au poste : le partage ne se désigne pas
            ' lui-même, sans quoi on ne saurait plus où chercher le jour où il change.
            fusion(CLE_PARTAGE) = LireCle(local, CLE_PARTAGE)

            EcrireFichier(CheminLocal, fusion, "Copie locale de secours - regeneree automatiquement")

        Catch
            ' Volontairement silencieux : voir le commentaire ci-dessus.
        End Try
    End Sub

    ''' <summary>
    ''' Écrit un fichier CLE=VALEUR, en créant le dossier au besoin.
    ''' </summary>
    ''' <param name="chemin">Fichier à écrire.</param>
    ''' <param name="valeurs">Clés à inscrire.</param>
    ''' <param name="titre">Ligne de commentaire placée en tête.</param>
    Private Shared Sub EcrireFichier(chemin As String, valeurs As Dictionary(Of String, String), titre As String)

        Dim dossier As String = Path.GetDirectoryName(chemin)
        If Not String.IsNullOrEmpty(dossier) AndAlso Not Directory.Exists(dossier) Then
            Directory.CreateDirectory(dossier)
        End If

        Dim contenu As New StringBuilder()
        contenu.AppendLine("# Wincompense TCHAD - " & titre)
        contenu.AppendLine("# Une ligne CLE=VALEUR. Le caractere # ouvre un commentaire.")
        contenu.AppendLine("# Ecrit le " & Date.Now.ToString("dd/MM/yyyy HH:mm"))
        contenu.AppendLine()

        For Each paire As KeyValuePair(Of String, String) In valeurs
            If String.IsNullOrWhiteSpace(paire.Value) Then Continue For
            contenu.AppendLine($"{paire.Key}={paire.Value}")
        Next

        File.WriteAllText(chemin, contenu.ToString(), Encoding.UTF8)
    End Sub

#End Region

#Region "Enregistrement depuis l'écran d'administration"

    ''' <summary>
    ''' Enregistre le serveur et la base, sur ce poste et — si demandé — sur le partage.
    '''
    ''' L'écriture sur le partage est faite EN PREMIER : c'est elle qui peut échouer, faute de
    ''' droits, et il vaut mieux ne rien avoir changé du tout que d'avoir un poste réglé sur un
    ''' serveur que les autres ignorent.
    ''' </summary>
    ''' <param name="serveur">Nom ou adresse de l'instance SQL Server.</param>
    ''' <param name="base">Nom de la base.</param>
    ''' <param name="delai">Délai de connexion, en secondes.</param>
    ''' <param name="cheminPartage">Fichier partagé. Chaîne vide pour n'écrire que localement.</param>
    ''' <param name="ecrireSurLePartage">Vrai pour propager le changement à toute la banque.</param>
    ''' <param name="messageErreur">Motif de l'échec, le cas échéant.</param>
    ''' <returns>Vrai si tout ce qui était demandé a été écrit.</returns>
    Public Shared Function Enregistrer(serveur As String, base As String, delai As Integer,
                                       cheminPartage As String, ecrireSurLePartage As Boolean,
                                       ByRef messageErreur As String) As Boolean

        messageErreur = String.Empty

        If String.IsNullOrWhiteSpace(serveur) Then
            messageErreur = "Le serveur n'est pas renseigné : sans lui, l'application n'a rien à joindre."
            Return False
        End If

        Dim valeurs As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
            {CLE_SERVEUR, serveur.Trim()},
            {CLE_BASE, If(String.IsNullOrWhiteSpace(base), BASE_PAR_DEFAUT, base.Trim())},
            {CLE_DELAI, If(delai > 0, delai, DELAI_PAR_DEFAUT).ToString()}
        }

        If ecrireSurLePartage Then

            If String.IsNullOrWhiteSpace(cheminPartage) Then
                messageErreur = "Aucun fichier partagé n'est indiqué : impossible de propager le changement."
                Return False
            End If

            Try
                EcrireFichier(cheminPartage, valeurs, "Connexion commune a tous les postes")
            Catch ex As Exception
                messageErreur = $"Écriture impossible sur {cheminPartage} : {ex.Message}" & Environment.NewLine &
                                "Vérifiez que vous avez le droit d'écrire sur ce partage. Rien n'a été modifié."
                Return False
            End Try
        End If

        Try
            Dim local As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

            For Each paire As KeyValuePair(Of String, String) In valeurs
                local(paire.Key) = paire.Value
            Next

            local(CLE_PARTAGE) = If(cheminPartage, String.Empty).Trim()

            EcrireFichier(CheminLocal, local, "Configuration de ce poste")

        Catch ex As Exception
            messageErreur = $"Écriture impossible dans {CheminLocal} : {ex.Message}"
            Return False
        End Try

        ' La prochaine lecture repart de zéro : le changement vaut sans quitter l'application.
        Oublier()
        Return True
    End Function

#End Region

End Class
