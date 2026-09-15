Option Strict On
Option Explicit On

Imports System.IO
Imports System.Reflection
Imports System.Text

''' <summary>
''' Annonce aux postes qu'une version plus récente est publiée.
'''
''' La banque déploie par programme d'installation, poste par poste. Rien n'avertissait un agent
''' qu'il travaillait sur une version dépassée : une correction livrée un lundi pouvait rester
''' ignorée d'un poste pendant des semaines, et deux agents produire des pièces différentes à
''' partir des mêmes rapports.
'''
''' Le dispositif tient en un fichier « version.txt » posé sur le partage, à côté de
''' « connexion.config ». L'informatique l'écrit au moment où elle publie ; les postes le lisent
''' au démarrage et se comparent à lui. Ils ne téléchargent rien et n'installent rien :
''' l'informatique garde la main sur le moment de la mise à jour, ce qui compte pendant la
''' marche en parallèle, où deux versions différentes fausseraient la comparaison.
'''
''' Rien de tout cela n'est bloquant. Un partage injoignable, un fichier absent ou mal écrit
''' laissent l'application démarrer sans un mot : on ne refuse pas la compense du matin parce
''' qu'un fichier d'annonce manque.
''' </summary>
Public NotInheritable Class MiseAJourWU

    Private Sub New()
    End Sub

    ''' <summary>Nom du fichier d'annonce, cherché dans le dossier du fichier partagé.</summary>
    Public Const FICHIER As String = "version.txt"

    Private Const CLE_VERSION As String = "VERSION"
    Private Const CLE_SETUP As String = "SETUP"
    Private Const CLE_NOTE As String = "NOTE"

    ''' <summary>Clé du fichier local retenant la dernière version déjà annoncée à cet utilisateur.</summary>
    Private Const CLE_DEJA_SIGNALEE As String = "VERSION_SIGNALEE"

#Region "Ce que le partage annonce"

    ''' <summary>Version publiée sur le partage, et ce qu'il faut en faire.</summary>
    Public NotInheritable Class VersionPubliee

        ''' <summary>Version annoncée, telle qu'elle est écrite sur le partage.</summary>
        Public Property Numero As String = String.Empty

        ''' <summary>Chemin du programme d'installation à lancer. Peut être vide.</summary>
        Public Property CheminSetup As String = String.Empty

        ''' <summary>Ce que la version apporte, en une ligne. Peut être vide.</summary>
        Public Property Note As String = String.Empty

        ''' <summary>Vrai si elle est réellement postérieure à celle qui tourne.</summary>
        Public Property EstPlusRecente As Boolean = False
    End Class

    ''' <summary>
    ''' Lit l'annonce du partage. Retourne Nothing si aucune mise à jour n'est à signaler —
    ''' partage absent, fichier absent, version illisible, ou version déjà installée.
    ''' </summary>
    Public Shared Function Verifier() As VersionPubliee

        Dim dossier As String = ConfigurationWU.DossierDuPartage
        If dossier.Length = 0 Then Return Nothing

        Dim chemin As String
        Try
            chemin = Path.Combine(dossier, FICHIER)
        Catch
            Return Nothing
        End Try

        Dim valeurs As Dictionary(Of String, String) = LireLeFichier(chemin)
        If valeurs.Count = 0 Then Return Nothing

        Dim annonce As String = Lire(valeurs, CLE_VERSION)
        If annonce.Length = 0 Then Return Nothing

        ' Deux numéros de version se comparent par leurs nombres, jamais par leur texte :
        ' « 1.10.0.0 » est postérieure à « 1.9.0.0 », alors qu'elle la précède alphabétiquement.
        Dim publiee As Version = Nothing
        If Not Version.TryParse(annonce, publiee) Then Return Nothing

        Dim courante As Version = VersionCourante()
        If courante Is Nothing Then Return Nothing

        If publiee <= courante Then Return Nothing

        Return New VersionPubliee() With {
            .Numero = annonce,
            .CheminSetup = Lire(valeurs, CLE_SETUP),
            .Note = Lire(valeurs, CLE_NOTE),
            .EstPlusRecente = True
        }
    End Function

    ''' <summary>Version de l'application qui tourne, ou Nothing si elle est illisible.</summary>
    Public Shared Function VersionCourante() As Version

        Try
            Return Assembly.GetExecutingAssembly().GetName().Version
        Catch
            Return Nothing
        End Try
    End Function

#End Region

#Region "Ne le dire qu'une fois par version"

    ''' <summary>
    ''' Vrai si cette version a déjà fait l'objet d'un message sur ce poste.
    '''
    ''' Répéter l'annonce chaque matin la ferait fermer sans être lue, et la prochaine — celle
    ''' qui compte vraiment — le serait aussi.
    ''' </summary>
    Public Shared Function DejaSignalee(numero As String) As Boolean

        Return String.Equals(ConfigurationWU.LireValeurLocale(CLE_DEJA_SIGNALEE),
                             If(numero, String.Empty).Trim(),
                             StringComparison.OrdinalIgnoreCase)
    End Function

    ''' <summary>Retient qu'on a signalé cette version, pour ne pas y revenir demain.</summary>
    Public Shared Sub MarquerSignalee(numero As String)
        ConfigurationWU.EcrireValeurLocale(CLE_DEJA_SIGNALEE, If(numero, String.Empty).Trim())
    End Sub

#End Region

#Region "Lecture du fichier d'annonce"

    ''' <summary>
    ''' Lit le fichier CLE=VALEUR de l'annonce. Retourne un dictionnaire vide en cas d'échec,
    ''' quel qu'il soit : ce fichier est un confort, jamais une condition de démarrage.
    ''' </summary>
    Private Shared Function LireLeFichier(chemin As String) As Dictionary(Of String, String)

        Dim valeurs As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

        Try
            If Not File.Exists(chemin) Then Return valeurs

            For Each ligne As String In File.ReadAllLines(chemin, Encoding.UTF8)

                Dim nette As String = ligne.Trim()
                If nette.Length = 0 OrElse nette.StartsWith("#", StringComparison.Ordinal) Then Continue For

                Dim separateur As Integer = nette.IndexOf("="c)
                If separateur <= 0 Then Continue For

                Dim cle As String = nette.Substring(0, separateur).Trim()
                If cle.Length > 0 Then valeurs(cle) = nette.Substring(separateur + 1).Trim()
            Next

        Catch
            ' Partage injoignable, droit refusé, fichier verrouillé : rien à signaler.
        End Try

        Return valeurs
    End Function

    Private Shared Function Lire(valeurs As Dictionary(Of String, String), cle As String) As String

        Dim valeur As String = Nothing
        If Not valeurs.TryGetValue(cle, valeur) Then Return String.Empty

        Return If(valeur, String.Empty).Trim()
    End Function

#End Region

End Class
