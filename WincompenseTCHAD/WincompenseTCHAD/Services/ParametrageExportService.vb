Option Strict On
Option Explicit On

Imports System.Data.SqlClient

''' <summary>
''' Sortie du paramétrage dans un fichier de secours.
'''
''' CE QUE CE FICHIER EST, ET CE QU'IL N'EST PAS
'''
''' Il n'est PAS une sauvegarde de la base. La sauvegarde d'une base SQL Server se fait sur le
''' serveur, par l'équipe qui le tient : `BACKUP DATABASE` demande un droit de niveau serveur
''' que l'application n'a pas et ne doit pas avoir, écrit sur le disque du serveur et non sur
''' celui du poste, et une sauvegarde qui ne se déclenche que si quelqu'un ouvre l'application
''' n'est pas une sauvegarde.
'''
''' Il est un FILET : de quoi remonter une installation neuve sans resaisir à la main des
''' centaines de sous-agents. C'est peu, et c'est précisément ce qu'une application de poste
''' peut promettre sans mentir.
''' </summary>
Public NotInheritable Class ParametrageExportService

    Private Sub New()
    End Sub

    ''' <summary>Nombre d'étapes annoncées à la barre de progression.</summary>
    Public Const ETAPES As Integer = 8

#Region "Lecture de la base"

    ''' <summary>
    ''' Rassemble le paramétrage courant.
    '''
    ''' Une table illisible n'interrompt pas la collecte : mieux vaut un fichier de secours
    ''' incomplet, et qui le dit, que pas de fichier du tout. Les motifs sont rendus dans
    ''' messageErreur, que l'écran affiche.
    ''' </summary>
    Public Shared Function Construire(progression As ProgressionWU,
                                      ByRef messageErreur As String) As LotParametrageWU

        messageErreur = String.Empty

        Dim lot As New LotParametrageWU() With {
            .Version = ParametrageFichierWU.VERSION,
            .DateExport = Date.Now,
            .ExportePar = SessionWU.Auteur
        }

        RenseignerLaProvenance(lot)

        Dim soucis As New List(Of String)()
        Dim motif As String = String.Empty

        Annoncer(progression, "Lecture des comptes comptables")
        lot.Comptes = WURepository.ChargerComptesSysteme(motif)
        If motif.Length > 0 Then soucis.Add("Comptes comptables : " & PremiereLigne(motif))

        Annoncer(progression, "Lecture des groupes statistiques")
        lot.Groupes = PdvRepository.ListerGroupes(String.Empty, motif)
        If motif.Length > 0 Then soucis.Add("Groupes statistiques : " & PremiereLigne(motif))

        Annoncer(progression, "Lecture des sous-agents")
        lot.SousAgents = PdvRepository.ListerSousAgents(String.Empty, motif)
        If motif.Length > 0 Then soucis.Add("Sous-agents : " & PremiereLigne(motif))

        Annoncer(progression, "Lecture des agences propres")
        lot.Agences = PdvRepository.ListerAgences(String.Empty, motif)
        If motif.Length > 0 Then soucis.Add("Agences propres : " & PremiereLigne(motif))

        Annoncer(progression, "Lecture des options de traitement")
        lot.Options = OptionsWU.ListerToutes(motif)
        If motif.Length > 0 Then soucis.Add("Options de traitement : " & PremiereLigne(motif))

        Annoncer(progression, "Lecture des narratives par nature")
        lot.Narratives = NarrativesParNature(motif)
        If motif.Length > 0 Then soucis.Add("Narratives par nature : " & PremiereLigne(motif))

        Annoncer(progression, "Lecture des utilisateurs")
        lot.Utilisateurs = UtilisateurRepository.Lister(motif)
        If motif.Length > 0 Then soucis.Add("Utilisateurs : " & PremiereLigne(motif))

        If soucis.Count > 0 Then
            messageErreur = "Le fichier a été constitué, mais il est INCOMPLET :" &
                            Environment.NewLine & " - " & String.Join(Environment.NewLine & " - ", soucis)
        End If

        Return lot
    End Function

    ''' <summary>
    ''' Les narratives par nature, converties en couples clé / valeur pour le fichier.
    '''
    ''' LA TABLE ABSENTE N'EST PAS UN INCIDENT. Une base qui n'a jamais joué le script des
    ''' narratives par nature n'en a aucune, et c'est un état normal : le fichier de secours
    ''' sort sans cette section, sans se déclarer incomplet pour autant. Signaler un manque
    ''' là où il n'y a rien à manquer apprendrait à l'administrateur à ignorer les
    ''' avertissements.
    ''' </summary>
    Private Shared Function NarrativesParNature(ByRef messageErreur As String) As List(Of OptionParametrageWU)

        messageErreur = String.Empty
        Dim lues As New List(Of OptionParametrageWU)

        If Not NarrativeRepository.Disponible() Then Return lues

        Dim modeles As Dictionary(Of NatureMouvementWU, String) =
            NarrativeRepository.Lire(messageErreur)

        ' L'ordre est celui des natures, et non celui de la base : deux exports de la même
        ' base doivent donner deux fichiers identiques, sinon l'empreinte ne prouve rien.
        For Each nature As NatureMouvementWU In NaturesMouvementWU.Toutes()

            If Not modeles.ContainsKey(nature) Then Continue For

            lues.Add(New OptionParametrageWU(NaturesMouvementWU.Code(nature),
                                             modeles(nature),
                                             NaturesMouvementWU.Intitule(nature)))
        Next

        Return lues
    End Function

    ''' <summary>
    ''' Inscrit d'où vient le lot : base et serveur, lus dans la chaîne de connexion.
    '''
    ''' Sans cela, deux fichiers exportés le même jour depuis deux installations différentes
    ''' seraient impossibles à distinguer — et on chargerait le paramétrage du Tchad dans la
    ''' base d'un autre pays sans s'en apercevoir.
    ''' </summary>
    Private Shared Sub RenseignerLaProvenance(lot As LotParametrageWU)

        Try
            Dim constructeur As New SqlConnectionStringBuilder(WURepository.ObtenirChaineConnexion())
            lot.Base = constructeur.InitialCatalog
            lot.Serveur = constructeur.DataSource

        Catch ex As ArgumentException
            ' Chaîne de connexion inhabituelle : le fichier sort quand même, sans provenance.
            lot.Base = String.Empty
            lot.Serveur = String.Empty
        Catch ex As FormatException
            lot.Base = String.Empty
            lot.Serveur = String.Empty
        End Try
    End Sub

#End Region

#Region "Écriture"

    ''' <summary>
    ''' Constitue le lot et l'écrit. Le message d'avertissement d'une collecte incomplète est
    ''' conservé : le fichier existe, mais celui qui le range doit savoir ce qui lui manque.
    ''' </summary>
    Public Shared Function Exporter(chemin As String, progression As ProgressionWU,
                                    ByRef messageAvertissement As String,
                                    ByRef messageErreur As String) As LotParametrageWU

        messageAvertissement = String.Empty
        messageErreur = String.Empty

        Dim lot As LotParametrageWU = Construire(progression, messageAvertissement)

        If lot.EstVide Then
            messageErreur = "Il n'y a rien à exporter : ni comptes comptables, ni groupes, " &
                            "ni points de vente n'ont pu être lus." & Environment.NewLine &
                            "Vérifiez la connexion à la base avant de recommencer."
            Return Nothing
        End If

        If Not ParametrageFichierWU.Ecrire(lot, chemin, progression, messageErreur) Then Return Nothing

        Return lot
    End Function

#End Region

#Region "Utilitaires"

    Private Shared Sub Annoncer(progression As ProgressionWU, libelle As String)
        If progression Is Nothing Then Return
        progression.Avancer(libelle)
    End Sub

    ''' <summary>La première ligne d'un message, pour une énumération qui doit rester courte.</summary>
    Private Shared Function PremiereLigne(texte As String) As String

        Dim fin As Integer = texte.IndexOfAny(New Char() {ControlChars.Cr, ControlChars.Lf})
        If fin < 0 Then Return texte
        Return texte.Substring(0, fin)
    End Function

#End Region

End Class
