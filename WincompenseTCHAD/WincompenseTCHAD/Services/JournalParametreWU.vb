Option Strict On
Option Explicit On

Imports System.Data
Imports System.Data.SqlClient

''' <summary>
''' LE JOURNAL DU PARAMÉTRAGE (table T_JournalParametreWU) : qui a changé quoi, quand, depuis
''' quel poste, et ce qu'il y avait avant.
'''
''' POURQUOI UN JOURNAL, ALORS QUE T_ParametreWU PORTE DÉJÀ ModifiePar ET DateModification
'''
''' Ces deux colonnes ne gardent que le DERNIER changement : le précédent est écrasé par le
''' suivant. Pour une case à cocher, c'est assez. Pour le libellé qui part sur toutes les
''' écritures du grand livre de la banque, non : le jour où un comptable demande pourquoi les
''' pièces de septembre ne portent pas la même phrase que celles d'octobre, il faut pouvoir
''' répondre autre chose que « quelqu'un l'a changé, un jour ».
'''
''' LE JOURNAL EST ÉCRIT DANS LA MÊME TRANSACTION QUE LE CHANGEMENT
'''
''' C'est tout l'objet de <see cref="INSERTION"/> : la ligne de journal n'est pas posée par un
''' second aller-retour, elle est ajoutée à la commande qui écrit le paramètre, sous la même
''' SqlTransaction. Deux écritures séparées produiraient tôt ou tard un changement sans trace
''' — ou une trace sans changement, ce qui est pire : elle accuserait quelqu'un d'une
''' modification qui n'a pas eu lieu.
'''
''' ON N'Y REVIENT JAMAIS
'''
''' Aucun droit d'UPDATE ni de DELETE n'est accordé sur cette table, à personne — voir le
''' script 23. Un journal qu'on peut réécrire n'est pas un journal.
'''
''' SON ABSENCE N'EMPÊCHE RIEN
'''
''' Si le script 23 n'a pas été exécuté, <see cref="Disponible"/> rend Faux, la ligne de
''' journal n'est pas demandée, et le paramètre s'enregistre comme avant. L'écran le dit, au
''' lieu de refuser le changement : c'est la règle de tout le paramétrage de ce projet.
''' </summary>
Public NotInheritable Class JournalParametreWU

    Private Sub New()
    End Sub

    Private Const TABLE As String = "T_JournalParametreWU"

    ''' <summary>Code d'erreur SQL Server signalant une table absente (« Invalid object name »).</summary>
    Private Const ERREUR_TABLE_ABSENTE As Integer = 208

    Public Const MESSAGE_TABLE_ABSENTE As String =
        "La table T_JournalParametreWU n'existe pas encore dans la base." & vbCrLf & vbCrLf &
        "Exécutez le script Scripts\23_NarrativeParNature.sql : il la crée." & vbCrLf &
        "Tant qu'elle est absente, le paramétrage fonctionne — mais ses modifications ne " &
        "laissent pas de trace."

#Region "L'insertion, à coudre dans la transaction de l'appelant"

    ''' <summary>
    ''' L'instruction qui pose une ligne de journal, à CONCATÉNER à la commande qui écrit le
    ''' paramètre pour qu'elles partagent sa transaction.
    '''
    ''' Les paramètres sont préfixés « j_ » afin de ne jamais entrer en collision avec ceux de
    ''' l'instruction à laquelle elle est cousue : @cle et @valeur y existent déjà, et une
    ''' collision silencieuse ferait journaliser autre chose que ce qui a été écrit.
    ''' </summary>
    Public Const INSERTION As String =
        "INSERT INTO " & TABLE &
        " (Cle, AncienneValeur, NouvelleValeur, ModifiePar, Poste, DateModification) " &
        "VALUES (@j_cle, @j_ancienne, @j_nouvelle, @j_auteur, @j_poste, GETDATE());"

    ''' <summary>
    ''' Pose les paramètres de <see cref="INSERTION"/> sur la commande de l'appelant.
    '''
    ''' L'ANCIENNE VALEUR PEUT ÊTRE VIDE, et ce n'est pas la même chose qu'un blanc : une clé
    ''' créée pour la première fois n'avait pas de valeur avant. Elle est donc écrite NULL,
    ''' pour que le journal distingue « il n'y avait rien » de « il y avait une chaîne vide ».
    ''' </summary>
    Public Shared Sub AjouterLesParametres(commande As SqlCommand, cle As String,
                                           ancienne As String, nouvelle As String)

        commande.Parameters.Add("@j_cle", SqlDbType.NVarChar, 60).Value = If(cle, String.Empty).Trim()

        Dim avant As String = If(ancienne, String.Empty)
        If avant.Length = 0 Then
            commande.Parameters.Add("@j_ancienne", SqlDbType.NVarChar, 255).Value = DBNull.Value
        Else
            commande.Parameters.Add("@j_ancienne", SqlDbType.NVarChar, 255).Value = avant
        End If

        commande.Parameters.Add("@j_nouvelle", SqlDbType.NVarChar, 255).Value = If(nouvelle, String.Empty)
        commande.Parameters.Add("@j_auteur", SqlDbType.NVarChar, 50).Value = SessionWU.Auteur
        commande.Parameters.Add("@j_poste", SqlDbType.NVarChar, 100).Value = SessionWU.Poste
    End Sub

#End Region

#Region "Disponibilité"

    ''' <summary>
    ''' Vrai si la table existe. Nothing tant que la question n'a pas été posée : la
    ''' distinction compte, Faux signifierait « vérifié, elle n'est pas là ».
    ''' </summary>
    Private Shared _disponible As Boolean?

    ''' <summary>Fait reposer la question. À appeler quand la connexion change de base.</summary>
    Public Shared Sub Oublier()
        _disponible = Nothing
    End Sub

    ''' <summary>
    ''' La table existe-t-elle ? La réponse est mise en cache le temps de la session : une
    ''' table ne se crée pas entre deux enregistrements, et la question serait posée à chaque
    ''' écriture de paramètre.
    '''
    ''' UN ÉCHEC DE LECTURE N'EST PAS MIS EN CACHE. Base injoignable au démarrage, puis
    ''' joignable ensuite : une réponse « non » retenue pour la session priverait de journal
    ''' tous les changements du reste de la journée.
    ''' </summary>
    Public Shared Function Disponible() As Boolean

        If _disponible.HasValue Then Return _disponible.Value

        Const lecture As String =
            "SELECT COUNT(*) FROM sys.tables WHERE name = @nom AND schema_id = SCHEMA_ID('dbo')"

        Try
            Using connexion As SqlConnection = WURepository.CreerConnexion()
                connexion.Open()

                Using commande As New SqlCommand(lecture, connexion)
                    commande.Parameters.Add("@nom", SqlDbType.NVarChar, 128).Value = TABLE

                    Dim nombre As Object = commande.ExecuteScalar()
                    _disponible = nombre IsNot Nothing AndAlso nombre IsNot DBNull.Value AndAlso
                                  Convert.ToInt32(nombre, Globalization.CultureInfo.InvariantCulture) > 0
                End Using
            End Using

        Catch ex As SqlException
            Return False

        Catch ex As InvalidOperationException
            Return False
        End Try

        Return _disponible.Value
    End Function

#End Region

#Region "Relecture"

    ''' <summary>Une modification de paramètre, telle que le journal l'a gardée.</summary>
    Public NotInheritable Class Modification

        Public Property Cle As String = String.Empty
        Public Property AncienneValeur As String = String.Empty
        Public Property NouvelleValeur As String = String.Empty
        Public Property ModifiePar As String = String.Empty
        Public Property Poste As String = String.Empty
        Public Property DateModification As Date

        ''' <summary>
        ''' Ce que la clé désigne, en clair : l'intitulé de la nature pour un libellé par
        ''' nature, la clé brute pour une option. C'est cette colonne que la grille montre, et
        ''' non le code : « TVA collectée » se lit, « NATURE_TVA » se déchiffre.
        ''' </summary>
        Public ReadOnly Property Intitule As String
            Get
                Return NaturesMouvementWU.IntituleDeCle(Cle)
            End Get
        End Property

        ''' <summary>
        ''' L'ancienne valeur telle qu'on l'affiche. Une clé créée pour la première fois n'en
        ''' avait pas : le dire vaut mieux que montrer une case vide, qu'on lirait comme un
        ''' libellé effacé.
        ''' </summary>
        Public ReadOnly Property Avant As String
            Get
                If String.IsNullOrEmpty(AncienneValeur) Then Return "(aucune valeur)"
                Return AncienneValeur
            End Get
        End Property
    End Class

    ''' <summary>
    ''' Les modifications, de la plus récente à la plus ancienne.
    '''
    ''' BORNÉ, parce qu'un journal se consulte et ne se télécharge pas : au bout de deux ans,
    ''' une grille qui charge tout met l'écran plusieurs secondes à s'ouvrir pour montrer
    ''' trente lignes qu'on ne fera jamais défiler.
    ''' </summary>
    Public Shared Function Lister(limite As Integer, ByRef messageErreur As String) As List(Of Modification)

        messageErreur = String.Empty

        Dim lignes As New List(Of Modification)
        Dim nombre As Integer = If(limite > 0, limite, 200)

        Const lecture As String =
            "SELECT TOP (@limite) Cle, AncienneValeur, NouvelleValeur, ModifiePar, Poste, DateModification " &
            "FROM   " & TABLE & " " &
            "ORDER BY DateModification DESC, Id DESC"

        Try
            Using connexion As SqlConnection = WURepository.CreerConnexion()
                connexion.Open()

                Using commande As New SqlCommand(lecture, connexion)
                    commande.Parameters.Add("@limite", SqlDbType.Int).Value = nombre

                    Using lecteur As SqlDataReader = commande.ExecuteReader()
                        While lecteur.Read()
                            lignes.Add(New Modification() With {
                                .Cle = LireChaine(lecteur, 0),
                                .AncienneValeur = LireChaine(lecteur, 1),
                                .NouvelleValeur = LireChaine(lecteur, 2),
                                .ModifiePar = LireChaine(lecteur, 3),
                                .Poste = LireChaine(lecteur, 4),
                                .DateModification = If(lecteur.IsDBNull(5), Date.MinValue, lecteur.GetDateTime(5))
                            })
                        End While
                    End Using
                End Using
            End Using

        Catch ex As SqlException
            messageErreur = If(ex.Number = ERREUR_TABLE_ABSENTE,
                               MESSAGE_TABLE_ABSENTE,
                               $"Lecture du journal du paramétrage impossible : {ex.Message}")

        Catch ex As InvalidOperationException
            messageErreur = $"Connexion SQL Server indisponible : {ex.Message}"
        End Try

        Return lignes
    End Function

    Private Shared Function LireChaine(lecteur As SqlDataReader, position As Integer) As String

        If lecteur.IsDBNull(position) Then Return String.Empty
        Return Convert.ToString(lecteur.GetValue(position)).Trim()
    End Function

#End Region

End Class
