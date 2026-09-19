Option Strict On
Option Explicit On

Imports System.Data
Imports System.Data.SqlClient

''' <summary>
''' Pièces comptables conservées (table T_PieceWU) : écriture au moment de l'historisation,
''' relecture pour consultation.
'''
''' POURQUOI CONSERVER LA PIÈCE PLUTÔT QUE DE LA RECALCULER
'''
''' L'historique garde les volumes, les montants et les totaux de commissions. Il ne garde pas
''' le taux du sous-agent ce jour-là, ni ses comptes de compensation et de commission, ni la
''' ligne d'écart posée sur le compte inter bancaire. Reconstituer une pièce ancienne avec le
''' paramétrage d'aujourd'hui réécrirait donc le passé : un sous-agent passé de 70 % à 60 %
''' ferait apparaître une pièce qui n'a jamais été visée ni signée.
'''
''' Une pièce comptable est un justificatif. « À peu près la même » n'a pas de sens devant un
''' inspecteur. Cette classe conserve donc la pièce TELLE QU'ELLE A ÉTÉ PRODUITE, ligne à
''' ligne : la consulter, c'est relire ce qui a été écrit, et rien n'est recalculé.
'''
''' Comme les autres dépôts, aucune exception SQL ne remonte à l'interface : chaque fonction
''' retourne un booléen de réussite et pose un message explicite, en français.
''' </summary>
Public NotInheritable Class PieceRepository

    Private Sub New()
    End Sub

    Private Const TABLE_PIECE As String = "T_PieceWU"

    ''' <summary>Code d'erreur SQL Server signalant une table absente (« Invalid object name »).</summary>
    Private Const ERREUR_TABLE_ABSENTE As Integer = 208

    Public Const MESSAGE_TABLE_ABSENTE As String =
        "La table T_PieceWU n'existe pas encore dans la base." & vbCrLf & vbCrLf &
        "Exécutez le script Scripts\13_PiecesComptables.sql : il la crée." & vbCrLf &
        "Tant qu'elle est absente, les pièces ne sont pas conservées et l'écran de consultation " &
        "reste vide. Les journées déjà comptabilisées, elles, ne pourront jamais être retrouvées " &
        "telles qu'elles ont été produites."

#Region "Description d'une journée conservée"

    ''' <summary>
    ''' Ce que la liste des journées affiche : de quoi choisir sans ouvrir la pièce.
    ''' </summary>
    Public NotInheritable Class JourneeConservee

        Public Property DateActivite As Date
        Public Property NombreEcritures As Integer
        Public Property TotalDebit As Long
        Public Property TotalCredit As Long
        Public Property DateEnregistrement As Date
        Public Property EnregistrePar As String = String.Empty

        ''' <summary>
        ''' Vrai si la pièce s'équilibre. Elle le fait toujours — le contrôle d'équilibre
        ''' précède l'historisation — mais l'afficher permet de le CONSTATER plutôt que de le
        ''' croire, ce qui est tout l'objet d'une consultation d'archive.
        ''' </summary>
        Public ReadOnly Property Equilibree As Boolean
            Get
                Return TotalDebit = TotalCredit
            End Get
        End Property
    End Class

#End Region

#Region "Écriture"

    ''' <summary>
    ''' Écrit la pièce d'une journée, DANS LA TRANSACTION DE L'APPELANT.
    '''
    ''' La connexion et la transaction sont reçues et non créées : la pièce et l'historique
    ''' décrivent la même journée, et deux documents d'une même journée ne doivent jamais
    ''' diverger. Si l'un échoue, aucun des deux n'est écrit.
    '''
    ''' L'écriture est RÉPÉTABLE : les lignes existantes de la journée sont d'abord supprimées.
    ''' Une journée regénérée remplace donc intégralement la précédente.
    ''' </summary>
    ''' <param name="jour">Journée comptabilisée.</param>
    ''' <param name="dtPiece">Pièce produite : colonnes Compte, Libelle, Debit, Credit, CodeAgence.</param>
    ''' <param name="connexion">Connexion ouverte de l'appelant.</param>
    ''' <param name="transaction">Transaction en cours de l'appelant.</param>
    ''' <returns>Nombre d'écritures conservées.</returns>
    Public Shared Function Enregistrer(jour As Date, dtPiece As DataTable,
                                       connexion As SqlConnection, transaction As SqlTransaction) As Integer

        If dtPiece Is Nothing OrElse dtPiece.Rows.Count = 0 Then Return 0

        Const suppression As String = "DELETE FROM " & TABLE_PIECE & " WHERE DateActivite = @jour"

        Const insertion As String =
            "INSERT INTO " & TABLE_PIECE & " (DateActivite, Ligne, Compte, Libelle, Debit, Credit, " &
            "CodeAgence, DateEnregistrement, EnregistrePar) " &
            "VALUES (@jour, @ligne, @compte, @libelle, @debit, @credit, @codeAgence, GETDATE(), @auteur)"

        Using commande As New SqlCommand(suppression, connexion, transaction)
            commande.Parameters.Add("@jour", SqlDbType.Date).Value = jour.Date
            commande.ExecuteNonQuery()
        End Using

        Dim rang As Integer = 0
        Dim porteLeCodeAgence As Boolean = dtPiece.Columns.Contains("CodeAgence")

        For Each ligne As DataRow In dtPiece.Rows

            Dim debit As Long = Convert.ToInt64(ligne("Debit"))
            Dim credit As Long = Convert.ToInt64(ligne("Credit"))

            ' Une écriture sans montant n'a rien à faire dans une pièce. Les auxiliaires de
            ' PieceComptableService l'écartent déjà ; le contrôle est ici pour que la table ne
            ' dépende pas de cette politesse — sa contrainte CHECK refuserait la ligne, et
            ' c'est toute la journée qui ne serait pas historisée.
            If debit = 0L AndAlso credit = 0L Then Continue For

            rang += 1

            Using commande As New SqlCommand(insertion, connexion, transaction)

                commande.Parameters.Add("@jour", SqlDbType.Date).Value = jour.Date
                commande.Parameters.Add("@ligne", SqlDbType.Int).Value = rang
                commande.Parameters.Add("@compte", SqlDbType.NVarChar, 50).Value =
                    Convert.ToString(ligne("Compte"))
                commande.Parameters.Add("@libelle", SqlDbType.NVarChar, 255).Value =
                    Convert.ToString(ligne("Libelle"))
                commande.Parameters.Add("@debit", SqlDbType.BigInt).Value = debit
                commande.Parameters.Add("@credit", SqlDbType.BigInt).Value = credit
                commande.Parameters.Add("@codeAgence", SqlDbType.NVarChar, 50).Value =
                    If(porteLeCodeAgence, Convert.ToString(ligne("CodeAgence")), String.Empty)
                commande.Parameters.Add("@auteur", SqlDbType.NVarChar, 100).Value = SessionWU.Auteur

                commande.ExecuteNonQuery()
            End Using
        Next

        Return rang
    End Function

#End Region

#Region "Lecture"

    ''' <summary>
    ''' Les journées dont une pièce est conservée, la plus récente d'abord.
    '''
    ''' Une table absente n'est pas une erreur de fonctionnement mais une installation
    ''' incomplète : le message le dit, et l'écran affiche une liste vide plutôt que de se
    ''' fermer sur une exception.
    ''' </summary>
    Public Shared Function ListerJournees(ByRef messageErreur As String) As List(Of JourneeConservee)

        messageErreur = String.Empty
        Dim journees As New List(Of JourneeConservee)()

        Const lecture As String =
            "SELECT  DateActivite," &
            "        ecritures   = COUNT(*)," &
            "        totalDebit  = SUM(Debit)," &
            "        totalCredit = SUM(Credit)," &
            "        enregistree = MAX(DateEnregistrement)," &
            "        auteur      = MIN(EnregistrePar) " &
            "FROM    " & TABLE_PIECE & " " &
            "GROUP BY DateActivite " &
            "ORDER BY DateActivite DESC"

        Try
            Using connexion As SqlConnection = WURepository.CreerConnexion()
                connexion.Open()

                Using commande As New SqlCommand(lecture, connexion)
                    Using lecteur As SqlDataReader = commande.ExecuteReader()

                        While lecteur.Read()
                            journees.Add(New JourneeConservee() With {
                                .DateActivite = lecteur.GetDateTime(0),
                                .NombreEcritures = lecteur.GetInt32(1),
                                .TotalDebit = LireEntier(lecteur, 2),
                                .TotalCredit = LireEntier(lecteur, 3),
                                .DateEnregistrement = lecteur.GetDateTime(4),
                                .EnregistrePar = LireChaine(lecteur, 5)
                            })
                        End While
                    End Using
                End Using
            End Using

        Catch ex As SqlException
            messageErreur = If(ex.Number = ERREUR_TABLE_ABSENTE,
                               MESSAGE_TABLE_ABSENTE,
                               $"Lecture des pièces conservées impossible : {ex.Message}")

        Catch ex As InvalidOperationException
            messageErreur = $"Connexion SQL Server indisponible : {ex.Message}"
        End Try

        Return journees
    End Function

    ''' <summary>
    ''' La pièce d'une journée, dans la forme EXACTE que le reste de l'application attend d'une
    ''' pièce : mêmes colonnes, même ordre, même nom de table.
    '''
    ''' C'est ce qui permet de la présenter avec FrmPieceComptable, de l'exporter avec
    ''' PieceExcelWU et d'en reconstruire le fichier core banking avec CoreBankingService, sans
    ''' qu'aucun d'eux n'ait à savoir qu'elle vient d'une archive.
    ''' </summary>
    Public Shared Function Charger(jour As Date, ByRef messageErreur As String) As DataTable

        messageErreur = String.Empty

        Dim dt As New DataTable("dtPiece")
        dt.Columns.Add("Compte", GetType(String))
        dt.Columns.Add("Libelle", GetType(String))
        dt.Columns.Add("Debit", GetType(Long))
        dt.Columns.Add("Credit", GetType(Long))
        dt.Columns.Add("CodeAgence", GetType(String))

        Const lecture As String =
            "SELECT Compte, Libelle, Debit, Credit, CodeAgence " &
            "FROM   " & TABLE_PIECE & " " &
            "WHERE  DateActivite = @jour " &
            "ORDER BY Ligne"

        Try
            Using connexion As SqlConnection = WURepository.CreerConnexion()
                connexion.Open()

                Using commande As New SqlCommand(lecture, connexion)
                    commande.Parameters.Add("@jour", SqlDbType.Date).Value = jour.Date

                    Using lecteur As SqlDataReader = commande.ExecuteReader()
                        While lecteur.Read()
                            dt.Rows.Add(LireChaine(lecteur, 0), LireChaine(lecteur, 1),
                                        LireEntier(lecteur, 2), LireEntier(lecteur, 3),
                                        LireChaine(lecteur, 4))
                        End While
                    End Using
                End Using
            End Using

        Catch ex As SqlException
            messageErreur = If(ex.Number = ERREUR_TABLE_ABSENTE,
                               MESSAGE_TABLE_ABSENTE,
                               $"Lecture de la pièce du {jour:dd/MM/yyyy} impossible : {ex.Message}")

        Catch ex As InvalidOperationException
            messageErreur = $"Connexion SQL Server indisponible : {ex.Message}"
        End Try

        Return dt
    End Function

#End Region

#Region "Utilitaires de lecture"

    Private Shared Function LireChaine(lecteur As SqlDataReader, position As Integer) As String
        If lecteur.IsDBNull(position) Then Return String.Empty
        Return lecteur.GetString(position)
    End Function

    Private Shared Function LireEntier(lecteur As SqlDataReader, position As Integer) As Long
        If lecteur.IsDBNull(position) Then Return 0L
        Return Convert.ToInt64(lecteur.GetValue(position))
    End Function

#End Region

End Class
