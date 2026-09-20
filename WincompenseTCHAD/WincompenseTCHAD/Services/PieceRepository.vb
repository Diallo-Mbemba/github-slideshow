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
    Private Const TABLE_PIECE_ARCHIVE As String = "T_PieceAnnuleeWU"
    Private Const TABLE_ANNULATION As String = "T_AnnulationWU"

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
        ''' Identifiant de l'annulation qui a retiré cette journée, ou 0 si elle est vivante.
        '''
        ''' C'est lui, et non la date, qui identifie une archive : une même journée peut être
        ''' comptabilisée, annulée, recomptabilisée et annulée de nouveau.
        ''' </summary>
        Public Property IdAnnulation As Long = 0L

        ''' <summary>Motif de l'annulation, tel qu'enregistré (RAPPORT_VIDE, …).</summary>
        Public Property MotifAnnulation As String = String.Empty

        Public Property DateAnnulation As Date?
        Public Property AnnuleePar As String = String.Empty

        ''' <summary>Vrai si cette ligne est une archive et non la journée en vigueur.</summary>
        Public ReadOnly Property Annulee As Boolean
            Get
                Return IdAnnulation > 0L
            End Get
        End Property

        ''' <summary>« Comptabilisée » ou « ANNULÉE », pour la colonne d'état de la liste.</summary>
        Public ReadOnly Property LibelleEtat As String
            Get
                If Not Annulee Then Return "Comptabilisée"
                Return "ANNULÉE — " & AnnulationWU.LibelleDepuisCode(MotifAnnulation)
            End Get
        End Property

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
    ''' Les journées dont une pièce est conservée, la plus récente d'abord — les journées
    ''' en vigueur ET celles qui ont été annulées.
    '''
    ''' POURQUOI LES DEUX DANS UNE SEULE LISTE
    '''
    ''' Une journée annulée ne compte plus dans les rapports, mais elle a existé : elle a
    ''' porté des écritures, peut-être un fichier core banking, et quelqu'un l'a retirée
    ''' pour une raison écrite. La faire disparaître de l'écran reviendrait à annuler
    ''' l'annulation elle-même, qui est précisément ce qu'un auditeur vient chercher.
    '''
    ''' UNE MÊME DATE PEUT APPARAÎTRE PLUSIEURS FOIS
    '''
    ''' Comptabilisée, annulée, recomptabilisée : la journée en vigueur vient en tête, ses
    ''' archives en dessous, dans l'ordre où elles se sont produites.
    '''
    ''' Une table absente n'est pas une erreur de fonctionnement mais une installation
    ''' incomplète : le message le dit, et l'écran affiche une liste vide plutôt que de se
    ''' fermer sur une exception. Si seules les tables d'annulation manquent, la liste des
    ''' journées vivantes s'affiche quand même : l'archive est un complément, pas une
    ''' condition.
    ''' </summary>
    Public Shared Function ListerJournees(ByRef messageErreur As String) As List(Of JourneeConservee)

        messageErreur = String.Empty
        Dim journees As New List(Of JourneeConservee)()

        Try
            Using connexion As SqlConnection = WURepository.CreerConnexion()
                connexion.Open()

                LireLesJourneesVivantes(connexion, journees)

                If ArchivePresente(connexion) Then
                    LireLesJourneesAnnulees(connexion, journees)
                End If
            End Using

        Catch ex As SqlException
            messageErreur = If(ex.Number = ERREUR_TABLE_ABSENTE,
                               MESSAGE_TABLE_ABSENTE,
                               $"Lecture des pièces conservées impossible : {ex.Message}")

        Catch ex As InvalidOperationException
            messageErreur = $"Connexion SQL Server indisponible : {ex.Message}"
        End Try

        ' La plus récente d'abord ; à date égale, la journée en vigueur avant ses archives.
        journees.Sort(Function(gauche, droite)
                          Dim ordre As Integer = droite.DateActivite.CompareTo(gauche.DateActivite)
                          If ordre <> 0 Then Return ordre
                          Return gauche.IdAnnulation.CompareTo(droite.IdAnnulation)
                      End Function)

        Return journees
    End Function

    ''' <summary>
    ''' Vrai si les tables d'annulation sont en place. La question se pose AVANT d'écrire
    ''' une requête qui les nomme : SQL Server refuse une requête dont une table manque,
    ''' même si le reste est valable — la liste entière échouerait pour une installation
    ''' incomplète.
    ''' </summary>
    Private Shared Function ArchivePresente(connexion As SqlConnection) As Boolean

        Const question As String =
            "SELECT CASE WHEN OBJECT_ID(N'dbo.T_AnnulationWU') IS NOT NULL " &
            "        AND OBJECT_ID(N'dbo.T_PieceAnnuleeWU') IS NOT NULL " &
            "       THEN 1 ELSE 0 END"

        Using commande As New SqlCommand(question, connexion)
            Dim valeur As Object = commande.ExecuteScalar()
            If valeur Is Nothing OrElse valeur Is DBNull.Value Then Return False
            Return Convert.ToInt32(valeur, Globalization.CultureInfo.InvariantCulture) = 1
        End Using
    End Function

    Private Shared Sub LireLesJourneesVivantes(connexion As SqlConnection,
                                               journees As List(Of JourneeConservee))

        Const lecture As String =
            "SELECT  DateActivite," &
            "        ecritures   = COUNT(*)," &
            "        totalDebit  = SUM(Debit)," &
            "        totalCredit = SUM(Credit)," &
            "        enregistree = MAX(DateEnregistrement)," &
            "        auteur      = MIN(EnregistrePar) " &
            "FROM    " & TABLE_PIECE & " " &
            "GROUP BY DateActivite"

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
    End Sub

    ''' <summary>
    ''' Les journées annulées. Les totaux viennent de l'en-tête d'annulation, qui les a
    ''' figés au moment du retrait ; la date de conservation et son auteur viennent des
    ''' lignes archivées, parce que ce sont ceux de la comptabilisation d'origine et non
    ''' ceux de l'annulation.
    ''' </summary>
    Private Shared Sub LireLesJourneesAnnulees(connexion As SqlConnection,
                                               journees As List(Of JourneeConservee))

        Const lecture As String =
            "SELECT  a.DateActivite," &
            "        a.NombreEcritures," &
            "        a.TotalDebit," &
            "        a.TotalCredit," &
            "        enregistree = ISNULL(p.enregistree, a.DateAutorisation)," &
            "        auteur      = ISNULL(p.auteur, a.DemandeePar)," &
            "        a.IdAnnulation," &
            "        a.Motif," &
            "        a.DateAutorisation," &
            "        a.AutoriseePar " &
            "FROM    " & TABLE_ANNULATION & " AS a " &
            "LEFT JOIN (SELECT IdAnnulation," &
            "                  enregistree = MAX(DateEnregistrement)," &
            "                  auteur      = MIN(EnregistrePar) " &
            "           FROM   " & TABLE_PIECE_ARCHIVE & " " &
            "           GROUP BY IdAnnulation) AS p ON p.IdAnnulation = a.IdAnnulation"

        Using commande As New SqlCommand(lecture, connexion)
            Using lecteur As SqlDataReader = commande.ExecuteReader()

                While lecteur.Read()
                    journees.Add(New JourneeConservee() With {
                        .DateActivite = lecteur.GetDateTime(0),
                        .NombreEcritures = lecteur.GetInt32(1),
                        .TotalDebit = LireEntier(lecteur, 2),
                        .TotalCredit = LireEntier(lecteur, 3),
                        .DateEnregistrement = lecteur.GetDateTime(4),
                        .EnregistrePar = LireChaine(lecteur, 5),
                        .IdAnnulation = LireEntier(lecteur, 6),
                        .MotifAnnulation = LireChaine(lecteur, 7),
                        .DateAnnulation = lecteur.GetDateTime(8),
                        .AnnuleePar = LireChaine(lecteur, 9)
                    })
                End While
            End Using
        End Using
    End Sub
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

        Dim dt As DataTable = TableVide()

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

    ''' <summary>
    ''' La pièce d'une journée ANNULÉE, dans la même forme que celle d'une journée vivante.
    '''
    ''' Elle se charge par identifiant d'annulation et non par date : une même journée peut
    ''' avoir été annulée plusieurs fois, et la date ne dirait pas laquelle est demandée.
    '''
    ''' Les colonnes sont identiques à celles de Charger : l'écran de consultation, l'export
    ''' Excel et le fichier core banking n'ont pas à savoir d'où vient la pièce qu'on leur
    ''' donne. Une pièce archivée reste une pièce.
    ''' </summary>
    Public Shared Function ChargerAnnulee(idAnnulation As Long, ByRef messageErreur As String) As DataTable

        messageErreur = String.Empty

        Dim dt As DataTable = TableVide()

        Const lecture As String =
            "SELECT Compte, Libelle, Debit, Credit, CodeAgence " &
            "FROM   " & TABLE_PIECE_ARCHIVE & " " &
            "WHERE  IdAnnulation = @id " &
            "ORDER BY Ligne"

        Try
            Using connexion As SqlConnection = WURepository.CreerConnexion()
                connexion.Open()

                Using commande As New SqlCommand(lecture, connexion)
                    commande.Parameters.Add("@id", SqlDbType.BigInt).Value = idAnnulation

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
                               AnnulationRepository.MESSAGE_TABLES_ABSENTES,
                               $"Lecture de la pièce archivée impossible : {ex.Message}")

        Catch ex As InvalidOperationException
            messageErreur = $"Connexion SQL Server indisponible : {ex.Message}"
        End Try

        Return dt
    End Function

    ''' <summary>
    ''' Une pièce vide, aux colonnes exactes qu'attend le reste de l'application. Définie
    ''' une fois : deux listes de colonnes qui doivent rester identiques ne le restent pas.
    ''' </summary>
    Private Shared Function TableVide() As DataTable

        Dim dt As New DataTable("dtPiece")
        dt.Columns.Add("Compte", GetType(String))
        dt.Columns.Add("Libelle", GetType(String))
        dt.Columns.Add("Debit", GetType(Long))
        dt.Columns.Add("Credit", GetType(Long))
        dt.Columns.Add("CodeAgence", GetType(String))
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
