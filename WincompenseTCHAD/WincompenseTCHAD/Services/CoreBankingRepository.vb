Option Strict On
Option Explicit On

Imports System.Data
Imports System.Data.SqlClient

''' <summary>
''' Trace des fichiers produits pour le core banking (table T_FichierCoreBankingWU).
'''
''' POURQUOI TRACER UN FICHIER QU'ON NE STOCKE PAS
'''
''' Le fichier n'est pas conservé, et n'a pas à l'être : il dérive entièrement de la pièce,
''' par les mêmes règles. Mais le FAIT de l'avoir produit ne se déduit de rien, et c'est le
''' seul moment où la journée quitte Wincompense pour entrer dans les livres de la banque.
'''
''' Sans cette trace, l'application ne peut ni avertir avant de recomptabiliser une journée
''' déjà partie, ni prévenir avant de l'annuler. Avec elle, les deux avertissements se
''' posent d'eux-mêmes, au lieu de reposer sur la mémoire de l'agent.
'''
''' CE QUE LA TRACE NE DIT PAS : que le fichier a été INJECTÉ. L'application écrit un
''' classeur ; elle ne voit pas ce que le core banking en fait. Elle dit « ce fichier est
''' sorti, tel jour, par telle personne » — la suite se demande à l'agent.
'''
''' Comme les autres dépôts, aucune exception SQL ne remonte à l'interface.
''' </summary>
Public NotInheritable Class CoreBankingRepository

    Private Sub New()
    End Sub

    Private Const TABLE As String = "T_FichierCoreBankingWU"

    ''' <summary>Code d'erreur SQL Server signalant une table absente (« Invalid object name »).</summary>
    Private Const ERREUR_TABLE_ABSENTE As Integer = 208

    Public Const MESSAGE_TABLE_ABSENTE As String =
        "La table T_FichierCoreBankingWU n'existe pas encore dans la base." & vbCrLf & vbCrLf &
        "Exécutez le script Scripts\15_FichierCoreBanking.sql : il la crée." & vbCrLf &
        "Tant qu'elle est absente, l'application ne sait pas quelles journées sont déjà " &
        "parties vers le core banking, et ne peut donc pas en avertir."

#Region "Description d'une production"

    ''' <summary>
    ''' Un fichier produit, tel que les avertissements l'annoncent.
    ''' </summary>
    Public NotInheritable Class Production

        Public Property DateActivite As Date
        Public Property DateValeur As Date
        Public Property NumeroLot As String = String.Empty
        Public Property NomFichier As String = String.Empty
        Public Property NombreLignes As Integer = 0
        Public Property TotalDebit As Long = 0L
        Public Property DateProduction As Date
        Public Property ProduitPar As String = String.Empty

        ''' <summary>
        ''' La phrase d'avertissement, telle qu'elle s'affiche avant de recomptabiliser ou
        ''' d'annuler la journée. Elle nomme la personne : c'est elle qu'il faut appeler
        ''' pour savoir ce que le fichier est devenu.
        ''' </summary>
        Public ReadOnly Property Avertissement As String
            Get
                Dim phrase As String =
                    $"Le fichier destiné au core banking de cette journée a été produit le " &
                    $"{DateProduction:dd/MM/yyyy à HH:mm}"

                If ProduitPar.Length > 0 Then phrase &= $" par {ProduitPar}"
                If NumeroLot.Length > 0 Then phrase &= $" (lot {NumeroLot})"

                Return phrase & "."
            End Get
        End Property
    End Class

#End Region

#Region "Écriture"

    ''' <summary>
    ''' Enregistre qu'un fichier vient d'être produit.
    '''
    ''' L'échec n'est JAMAIS bloquant : le fichier existe, il est ouvert dans Excel, et
    ''' refuser la production parce que la trace n'a pas pu s'écrire ferait perdre le
    ''' travail pour sauver le carnet. L'appelant reçoit le message et décide de l'afficher
    ''' ou de se taire.
    ''' </summary>
    Public Shared Function Enregistrer(dateActivite As Date, dateValeur As Date, numeroLot As String,
                                       cheminFichier As String, nombreLignes As Integer,
                                       totalDebit As Long, totalCredit As Long,
                                       ByRef messageErreur As String) As Boolean

        messageErreur = String.Empty

        Const insertion As String =
            "INSERT INTO " & TABLE & " (DateActivite, DateValeur, NumeroLot, NomFichier, " &
            "CheminFichier, NombreLignes, TotalDebit, TotalCredit, DateProduction, ProduitPar) " &
            "VALUES (@jour, @valeur, @lot, @nom, @chemin, @lignes, @debit, @credit, " &
            "GETDATE(), @auteur)"

        Dim chemin As String = If(cheminFichier, String.Empty)
        Dim nom As String = String.Empty

        Try
            If chemin.Length > 0 Then nom = IO.Path.GetFileName(chemin)
        Catch ex As ArgumentException
            ' Chemin exotique : le nom reste vide, ce qui n'empêche pas la trace.
            nom = String.Empty
        End Try

        Try
            Using connexion As SqlConnection = WURepository.CreerConnexion()
                connexion.Open()

                Using commande As New SqlCommand(insertion, connexion)

                    commande.Parameters.Add("@jour", SqlDbType.Date).Value = dateActivite.Date
                    commande.Parameters.Add("@valeur", SqlDbType.Date).Value = dateValeur.Date
                    commande.Parameters.Add("@lot", SqlDbType.NVarChar, 10).Value = If(numeroLot, String.Empty)
                    commande.Parameters.Add("@nom", SqlDbType.NVarChar, 255).Value = nom
                    commande.Parameters.Add("@chemin", SqlDbType.NVarChar, 500).Value = Tronquer(chemin, 500)
                    commande.Parameters.Add("@lignes", SqlDbType.Int).Value = nombreLignes
                    commande.Parameters.Add("@debit", SqlDbType.BigInt).Value = totalDebit
                    commande.Parameters.Add("@credit", SqlDbType.BigInt).Value = totalCredit
                    commande.Parameters.Add("@auteur", SqlDbType.NVarChar, 50).Value = SessionWU.Auteur

                    commande.ExecuteNonQuery()
                End Using
            End Using

        Catch ex As SqlException
            messageErreur = If(ex.Number = ERREUR_TABLE_ABSENTE,
                               MESSAGE_TABLE_ABSENTE,
                               $"Production non tracée : {ex.Message}")
            Return False

        Catch ex As InvalidOperationException
            messageErreur = $"Connexion SQL Server indisponible : {ex.Message}"
            Return False
        End Try

        Return True
    End Function

    ''' <summary>Un chemin trop long est coupé plutôt que de faire échouer l'insertion.</summary>
    Private Shared Function Tronquer(texte As String, longueur As Integer) As String
        If texte.Length <= longueur Then Return texte
        Return texte.Substring(0, longueur)
    End Function

#End Region

#Region "Lecture"

    ''' <summary>
    ''' La dernière production connue pour une journée, ou Nothing s'il n'y en a aucune.
    '''
    ''' « Aucune production connue » n'est pas « jamais produit » : les fichiers sortis avant
    ''' la mise en service de cette table ne sont pas tracés. Les écrans le disent ainsi.
    '''
    ''' Une table absente n'alarme pas : l'application se comporte comme avant la trace,
    ''' c'est-à-dire sans avertissement, plutôt que de bloquer une comptabilisation.
    ''' </summary>
    Public Shared Function DerniereProduction(jour As Date) As Production

        Const lecture As String =
            "SELECT TOP 1 DateActivite, DateValeur, NumeroLot, NomFichier, NombreLignes, " &
            "       TotalDebit, DateProduction, ProduitPar " &
            "FROM   " & TABLE & " " &
            "WHERE  DateActivite = @jour " &
            "ORDER BY DateProduction DESC, IdFichier DESC"

        Try
            Using connexion As SqlConnection = WURepository.CreerConnexion()
                connexion.Open()

                Using commande As New SqlCommand(lecture, connexion)
                    commande.Parameters.Add("@jour", SqlDbType.Date).Value = jour.Date

                    Using lecteur As SqlDataReader = commande.ExecuteReader()

                        If Not lecteur.Read() Then Return Nothing

                        Return New Production() With {
                            .DateActivite = lecteur.GetDateTime(0),
                            .DateValeur = lecteur.GetDateTime(1),
                            .NumeroLot = LireChaine(lecteur, 2),
                            .NomFichier = LireChaine(lecteur, 3),
                            .NombreLignes = LireEntier(lecteur, 4),
                            .TotalDebit = LireEntierLong(lecteur, 5),
                            .DateProduction = lecteur.GetDateTime(6),
                            .ProduitPar = LireChaine(lecteur, 7)
                        }
                    End Using
                End Using
            End Using

        Catch ex As SqlException
            ' Table absente ou base injoignable : pas d'avertissement plutôt qu'un faux.
            Return Nothing
        Catch ex As InvalidOperationException
            Return Nothing
        End Try
    End Function

#End Region

#Region "Utilitaires de lecture"

    Private Shared Function LireChaine(lecteur As SqlDataReader, position As Integer) As String
        If lecteur.IsDBNull(position) Then Return String.Empty
        Return lecteur.GetValue(position).ToString().Trim()
    End Function

    Private Shared Function LireEntier(lecteur As SqlDataReader, position As Integer) As Integer
        If lecteur.IsDBNull(position) Then Return 0
        Return Convert.ToInt32(lecteur.GetValue(position), Globalization.CultureInfo.InvariantCulture)
    End Function

    Private Shared Function LireEntierLong(lecteur As SqlDataReader, position As Integer) As Long
        If lecteur.IsDBNull(position) Then Return 0L
        Return Convert.ToInt64(lecteur.GetValue(position), Globalization.CultureInfo.InvariantCulture)
    End Function

#End Region

End Class
