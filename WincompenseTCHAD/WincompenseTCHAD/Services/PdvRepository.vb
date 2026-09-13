Option Strict On
Option Explicit On

Imports System.Data
Imports System.Data.SqlClient

''' <summary>
''' Gestion (création, lecture, modification, suppression) des points de vente Western Union
''' dans SQL Server : sous-agents (T_Pdv_SA) et agences propres Ecobank (T_Pdv_EC).
'''
''' Séparé de WURepository, qui ne fait que LIRE ces tables pour la compensation quotidienne :
''' le chemin de comptabilisation ne doit jamais pouvoir écrire dans le paramétrage.
'''
''' Aucune exception SQL ne remonte à l'interface : chaque fonction retourne un booléen de
''' réussite et pose un message explicite, en français, dans son paramètre messageErreur.
'''
''' Les champs laissés vides sont écrits comme chaînes vides, jamais comme NULL : plusieurs
''' colonnes facultatives du point de vue de la saisie (GroupeStatistique, codeagence,
''' [CodeAgenc-Voyager]) sont déclarées NOT NULL en base, où un NULL ferait échouer l'écriture.
''' La lecture, elle, traite indifféremment NULL et chaîne vide.
''' </summary>
Public NotInheritable Class PdvRepository

    Private Sub New()
    End Sub

    ''' <summary>Codes d'erreur SQL Server signalant une violation de clé primaire ou d'unicité.</summary>
    Private Const ERREUR_CLE_DUPLIQUEE As Integer = 2627
    Private Const ERREUR_INDEX_UNIQUE As Integer = 2601

#Region "Sous-agents (T_Pdv_SA)"

    ''' <summary>
    ''' Liste les sous-agents, éventuellement restreints à ceux dont l'Account ou la désignation
    ''' contient le filtre saisi. Retourne une liste vide (jamais Nothing) en cas d'erreur.
    ''' </summary>
    ''' <param name="filtre">Texte recherché, ou chaîne vide pour tout lister.</param>
    Public Shared Function ListerSousAgents(filtre As String, ByRef messageErreur As String) As List(Of PointDeVenteSA)

        messageErreur = String.Empty
        Dim resultat As New List(Of PointDeVenteSA)

        Dim requete As String =
            "SELECT Code_Pdv, Designationagence, GroupeStatistique, Taux, " &
            "CompteCompense, CompteCommission, codeagence FROM T_Pdv_SA"

        Dim recherche As String = If(filtre, String.Empty).Trim()
        If recherche.Length > 0 Then
            requete &= " WHERE Code_Pdv LIKE @filtre OR Designationagence LIKE @filtre"
        End If
        requete &= " ORDER BY Code_Pdv"

        Try
            Using connexion As SqlConnection = WURepository.CreerConnexion()
                connexion.Open()

                Using commande As New SqlCommand(requete, connexion)

                    If recherche.Length > 0 Then
                        ' Les caractères génériques de LIKE sont neutralisés : une recherche sur
                        ' « % » doit chercher un pourcentage, pas tout ramener.
                        commande.Parameters.Add("@filtre", SqlDbType.NVarChar, 255).Value = "%" & EchapperLike(recherche) & "%"
                    End If

                    Using lecteur As SqlDataReader = commande.ExecuteReader()
                        While lecteur.Read()
                            resultat.Add(New PointDeVenteSA() With {
                                .CodePdv = LireChaine(lecteur, "Code_Pdv"),
                                .Designation = LireChaine(lecteur, "Designationagence"),
                                .GroupeStatistique = LireChaine(lecteur, "GroupeStatistique"),
                                .Taux = LireDecimal(lecteur, "Taux"),
                                .CompteCompense = LireChaine(lecteur, "CompteCompense"),
                                .CompteCommission = LireChaine(lecteur, "CompteCommission"),
                                .CodeAgence = LireChaine(lecteur, "codeagence")
                            })
                        End While
                    End Using
                End Using
            End Using

        Catch ex As SqlException
            messageErreur = $"Lecture des sous-agents impossible : {ex.Message}"
        Catch ex As InvalidOperationException
            messageErreur = $"Connexion SQL Server indisponible : {ex.Message}"
        End Try

        Return resultat
    End Function

    ''' <summary>Crée un sous-agent. Échoue si l'Account existe déjà dans T_Pdv_SA.</summary>
    Public Shared Function AjouterSousAgent(pdv As PointDeVenteSA, ByRef messageErreur As String) As Boolean

        messageErreur = String.Empty
        If pdv Is Nothing Then
            messageErreur = "Aucune fiche à enregistrer."
            Return False
        End If

        Const requete As String =
            "INSERT INTO T_Pdv_SA (Code_Pdv, Designationagence, GroupeStatistique, Taux, " &
            "CompteCompense, CompteCommission, codeagence) " &
            "VALUES (@code, @designation, @groupe, @taux, @compense, @commission, @codeagence)"

        Return ExecuterEcritureSA(requete, pdv, pdv.CodePdv, "créer", messageErreur)
    End Function

    ''' <summary>
    ''' Modifie un sous-agent existant, désigné par son Account. L'Account lui-même n'est jamais
    ''' modifié : c'est la clé sous laquelle les rapports Western Union désignent le point de
    ''' vente, la renommer reviendrait à perdre le lien avec l'historique.
    ''' </summary>
    Public Shared Function ModifierSousAgent(pdv As PointDeVenteSA, ByRef messageErreur As String) As Boolean

        messageErreur = String.Empty
        If pdv Is Nothing Then
            messageErreur = "Aucune fiche à enregistrer."
            Return False
        End If

        Const requete As String =
            "UPDATE T_Pdv_SA SET Designationagence = @designation, GroupeStatistique = @groupe, " &
            "Taux = @taux, CompteCompense = @compense, CompteCommission = @commission, " &
            "codeagence = @codeagence WHERE Code_Pdv = @code"

        Return ExecuterEcritureSA(requete, pdv, pdv.CodePdv, "modifier", messageErreur)
    End Function

    ''' <summary>Supprime un sous-agent par son Account.</summary>
    Public Shared Function SupprimerSousAgent(codePdv As String, ByRef messageErreur As String) As Boolean
        Return Supprimer("T_Pdv_SA", "Code_Pdv", codePdv, "sous-agent", messageErreur)
    End Function

    ''' <summary>Exécute un INSERT ou un UPDATE sur T_Pdv_SA avec les mêmes paramètres.</summary>
    Private Shared Function ExecuterEcritureSA(requete As String, pdv As PointDeVenteSA,
                                               code As String, action As String,
                                               ByRef messageErreur As String) As Boolean
        Try
            Using connexion As SqlConnection = WURepository.CreerConnexion()
                connexion.Open()

                Using commande As New SqlCommand(requete, connexion)
                    commande.Parameters.Add("@code", SqlDbType.NVarChar, 255).Value = pdv.CodePdv
                    commande.Parameters.Add("@designation", SqlDbType.NVarChar, 255).Value = pdv.Designation
                    commande.Parameters.Add("@groupe", SqlDbType.NVarChar, 255).Value = pdv.GroupeStatistique
                    ' La colonne Taux est de type DECIMAL(4,2) : précision et échelle sont cadrées
                    ' explicitement, et la valeur arrondie à deux décimales côté application, afin
                    ' qu'aucun arrondi ne se produise en silence dans SQL Server.
                    Dim parametreTaux As SqlParameter = commande.Parameters.Add("@taux", SqlDbType.Decimal)
                    parametreTaux.Precision = 4
                    parametreTaux.Scale = 2
                    parametreTaux.Value = Decimal.Round(pdv.Taux, 2)
                    commande.Parameters.Add("@compense", SqlDbType.NVarChar, 255).Value = pdv.CompteCompense
                    commande.Parameters.Add("@commission", SqlDbType.NVarChar, 255).Value = pdv.CompteCommission
                    commande.Parameters.Add("@codeagence", SqlDbType.NVarChar, 255).Value = pdv.CodeAgence

                    If commande.ExecuteNonQuery() = 0 Then
                        messageErreur = $"Aucune fiche modifiée : le sous-agent « {code} » n'existe plus dans T_Pdv_SA." &
                                        Environment.NewLine & "Il a peut-être été supprimé depuis un autre poste."
                        Return False
                    End If
                End Using
            End Using

        Catch ex As SqlException
            messageErreur = MessageErreurEcriture(ex, code, "sous-agent", action)
            Return False
        Catch ex As InvalidOperationException
            messageErreur = $"Connexion SQL Server indisponible : {ex.Message}"
            Return False
        End Try

        Return True
    End Function

    ''' <summary>
    ''' Liste les groupes statistiques DÉJÀ utilisés par au moins un sous-agent.
    '''
    ''' Il n'existe pas de table de groupes : un groupe n'a d'existence que par les sous-agents
    ''' qui le portent. « Créer un groupe » revient donc simplement à saisir un libellé encore
    ''' inconnu — et « supprimer » le dernier sous-agent d'un groupe le fait disparaître de
    ''' cette liste. C'est précisément pourquoi la saisie doit passer par une sélection : sans
    ''' cela, la moindre faute de frappe crée un groupe parasite impossible à distinguer.
    ''' </summary>
    ''' <returns>Libellés distincts, triés. Liste vide (jamais Nothing) en cas d'erreur.</returns>
    Public Shared Function ListerGroupesStatistiques(ByRef messageErreur As String) As List(Of String)

        messageErreur = String.Empty
        Dim resultat As New List(Of String)

        Const requete As String =
            "SELECT DISTINCT LTRIM(RTRIM(GroupeStatistique)) AS Groupe FROM T_Pdv_SA " &
            "WHERE GroupeStatistique IS NOT NULL AND LTRIM(RTRIM(GroupeStatistique)) <> '' " &
            "ORDER BY Groupe"

        Try
            Using connexion As SqlConnection = WURepository.CreerConnexion()
                connexion.Open()

                Using commande As New SqlCommand(requete, connexion)
                    Using lecteur As SqlDataReader = commande.ExecuteReader()
                        While lecteur.Read()
                            Dim groupe As String = LireChaine(lecteur, "Groupe")
                            If groupe.Length > 0 Then resultat.Add(groupe)
                        End While
                    End Using
                End Using
            End Using

        Catch ex As SqlException
            messageErreur = $"Lecture des groupes statistiques impossible : {ex.Message}"
        Catch ex As InvalidOperationException
            messageErreur = $"Connexion SQL Server indisponible : {ex.Message}"
        End Try

        Return resultat
    End Function

#End Region

#Region "Agences propres Ecobank (T_Pdv_EC)"

    ''' <summary>Liste les agences propres, éventuellement filtrées sur l'Account ou la désignation.</summary>
    Public Shared Function ListerAgences(filtre As String, ByRef messageErreur As String) As List(Of PointDeVenteEC)

        messageErreur = String.Empty
        Dim resultat As New List(Of PointDeVenteEC)

        Dim requete As String = "SELECT Codesite, Designationagence, [CodeAgenc-Voyager] FROM T_Pdv_EC"

        Dim recherche As String = If(filtre, String.Empty).Trim()
        If recherche.Length > 0 Then
            requete &= " WHERE Codesite LIKE @filtre OR Designationagence LIKE @filtre"
        End If
        requete &= " ORDER BY Codesite"

        Try
            Using connexion As SqlConnection = WURepository.CreerConnexion()
                connexion.Open()

                Using commande As New SqlCommand(requete, connexion)

                    If recherche.Length > 0 Then
                        commande.Parameters.Add("@filtre", SqlDbType.NVarChar, 255).Value = "%" & EchapperLike(recherche) & "%"
                    End If

                    Using lecteur As SqlDataReader = commande.ExecuteReader()
                        While lecteur.Read()
                            resultat.Add(New PointDeVenteEC() With {
                                .CodeSite = LireChaine(lecteur, "Codesite"),
                                .Designation = LireChaine(lecteur, "Designationagence"),
                                .CodeAgenceVoyager = LireChaine(lecteur, "CodeAgenc-Voyager")
                            })
                        End While
                    End Using
                End Using
            End Using

        Catch ex As SqlException
            messageErreur = $"Lecture des agences impossible : {ex.Message}"
        Catch ex As InvalidOperationException
            messageErreur = $"Connexion SQL Server indisponible : {ex.Message}"
        End Try

        Return resultat
    End Function

    ''' <summary>Crée une agence propre. Échoue si l'Account existe déjà dans T_Pdv_EC.</summary>
    Public Shared Function AjouterAgence(pdv As PointDeVenteEC, ByRef messageErreur As String) As Boolean

        messageErreur = String.Empty
        If pdv Is Nothing Then
            messageErreur = "Aucune fiche à enregistrer."
            Return False
        End If

        Const requete As String =
            "INSERT INTO T_Pdv_EC (Codesite, Designationagence, [CodeAgenc-Voyager]) " &
            "VALUES (@code, @designation, @voyager)"

        Return ExecuterEcritureEC(requete, pdv, "créer", messageErreur)
    End Function

    ''' <summary>Modifie une agence propre existante, désignée par son Account (jamais modifié).</summary>
    Public Shared Function ModifierAgence(pdv As PointDeVenteEC, ByRef messageErreur As String) As Boolean

        messageErreur = String.Empty
        If pdv Is Nothing Then
            messageErreur = "Aucune fiche à enregistrer."
            Return False
        End If

        Const requete As String =
            "UPDATE T_Pdv_EC SET Designationagence = @designation, [CodeAgenc-Voyager] = @voyager " &
            "WHERE Codesite = @code"

        Return ExecuterEcritureEC(requete, pdv, "modifier", messageErreur)
    End Function

    ''' <summary>Supprime une agence propre par son Account.</summary>
    Public Shared Function SupprimerAgence(codeSite As String, ByRef messageErreur As String) As Boolean
        Return Supprimer("T_Pdv_EC", "Codesite", codeSite, "agence", messageErreur)
    End Function

    Private Shared Function ExecuterEcritureEC(requete As String, pdv As PointDeVenteEC,
                                               action As String, ByRef messageErreur As String) As Boolean
        Try
            Using connexion As SqlConnection = WURepository.CreerConnexion()
                connexion.Open()

                Using commande As New SqlCommand(requete, connexion)
                    commande.Parameters.Add("@code", SqlDbType.NVarChar, 255).Value = pdv.CodeSite
                    commande.Parameters.Add("@designation", SqlDbType.NVarChar, 255).Value = pdv.Designation
                    commande.Parameters.Add("@voyager", SqlDbType.NVarChar, 255).Value = pdv.CodeAgenceVoyager

                    If commande.ExecuteNonQuery() = 0 Then
                        messageErreur = $"Aucune fiche modifiée : l'agence « {pdv.CodeSite} » n'existe plus dans T_Pdv_EC." &
                                        Environment.NewLine & "Elle a peut-être été supprimée depuis un autre poste."
                        Return False
                    End If
                End Using
            End Using

        Catch ex As SqlException
            messageErreur = MessageErreurEcriture(ex, pdv.CodeSite, "agence", action)
            Return False
        Catch ex As InvalidOperationException
            messageErreur = $"Connexion SQL Server indisponible : {ex.Message}"
            Return False
        End Try

        Return True
    End Function

#End Region

#Region "Contrôle d'intégrité entre les deux tables"

    ''' <summary>
    ''' Indique si un Account figure déjà dans l'AUTRE table de points de vente.
    '''
    ''' Un même Account présent dans T_Pdv_SA et T_Pdv_EC est une incohérence de paramétrage :
    ''' la recherche interroge T_Pdv_SA en premier, l'entrée agence propre serait donc
    ''' silencieusement ignorée et le point de vente traité comme un sous-agent. La base
    ''' n'interdit pas ce doublon — seule l'application peut le signaler.
    ''' </summary>
    ''' <returns>True si l'Account existe dans l'autre table. False en cas d'erreur (non bloquant).</returns>
    Public Shared Function ExisteDansAutreTable(account As String, estSousAgent As Boolean) As Boolean

        If String.IsNullOrWhiteSpace(account) Then Return False

        Dim requete As String = If(estSousAgent,
                                   "SELECT COUNT(*) FROM T_Pdv_EC WHERE Codesite = @code",
                                   "SELECT COUNT(*) FROM T_Pdv_SA WHERE Code_Pdv = @code")

        Try
            Using connexion As SqlConnection = WURepository.CreerConnexion()
                connexion.Open()

                Using commande As New SqlCommand(requete, connexion)
                    commande.Parameters.Add("@code", SqlDbType.NVarChar, 255).Value = account.Trim()

                    Dim valeur As Object = commande.ExecuteScalar()
                    If valeur Is Nothing OrElse valeur Is DBNull.Value Then Return False

                    Return Convert.ToInt32(valeur, Globalization.CultureInfo.InvariantCulture) > 0
                End Using
            End Using

        Catch ex As SqlException
            ' Contrôle de confort : son échec ne doit jamais empêcher un enregistrement.
            Return False
        Catch ex As InvalidOperationException
            Return False
        End Try
    End Function

#End Region

#Region "Utilitaires internes"

    ''' <summary>Suppression générique d'une fiche par sa clé, pour les deux tables.</summary>
    Private Shared Function Supprimer(table As String, colonneCle As String, valeurCle As String,
                                      libelle As String, ByRef messageErreur As String) As Boolean

        messageErreur = String.Empty

        If String.IsNullOrWhiteSpace(valeurCle) Then
            messageErreur = "Aucune fiche sélectionnée."
            Return False
        End If

        Try
            Using connexion As SqlConnection = WURepository.CreerConnexion()
                connexion.Open()

                Using commande As New SqlCommand($"DELETE FROM {table} WHERE {colonneCle} = @code", connexion)
                    commande.Parameters.Add("@code", SqlDbType.NVarChar, 255).Value = valeurCle.Trim()

                    If commande.ExecuteNonQuery() = 0 Then
                        messageErreur = $"Aucune suppression : le {libelle} « {valeurCle} » n'existe plus dans {table}."
                        Return False
                    End If
                End Using
            End Using

        Catch ex As SqlException
            messageErreur = $"Suppression du {libelle} « {valeurCle} » impossible : {ex.Message}"
            Return False
        Catch ex As InvalidOperationException
            messageErreur = $"Connexion SQL Server indisponible : {ex.Message}"
            Return False
        End Try

        Return True
    End Function

    ''' <summary>
    ''' Traduit une exception SQL d'écriture en message compréhensible. Le cas le plus fréquent
    ''' — l'Account existe déjà — mérite une explication, pas un code d'erreur SQL Server.
    ''' </summary>
    Private Shared Function MessageErreurEcriture(ex As SqlException, code As String,
                                                  libelle As String, action As String) As String

        If ex.Number = ERREUR_CLE_DUPLIQUEE OrElse ex.Number = ERREUR_INDEX_UNIQUE Then
            Return $"L'Account « {code} » est déjà enregistré comme {libelle}." & Environment.NewLine &
                   "Chaque Account ne peut figurer qu'une seule fois : sélectionnez la fiche existante pour la modifier."
        End If

        Return $"Impossible de {action} le {libelle} « {code} » : {ex.Message}"
    End Function

    ''' <summary>
    ''' Neutralise les caractères génériques de LIKE (%, _, [) dans un texte de recherche,
    ''' au moyen d'une clause ESCAPE implicite par crochets.
    ''' </summary>
    Private Shared Function EchapperLike(texte As String) As String
        Return texte.Replace("[", "[[]").Replace("%", "[%]").Replace("_", "[_]")
    End Function

    Private Shared Function LireChaine(lecteur As SqlDataReader, colonne As String) As String
        Dim index As Integer = lecteur.GetOrdinal(colonne)
        If lecteur.IsDBNull(index) Then Return String.Empty
        Return lecteur.GetValue(index).ToString().Trim()
    End Function

    Private Shared Function LireDecimal(lecteur As SqlDataReader, colonne As String) As Decimal
        Dim index As Integer = lecteur.GetOrdinal(colonne)
        If lecteur.IsDBNull(index) Then Return 0D
        Return WUReportService.ToDecimalSafe(lecteur.GetValue(index))
    End Function

#End Region

End Class
