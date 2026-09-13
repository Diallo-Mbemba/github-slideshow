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
            "CompteCompense, CompteCommission, codeagence, DateCreation, CreePar) " &
            "VALUES (@code, @designation, @groupe, @taux, @compense, @commission, @codeagence, " &
            "GETDATE(), @auteur)"

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
            "codeagence = @codeagence, DateModification = GETDATE(), ModifiePar = @auteur " &
            "WHERE Code_Pdv = @code"

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
                    commande.Parameters.Add("@auteur", SqlDbType.NVarChar, 50).Value = SessionWU.Auteur

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

#End Region

#Region "Groupes statistiques (T_GroupeStatistique)"

    ''' <summary>Code d'erreur SQL Server signalant une table absente (« Invalid object name »).</summary>
    Private Const ERREUR_TABLE_ABSENTE As Integer = 208

    ''' <summary>
    ''' Message posé lorsque la table des groupes n'existe pas encore : la migration n'a pas
    ''' été jouée. Le dire explicitement évite de laisser croire à une panne de la base.
    ''' </summary>
    Public Const MESSAGE_TABLE_GROUPES_ABSENTE As String =
        "La table T_GroupeStatistique n'existe pas encore dans la base." & vbCrLf & vbCrLf &
        "Exécutez le script Scripts\04_GroupeStatistique.sql : il crée la table et y reprend " &
        "automatiquement les groupes déjà présents dans T_Pdv_SA."

    ''' <summary>
    ''' Liste les groupes statistiques : ceux de T_GroupeStatistique, ET ceux qui ne sont encore
    ''' portés que par des sous-agents de T_Pdv_SA — les groupes « hérités ».
    '''
    ''' Ne lister que la table condamnerait l'exploitation tant que la migration n'a pas abouti :
    ''' les anciens groupes disparaîtraient des listes déroulantes et plus aucun sous-agent ne
    ''' pourrait leur être rattaché. Les groupes hérités sont donc repris avec les valeurs lues
    ''' dans T_Pdv_SA et marqués EstEnregistre = False ; une confirmation suffit à les enregistrer.
    '''
    ''' Si T_GroupeStatistique n'existe pas encore, seuls les groupes hérités sont retournés :
    ''' l'application reste pleinement utilisable avant la migration.
    '''
    ''' Pour les groupes enregistrés, NombreDesynchronises compte les sous-agents dont les
    ''' colonnes de T_Pdv_SA — lues par la comptabilisation quotidienne — ont dérivé des valeurs
    ''' du groupe.
    ''' </summary>
    ''' <param name="filtre">Texte recherché dans le libellé ou les comptes, ou chaîne vide.</param>
    Public Shared Function ListerGroupes(filtre As String, ByRef messageErreur As String) As List(Of GroupeStatistiqueWU)

        messageErreur = String.Empty

        Dim recherche As String = If(filtre, String.Empty).Trim()

        Try
            Return ExecuterListeGroupes(recherche, avecTableGroupes:=True)

        Catch ex As SqlException When ex.Number = ERREUR_TABLE_ABSENTE
            ' Migration non jouée : on se rabat sur les seuls groupes hérités de T_Pdv_SA.
            Try
                Dim herites As List(Of GroupeStatistiqueWU) = ExecuterListeGroupes(recherche, avecTableGroupes:=False)
                messageErreur = String.Empty
                Return herites
            Catch exSecours As SqlException
                messageErreur = $"Lecture des groupes statistiques impossible : {exSecours.Message}"
            Catch exSecours As InvalidOperationException
                messageErreur = $"Connexion SQL Server indisponible : {exSecours.Message}"
            End Try

        Catch ex As SqlException
            messageErreur = $"Lecture des groupes statistiques impossible : {ex.Message}"
        Catch ex As InvalidOperationException
            messageErreur = $"Connexion SQL Server indisponible : {ex.Message}"
        End Try

        Return New List(Of GroupeStatistiqueWU)
    End Function

    ''' <summary>
    ''' Exécute la lecture des groupes. Avec avecTableGroupes = False, la requête ne référence
    ''' pas du tout T_GroupeStatistique : c'est ce qui permet de fonctionner avant la migration.
    ''' Les exceptions ne sont PAS capturées ici : l'appelant décide du repli.
    ''' </summary>
    Private Shared Function ExecuterListeGroupes(recherche As String, avecTableGroupes As Boolean) As List(Of GroupeStatistiqueWU)

        Dim resultat As New List(Of GroupeStatistiqueWU)

        ' Valeurs des groupes tels que portés par les sous-agents : servent aux groupes hérités.
        Const groupesPdv As String =
            "SELECT LTRIM(RTRIM(GroupeStatistique)) AS Groupe, " &
            "       MIN(CompteCompense) AS CompteActivite, " &
            "       MIN(CompteCommission) AS CompteCommission, " &
            "       MIN(Taux) AS Taux, COUNT(*) AS NombreSousAgents " &
            "FROM T_Pdv_SA " &
            "WHERE GroupeStatistique IS NOT NULL AND LTRIM(RTRIM(GroupeStatistique)) <> '' " &
            "GROUP BY LTRIM(RTRIM(GroupeStatistique))"

        Dim requete As String

        If avecTableGroupes Then
            requete =
                "SELECT g.Groupe, g.CompteActivite, g.CompteCommission, g.Taux, " &
                "  (SELECT COUNT(*) FROM T_Pdv_SA p " &
                "    WHERE LTRIM(RTRIM(p.GroupeStatistique)) = g.Groupe) AS NombreSousAgents, " &
                "  (SELECT COUNT(*) FROM T_Pdv_SA p " &
                "    WHERE LTRIM(RTRIM(p.GroupeStatistique)) = g.Groupe " &
                "      AND (p.CompteCompense <> g.CompteActivite " &
                "        OR p.CompteCommission <> g.CompteCommission " &
                "        OR p.Taux <> g.Taux)) AS NombreDesynchronises, " &
                "  1 AS EstEnregistre " &
                "FROM T_GroupeStatistique g " &
                "UNION ALL " &
                "SELECT h.Groupe, h.CompteActivite, h.CompteCommission, h.Taux, " &
                "       h.NombreSousAgents, 0 AS NombreDesynchronises, 0 AS EstEnregistre " &
                $"FROM ({groupesPdv}) h " &
                "WHERE NOT EXISTS (SELECT 1 FROM T_GroupeStatistique g2 WHERE g2.Groupe = h.Groupe)"
        Else
            requete =
                "SELECT h.Groupe, h.CompteActivite, h.CompteCommission, h.Taux, " &
                "       h.NombreSousAgents, 0 AS NombreDesynchronises, 0 AS EstEnregistre " &
                $"FROM ({groupesPdv}) h"
        End If

        ' Le filtre est appliqué autour de l'union : il porte de la même façon sur les deux sources.
        If recherche.Length > 0 Then
            requete = $"SELECT * FROM ({requete}) t " &
                      "WHERE t.Groupe LIKE @filtre OR t.CompteActivite LIKE @filtre " &
                      "OR t.CompteCommission LIKE @filtre ORDER BY t.Groupe"
        Else
            requete = $"SELECT * FROM ({requete}) t ORDER BY t.Groupe"
        End If

        Using connexion As SqlConnection = WURepository.CreerConnexion()
            connexion.Open()

            Using commande As New SqlCommand(requete, connexion)

                If recherche.Length > 0 Then
                    commande.Parameters.Add("@filtre", SqlDbType.NVarChar, 255).Value = "%" & EchapperLike(recherche) & "%"
                End If

                Using lecteur As SqlDataReader = commande.ExecuteReader()
                    While lecteur.Read()
                        resultat.Add(New GroupeStatistiqueWU() With {
                            .Nom = LireChaine(lecteur, "Groupe"),
                            .CompteActivite = LireChaine(lecteur, "CompteActivite"),
                            .CompteCommission = LireChaine(lecteur, "CompteCommission"),
                            .Taux = LireDecimal(lecteur, "Taux"),
                            .NombreSousAgents = LireEntier(lecteur, "NombreSousAgents"),
                            .NombreDesynchronises = LireEntier(lecteur, "NombreDesynchronises"),
                            .EstEnregistre = LireEntier(lecteur, "EstEnregistre") <> 0
                        })
                    End While
                End Using
            End Using
        End Using

        Return resultat
    End Function

    ''' <summary>Crée un groupe statistique.</summary>
    Public Shared Function AjouterGroupe(groupe As GroupeStatistiqueWU, ByRef messageErreur As String) As Boolean

        messageErreur = String.Empty
        If groupe Is Nothing Then
            messageErreur = "Aucun groupe à enregistrer."
            Return False
        End If

        Const requete As String =
            "INSERT INTO T_GroupeStatistique (Groupe, CompteActivite, CompteCommission, Taux, " &
            "DateCreation, CreePar) " &
            "VALUES (@groupe, @activite, @commission, @taux, GETDATE(), @auteur)"

        Return ExecuterEcritureGroupe(requete, groupe, "créer", messageErreur)
    End Function

    ''' <summary>
    ''' Modifie un groupe existant, désigné par son libellé. Le libellé lui-même n'est jamais
    ''' modifié ici : il est la clé sous laquelle les sous-agents se rattachent au groupe.
    ''' Pour renommer un groupe, voir RenommerGroupe.
    ''' </summary>
    Public Shared Function ModifierGroupe(groupe As GroupeStatistiqueWU, ByRef messageErreur As String) As Boolean

        messageErreur = String.Empty
        If groupe Is Nothing Then
            messageErreur = "Aucun groupe à enregistrer."
            Return False
        End If

        Const requete As String =
            "UPDATE T_GroupeStatistique SET CompteActivite = @activite, " &
            "CompteCommission = @commission, Taux = @taux, " &
            "DateModification = GETDATE(), ModifiePar = @auteur WHERE Groupe = @groupe"

        Return ExecuterEcritureGroupe(requete, groupe, "modifier", messageErreur)
    End Function

    ''' <summary>
    ''' Supprime un groupe. Refusé s'il reste des sous-agents rattachés : ils perdraient leur
    ''' compte d'activité, leur compte de commission et leur taux sans qu'on s'en aperçoive.
    ''' </summary>
    Public Shared Function SupprimerGroupe(nomGroupe As String, ByRef messageErreur As String) As Boolean

        messageErreur = String.Empty

        If String.IsNullOrWhiteSpace(nomGroupe) Then
            messageErreur = "Aucun groupe sélectionné."
            Return False
        End If

        Try
            Using connexion As SqlConnection = WURepository.CreerConnexion()
                connexion.Open()

                Dim rattaches As Integer = CompterSousAgentsDuGroupe(nomGroupe, connexion)
                If rattaches > 0 Then
                    messageErreur = $"Le groupe « {nomGroupe} » ne peut pas être supprimé : " &
                                    $"{rattaches} sous-agent(s) y sont rattachés." & Environment.NewLine & Environment.NewLine &
                                    "Rattachez-les d'abord à un autre groupe."
                    Return False
                End If

                Using commande As New SqlCommand("DELETE FROM T_GroupeStatistique WHERE Groupe = @groupe", connexion)
                    commande.Parameters.Add("@groupe", SqlDbType.NVarChar, 255).Value = nomGroupe.Trim()

                    If commande.ExecuteNonQuery() = 0 Then
                        messageErreur = $"Aucune suppression : le groupe « {nomGroupe} » n'existe plus."
                        Return False
                    End If
                End Using
            End Using

        Catch ex As SqlException
            messageErreur = If(ex.Number = ERREUR_TABLE_ABSENTE,
                               MESSAGE_TABLE_GROUPES_ABSENTE,
                               $"Suppression du groupe « {nomGroupe} » impossible : {ex.Message}")
            Return False
        Catch ex As InvalidOperationException
            messageErreur = $"Connexion SQL Server indisponible : {ex.Message}"
            Return False
        End Try

        Return True
    End Function

    ''' <summary>
    ''' Reporte les valeurs d'un groupe sur TOUS les sous-agents qui le portent, afin que les
    ''' colonnes de T_Pdv_SA — lues par la comptabilisation quotidienne et par d'autres
    ''' applications — restent le reflet exact du groupe.
    ''' </summary>
    ''' <param name="nombreModifies">Nombre de sous-agents effectivement mis à jour.</param>
    Public Shared Function SynchroniserSousAgentsDuGroupe(groupe As GroupeStatistiqueWU,
                                                          ByRef nombreModifies As Integer,
                                                          ByRef messageErreur As String) As Boolean

        messageErreur = String.Empty
        nombreModifies = 0

        If groupe Is Nothing OrElse String.IsNullOrWhiteSpace(groupe.Nom) Then
            messageErreur = "Aucun groupe à synchroniser."
            Return False
        End If

        Const requete As String =
            "UPDATE T_Pdv_SA SET CompteCompense = @activite, CompteCommission = @commission, " &
            "Taux = @taux, DateModification = GETDATE(), ModifiePar = @auteur " &
            "WHERE LTRIM(RTRIM(GroupeStatistique)) = @groupe"

        Try
            Using connexion As SqlConnection = WURepository.CreerConnexion()
                connexion.Open()

                Using commande As New SqlCommand(requete, connexion)
                    AjouterParametresGroupe(commande, groupe)
                    nombreModifies = commande.ExecuteNonQuery()
                End Using
            End Using

        Catch ex As SqlException
            messageErreur = $"Synchronisation des sous-agents du groupe « {groupe.Nom} » impossible : {ex.Message}"
            Return False
        Catch ex As InvalidOperationException
            messageErreur = $"Connexion SQL Server indisponible : {ex.Message}"
            Return False
        End Try

        Return True
    End Function

    Private Shared Function ExecuterEcritureGroupe(requete As String, groupe As GroupeStatistiqueWU,
                                                   action As String, ByRef messageErreur As String) As Boolean
        Try
            Using connexion As SqlConnection = WURepository.CreerConnexion()
                connexion.Open()

                Using commande As New SqlCommand(requete, connexion)
                    AjouterParametresGroupe(commande, groupe)

                    If commande.ExecuteNonQuery() = 0 Then
                        messageErreur = $"Aucune modification : le groupe « {groupe.Nom} » n'existe plus." &
                                        Environment.NewLine & "Il a peut-être été supprimé depuis un autre poste."
                        Return False
                    End If
                End Using
            End Using

        Catch ex As SqlException
            messageErreur = MessageErreurGroupe(ex, groupe, action)
            Return False
        Catch ex As InvalidOperationException
            messageErreur = $"Connexion SQL Server indisponible : {ex.Message}"
            Return False
        End Try

        Return True
    End Function

    Private Shared Sub AjouterParametresGroupe(commande As SqlCommand, groupe As GroupeStatistiqueWU)

        commande.Parameters.Add("@groupe", SqlDbType.NVarChar, 255).Value = groupe.Nom.Trim()
        commande.Parameters.Add("@activite", SqlDbType.NVarChar, 255).Value = groupe.CompteActivite
        commande.Parameters.Add("@commission", SqlDbType.NVarChar, 255).Value = groupe.CompteCommission

        ' La colonne Taux est de type DECIMAL(4,2) : précision et échelle cadrées explicitement,
        ' valeur arrondie côté application, pour qu'aucun arrondi ne se produise en silence.
        Dim parametreTaux As SqlParameter = commande.Parameters.Add("@taux", SqlDbType.Decimal)
        parametreTaux.Precision = 4
        parametreTaux.Scale = 2
        parametreTaux.Value = Decimal.Round(groupe.Taux, 2)

        commande.Parameters.Add("@auteur", SqlDbType.NVarChar, 50).Value = SessionWU.Auteur
    End Sub

    Private Shared Function CompterSousAgentsDuGroupe(nomGroupe As String, connexion As SqlConnection) As Integer
        Return CompterSousAgentsDuGroupe(nomGroupe, connexion, Nothing)
    End Function

    ''' <param name="transaction">Transaction en cours, ou Nothing hors transaction.</param>
    Private Shared Function CompterSousAgentsDuGroupe(nomGroupe As String, connexion As SqlConnection,
                                                      transaction As SqlTransaction) As Integer

        Using commande As New SqlCommand(
            "SELECT COUNT(*) FROM T_Pdv_SA WHERE LTRIM(RTRIM(GroupeStatistique)) = @groupe",
            connexion, transaction)

            commande.Parameters.Add("@groupe", SqlDbType.NVarChar, 255).Value = nomGroupe.Trim()

            Dim valeur As Object = commande.ExecuteScalar()
            If valeur Is Nothing OrElse valeur Is DBNull.Value Then Return 0

            Return Convert.ToInt32(valeur, Globalization.CultureInfo.InvariantCulture)
        End Using
    End Function

    ''' <summary>
    ''' Traduit une exception SQL d'écriture sur les groupes. Les deux index uniques posés sur
    ''' les comptes rendent la règle « un compte n'appartient qu'à un seul groupe » impossible à
    ''' violer : reste à dire lequel des deux comptes est déjà pris, plutôt qu'un code d'erreur.
    ''' </summary>
    Private Shared Function MessageErreurGroupe(ex As SqlException, groupe As GroupeStatistiqueWU, action As String) As String

        If ex.Number = ERREUR_TABLE_ABSENTE Then
            Return MESSAGE_TABLE_GROUPES_ABSENTE
        End If

        If ex.Number = ERREUR_CLE_DUPLIQUEE OrElse ex.Number = ERREUR_INDEX_UNIQUE Then

            ' Le message de SQL Server nomme l'index violé : il indique lequel des deux comptes
            ' — ou le libellé du groupe — est déjà utilisé ailleurs.
            If ex.Message.IndexOf("CompteActivite", StringComparison.OrdinalIgnoreCase) >= 0 Then
                Return $"Le compte d'activité « {groupe.CompteActivite} » appartient déjà à un autre groupe." &
                       Environment.NewLine & "Un compte ne peut appartenir qu'à un seul groupe statistique."
            End If

            If ex.Message.IndexOf("CompteCommission", StringComparison.OrdinalIgnoreCase) >= 0 Then
                Return $"Le compte de commission « {groupe.CompteCommission} » appartient déjà à un autre groupe." &
                       Environment.NewLine & "Un compte ne peut appartenir qu'à un seul groupe statistique."
            End If

            Return $"Le groupe « {groupe.Nom} » existe déjà." & Environment.NewLine &
                   "Sélectionnez-le dans la liste pour le modifier."
        End If

        Return $"Impossible de {action} le groupe « {groupe.Nom} » : {ex.Message}"
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
            "INSERT INTO T_Pdv_EC (Codesite, Designationagence, [CodeAgenc-Voyager], " &
            "DateCreation, CreePar) " &
            "VALUES (@code, @designation, @voyager, GETDATE(), @auteur)"

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
            "UPDATE T_Pdv_EC SET Designationagence = @designation, [CodeAgenc-Voyager] = @voyager, " &
            "DateModification = GETDATE(), ModifiePar = @auteur WHERE Codesite = @code"

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
                    commande.Parameters.Add("@auteur", SqlDbType.NVarChar, 50).Value = SessionWU.Auteur

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

#Region "Application d'une demande autorisée"

    '
    ' Ces méthodes sont le seul chemin d'écriture ouvert depuis l'arrivée du double regard.
    ' Elles s'exécutent dans la connexion et la transaction de la décision : l'écriture et
    ' le marquage de la demande tiennent ensemble, ou ni l'un ni l'autre.
    '
    ' Les colonnes de traçabilité reçoivent l'identifiant de CELUI QUI A SAISI, non de celui qui
    ' autorise : c'est lui l'auteur du contenu. Qui a autorisé se lit dans T_DemandeWU, qui
    ' conserve les deux noms — chaque table porte ainsi ce pour quoi elle est faite.
    '

    ''' <summary>Porte dans T_Pdv_SA une demande autorisée portant sur un sous-agent.</summary>
    Public Shared Function AppliquerSousAgent(demande As DemandeWU, connexion As SqlConnection,
                                              transaction As SqlTransaction,
                                              ByRef messageErreur As String) As Boolean

        messageErreur = String.Empty

        If demande.Operation = OperationWU.Suppression Then
            Return AppliquerSuppression("T_Pdv_SA", "Code_Pdv", demande.Cle, "sous-agent",
                                        connexion, transaction, messageErreur)
        End If

        Dim creation As Boolean = demande.Operation = OperationWU.Creation

        Dim requete As String =
            If(creation,
               "INSERT INTO T_Pdv_SA (Code_Pdv, Designationagence, GroupeStatistique, Taux, " &
               "CompteCompense, CompteCommission, codeagence, DateCreation, CreePar) " &
               "VALUES (@cle, @designation, @groupe, @taux, @activite, @commission, " &
               "@rattachement, GETDATE(), @auteur)",
               "UPDATE T_Pdv_SA SET Designationagence = @designation, GroupeStatistique = @groupe, " &
               "Taux = @taux, CompteCompense = @activite, CompteCommission = @commission, " &
               "codeagence = @rattachement, DateModification = GETDATE(), ModifiePar = @auteur " &
               "WHERE Code_Pdv = @cle")

        Return ExecuterApplication(requete, demande, creation, "sous-agent",
                                   connexion, transaction, messageErreur)
    End Function

    ''' <summary>Porte dans T_Pdv_EC une demande autorisée portant sur une agence propre.</summary>
    Public Shared Function AppliquerAgence(demande As DemandeWU, connexion As SqlConnection,
                                           transaction As SqlTransaction,
                                           ByRef messageErreur As String) As Boolean

        messageErreur = String.Empty

        If demande.Operation = OperationWU.Suppression Then
            Return AppliquerSuppression("T_Pdv_EC", "Codesite", demande.Cle, "agence",
                                        connexion, transaction, messageErreur)
        End If

        Dim creation As Boolean = demande.Operation = OperationWU.Creation

        Dim requete As String =
            If(creation,
               "INSERT INTO T_Pdv_EC (Codesite, Designationagence, [CodeAgenc-Voyager], " &
               "DateCreation, CreePar) " &
               "VALUES (@cle, @designation, @rattachement, GETDATE(), @auteur)",
               "UPDATE T_Pdv_EC SET Designationagence = @designation, " &
               "[CodeAgenc-Voyager] = @rattachement, DateModification = GETDATE(), " &
               "ModifiePar = @auteur WHERE Codesite = @cle")

        Return ExecuterApplication(requete, demande, creation, "agence",
                                   connexion, transaction, messageErreur)
    End Function

    ''' <summary>Porte une demande autorisée portant sur un groupe statistique.</summary>
    Public Shared Function AppliquerGroupe(demande As DemandeWU, connexion As SqlConnection,
                                           transaction As SqlTransaction,
                                           ByRef messageErreur As String) As Boolean

        messageErreur = String.Empty

        If demande.Operation = OperationWU.Synchronisation Then
            Return AppliquerSynchronisation(demande, connexion, transaction, messageErreur)
        End If

        If demande.Operation = OperationWU.Suppression Then

            ' Le garde-fou d'origine reste valable ici : supprimer un groupe encore porté par des
            ' sous-agents leur ferait perdre en silence leur compte d'activité, leur compte de
            ' commission et leur taux.
            Dim rattaches As Integer = CompterSousAgentsDuGroupe(demande.Cle, connexion, transaction)

            If rattaches > 0 Then
                messageErreur = $"Le groupe « {demande.Cle} » ne peut pas être supprimé : " &
                                $"{rattaches} sous-agent(s) y sont rattachés." & Environment.NewLine & Environment.NewLine &
                                "La demande ne peut pas être autorisée en l'état. Rattachez d'abord " &
                                "ces sous-agents à un autre groupe, ou rejetez la demande."
                Return False
            End If

            Return AppliquerSuppression("T_GroupeStatistique", "Groupe", demande.Cle, "groupe",
                                        connexion, transaction, messageErreur)
        End If

        Dim creation As Boolean = demande.Operation = OperationWU.Creation

        Dim requete As String =
            If(creation,
               "INSERT INTO T_GroupeStatistique (Groupe, CompteActivite, CompteCommission, Taux, " &
               "DateCreation, CreePar) " &
               "VALUES (@cle, @activite, @commission, @taux, GETDATE(), @auteur)",
               "UPDATE T_GroupeStatistique SET CompteActivite = @activite, " &
               "CompteCommission = @commission, Taux = @taux, DateModification = GETDATE(), " &
               "ModifiePar = @auteur WHERE Groupe = @cle")

        If Not ExecuterApplication(requete, demande, creation, "groupe",
                                   connexion, transaction, messageErreur) Then
            Return False
        End If

        ' Le groupe est la source de vérité ; les colonnes de T_Pdv_SA en sont le miroir, que la
        ' comptabilisation quotidienne continue de lire. Une modification autorisée les réaligne
        ' donc aussitôt, dans la même transaction — sans quoi la pièce comptable du lendemain
        ' utiliserait encore les anciens comptes, et personne ne verrait l'écart.
        '
        ' Une création n'a rien à réaligner : le groupe n'a pas encore de sous-agent.
        If creation Then Return True

        Return AppliquerSynchronisation(demande, connexion, transaction, messageErreur)
    End Function

    ''' <summary>
    ''' Reporte les valeurs d'un groupe sur tous les sous-agents qui le portent.
    '''
    ''' Aucune ligne touchée n'est une anomalie : le groupe n'a peut-être plus de sous-agent.
    ''' On laisse donc passer, contrairement aux autres opérations.
    ''' </summary>
    Private Shared Function AppliquerSynchronisation(demande As DemandeWU, connexion As SqlConnection,
                                                     transaction As SqlTransaction,
                                                     ByRef messageErreur As String) As Boolean

        Const requete As String =
            "UPDATE T_Pdv_SA SET CompteCompense = @activite, CompteCommission = @commission, " &
            "Taux = @taux, DateModification = GETDATE(), ModifiePar = @auteur " &
            "WHERE LTRIM(RTRIM(GroupeStatistique)) = @cle"

        Try
            Using commande As New SqlCommand(requete, connexion, transaction)
                AjouterParametresDemande(commande, demande)
                commande.ExecuteNonQuery()
            End Using

        Catch ex As SqlException
            messageErreur = $"Synchronisation des sous-agents du groupe « {demande.Cle} » " &
                            $"impossible : {ex.Message}"
            Return False
        End Try

        Return True
    End Function

    Private Shared Function AppliquerSuppression(table As String, colonneCle As String, cle As String,
                                                 objet As String, connexion As SqlConnection,
                                                 transaction As SqlTransaction,
                                                 ByRef messageErreur As String) As Boolean

        Try
            Using commande As New SqlCommand($"DELETE FROM {table} WHERE {colonneCle} = @cle",
                                             connexion, transaction)

                commande.Parameters.Add("@cle", SqlDbType.NVarChar, 255).Value = cle.Trim()

                If commande.ExecuteNonQuery() = 0 Then
                    messageErreur = $"Aucune suppression : le {objet} « {cle} » n'existe plus." &
                                    Environment.NewLine &
                                    "Il a peut-être été supprimé depuis un autre poste."
                    Return False
                End If
            End Using

        Catch ex As SqlException
            messageErreur = $"Suppression du {objet} « {cle} » impossible : {ex.Message}"
            Return False
        End Try

        Return True
    End Function

    Private Shared Function ExecuterApplication(requete As String, demande As DemandeWU,
                                                creation As Boolean, objet As String,
                                                connexion As SqlConnection, transaction As SqlTransaction,
                                                ByRef messageErreur As String) As Boolean
        Try
            Using commande As New SqlCommand(requete, connexion, transaction)

                AjouterParametresDemande(commande, demande)

                If commande.ExecuteNonQuery() = 0 Then
                    messageErreur = $"Aucune fiche modifiée : le {objet} « {demande.Cle} » " &
                                    "n'existe plus." & Environment.NewLine &
                                    "Il a peut-être été supprimé depuis un autre poste."
                    Return False
                End If
            End Using

        Catch ex As SqlException
            Dim action As String = If(creation, "créer", "modifier")
            messageErreur = MessageErreurEcriture(ex, demande.Cle, objet, action)
            Return False
        End Try

        Return True
    End Function

    ''' <summary>Paramètres communs à toutes les applications de demande.</summary>
    Private Shared Sub AjouterParametresDemande(commande As SqlCommand, demande As DemandeWU)

        commande.Parameters.Add("@cle", SqlDbType.NVarChar, 255).Value = demande.Cle.Trim()
        commande.Parameters.Add("@designation", SqlDbType.NVarChar, 255).Value = If(demande.Designation, String.Empty)
        commande.Parameters.Add("@groupe", SqlDbType.NVarChar, 255).Value = If(demande.GroupeStatistique, String.Empty)
        commande.Parameters.Add("@activite", SqlDbType.NVarChar, 255).Value = If(demande.CompteActivite, String.Empty)
        commande.Parameters.Add("@commission", SqlDbType.NVarChar, 255).Value = If(demande.CompteCommission, String.Empty)
        commande.Parameters.Add("@rattachement", SqlDbType.NVarChar, 255).Value = If(demande.CodeRattachement, String.Empty)

        ' DECIMAL(4,2) : précision et échelle cadrées, valeur arrondie côté application, pour
        ' qu'aucun arrondi ne se produise en silence dans SQL Server.
        Dim parametreTaux As SqlParameter = commande.Parameters.Add("@taux", SqlDbType.Decimal)
        parametreTaux.Precision = 4
        parametreTaux.Scale = 2
        parametreTaux.Value = Decimal.Round(demande.Taux, 2)

        commande.Parameters.Add("@auteur", SqlDbType.NVarChar, 50).Value = demande.SaisiPar
    End Sub

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

    Private Shared Function LireEntier(lecteur As SqlDataReader, colonne As String) As Integer
        Dim index As Integer = lecteur.GetOrdinal(colonne)
        If lecteur.IsDBNull(index) Then Return 0
        Return Convert.ToInt32(lecteur.GetValue(index), Globalization.CultureInfo.InvariantCulture)
    End Function

    Private Shared Function LireDecimal(lecteur As SqlDataReader, colonne As String) As Decimal
        Dim index As Integer = lecteur.GetOrdinal(colonne)
        If lecteur.IsDBNull(index) Then Return 0D
        Return WUReportService.ToDecimalSafe(lecteur.GetValue(index))
    End Function

#End Region

End Class
