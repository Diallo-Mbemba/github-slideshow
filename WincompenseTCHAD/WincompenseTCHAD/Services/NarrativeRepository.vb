Option Strict On
Option Explicit On

Imports System.Data
Imports System.Data.SqlClient

''' <summary>
''' LES LIBELLÉS PAR NATURE DE MOUVEMENT (table T_NarrativeNatureWU), et le paramétrage de
''' narrative en vigueur qu'ils composent avec le mode et le modèle global.
'''
''' POURQUOI UNE TABLE, ALORS QUE LE MODÈLE GLOBAL TIENT DANS T_ParametreWU
'''
''' Treize natures feraient treize clés dans une table conçue pour des options, et personne
''' ouvrant T_ParametreWU dans Management Studio ne verrait plus où sont les réglages de
''' procédure. Surtout : une table à elle permet de dire ce qu'une clé/valeur ne dit pas —
''' qui a saisi ce libellé-là, et quand.
'''
''' UNE NATURE ABSENTE DE LA TABLE N'EST PAS UNE NATURE SANS LIBELLÉ
'''
''' Elle retombe sur le modèle global (voir NarrativesWU.Modele). C'est ce qui rend la bascule
''' de mode SANS EFFET tant que la banque n'a rien saisi : une table vide, en mode par nature,
''' rend exactement la pièce d'avant.
'''
''' LE CACHE, ET POURQUOI IL EST NÉCESSAIRE
'''
''' VerifierEquilibrePiece a besoin du paramétrage pour le libellé de l'écart d'arrondi, et
''' elle est appelée sans contexte. Sans cache, l'export Excel d'un classeur de cinquante
''' feuilles rouvrirait la table cinquante fois. Oublier() le fait relire — après une
''' modification, ou quand la connexion change de base.
''' </summary>
Public NotInheritable Class NarrativeRepository

    Private Sub New()
    End Sub

    Private Const TABLE As String = "T_NarrativeNatureWU"

    ''' <summary>Code d'erreur SQL Server signalant une table absente (« Invalid object name »).</summary>
    Private Const ERREUR_TABLE_ABSENTE As Integer = 208

    Public Const MESSAGE_TABLE_ABSENTE As String =
        "La table T_NarrativeNatureWU n'existe pas encore dans la base." & vbCrLf & vbCrLf &
        "Exécutez le script Scripts\23_NarrativeParNature.sql : il la crée." & vbCrLf &
        "Tant qu'elle est absente, le mode « un libellé par nature » reste indisponible et " &
        "toutes les lignes portent le modèle global."

#Region "Cache"

    ''' <summary>
    ''' Les modèles déjà lus. Nothing tant qu'aucune lecture n'a eu lieu : la distinction
    ''' compte, un dictionnaire vide signifierait « lu, et aucune nature n'est personnalisée ».
    ''' </summary>
    Private Shared _modeles As Dictionary(Of NatureMouvementWU, String)

    ''' <summary>Fait relire les libellés par nature à la prochaine demande.</summary>
    Public Shared Sub Oublier()
        _modeles = Nothing
    End Sub

    Private Shared Function Modeles() As Dictionary(Of NatureMouvementWU, String)

        If _modeles IsNot Nothing Then Return _modeles

        Dim messageErreur As String = String.Empty
        Dim lus As Dictionary(Of NatureMouvementWU, String) = Lire(messageErreur)

        ' Une lecture qui a échoué n'est PAS mise en cache : la base peut redevenir joignable,
        ' et un paramétrage lu à vide une fois vaudrait pour toute la session.
        If messageErreur.Length > 0 Then Return lus

        _modeles = lus
        Return _modeles
    End Function

#End Region

#Region "Lecture"

    ''' <summary>
    ''' Le paramétrage en vigueur : le mode et le modèle global viennent de T_ParametreWU, les
    ''' treize modèles de cette table.
    '''
    ''' C'EST LUI QU'UNE PIÈCE LIT, UNE FOIS, À SA GÉNÉRATION. L'objet rendu est immuable :
    ''' une pièce ne mélange jamais deux paramétrages parce que quelqu'un a enregistré un
    ''' libellé pendant qu'elle se construisait.
    ''' </summary>
    Public Shared Function EnVigueur() As NarrativesWU

        Return New NarrativesWU(OptionsWU.NarrativeMode, OptionsWU.NarrativeModele, Modeles())
    End Function

    ''' <summary>
    ''' Les modèles saisis, nature par nature. Les natures absentes de la table n'y figurent
    ''' pas : c'est NarrativesWU qui décide de leur repli, et lui seul.
    ''' </summary>
    Public Shared Function Lire(ByRef messageErreur As String) As Dictionary(Of NatureMouvementWU, String)

        messageErreur = String.Empty

        Dim lus As New Dictionary(Of NatureMouvementWU, String)

        Const lecture As String = "SELECT Nature, Modele FROM " & TABLE

        Try
            Using connexion As SqlConnection = WURepository.CreerConnexion()
                connexion.Open()

                Using commande As New SqlCommand(lecture, connexion)
                    Using lecteur As SqlDataReader = commande.ExecuteReader()
                        While lecteur.Read()

                            Dim code As String = If(lecteur.IsDBNull(0), String.Empty,
                                                    Convert.ToString(lecteur.GetValue(0)).Trim())

                            ' Un code que le code applicatif ne connaît pas — une nature
                            ' retirée, ou une ligne saisie à la main — est IGNORÉ : la pièce
                            ' doit sortir, et elle sort avec le modèle global pour cette ligne.
                            Dim nature As NatureMouvementWU? = NaturesMouvementWU.DepuisCode(code)
                            If Not nature.HasValue Then Continue While

                            Dim modele As String = If(lecteur.IsDBNull(1), String.Empty,
                                                      Convert.ToString(lecteur.GetValue(1)).Trim())
                            If modele.Length = 0 Then Continue While

                            lus(nature.Value) = modele
                        End While
                    End Using
                End Using
            End Using

        Catch ex As SqlException
            messageErreur = If(ex.Number = ERREUR_TABLE_ABSENTE,
                               MESSAGE_TABLE_ABSENTE,
                               $"Lecture des libellés par nature impossible : {ex.Message}")

        Catch ex As InvalidOperationException
            messageErreur = $"Connexion SQL Server indisponible : {ex.Message}"
        End Try

        Return lus
    End Function

    ''' <summary>
    ''' Vrai si la table existe : c'est la condition du mode « un libellé par nature ».
    ''' L'écran s'en sert pour dire à la banque quel script exécuter, au lieu de lui proposer
    ''' un mode qui ne pourrait rien retenir.
    ''' </summary>
    Public Shared Function Disponible() As Boolean

        Dim messageErreur As String = String.Empty
        Lire(messageErreur)
        Return messageErreur.Length = 0
    End Function

#End Region

#Region "Écriture"

    ''' <summary>
    ''' Enregistre le modèle d'une nature. Réservé à l'administrateur, comme tout le
    ''' paramétrage de narrative.
    '''
    ''' UN MODÈLE VIDE EFFACE LA LIGNE, et c'est la façon de dire « cette nature reprend le
    ''' modèle global ». Garder une ligne vide donnerait une nature personnalisée à chaîne
    ''' vide, c'est-à-dire douze écritures sans libellé le jour où le mode bascule.
    '''
    ''' LA TRACE PART AVEC LE CHANGEMENT, dans la même transaction que lui : voir
    ''' JournalParametreWU, et la même raison qu'ailleurs — un changement sans trace laisse le
    ''' grand livre inexplicable, une trace sans changement accuse quelqu'un à tort.
    ''' </summary>
    Public Shared Function Enregistrer(nature As NatureMouvementWU, modele As String,
                                       ByRef messageErreur As String) As Boolean

        messageErreur = String.Empty

        If Not SessionWU.PeutGererLesComptesSystemes Then
            messageErreur = "Seul l'administrateur peut modifier les libellés de la pièce comptable."
            Return False
        End If

        Dim code As String = NaturesMouvementWU.Code(nature)
        Dim valeur As String = If(modele, String.Empty).Trim()

        Const ecriture As String =
            "UPDATE " & TABLE & " SET Modele = @modele, DateModification = GETDATE(), " &
            "ModifiePar = @auteur WHERE Nature = @nature; " &
            "IF @@ROWCOUNT = 0 " &
            "INSERT INTO " & TABLE & " (Nature, Modele, DateModification, ModifiePar) " &
            "VALUES (@nature, @modele, GETDATE(), @auteur);"

        Const effacement As String = "DELETE FROM " & TABLE & " WHERE Nature = @nature;"

        Dim journalDisponible As Boolean = JournalParametreWU.Disponible()

        Try
            Using connexion As SqlConnection = WURepository.CreerConnexion()
                connexion.Open()

                Using transaction As SqlTransaction = connexion.BeginTransaction()

                    Dim ancienne As String = ModeleVerrouille(connexion, transaction, code)

                    Dim aJournaliser As Boolean =
                        journalDisponible AndAlso
                        Not String.Equals(ancienne.Trim(), valeur, StringComparison.Ordinal)

                    Dim sql As String = If(valeur.Length = 0, effacement, ecriture)
                    If aJournaliser Then sql &= " " & JournalParametreWU.INSERTION

                    Using commande As New SqlCommand(sql, connexion, transaction)

                        commande.Parameters.Add("@nature", SqlDbType.NVarChar, 40).Value = code

                        ' Le DELETE ne porte pas ces deux paramètres : les ajouter quand même
                        ' serait inoffensif, les omettre quand il le faut ne l'est pas.
                        If valeur.Length > 0 Then
                            commande.Parameters.Add("@modele", SqlDbType.NVarChar, 255).Value = valeur
                            commande.Parameters.Add("@auteur", SqlDbType.NVarChar, 50).Value = SessionWU.Auteur
                        End If

                        If aJournaliser Then
                            JournalParametreWU.AjouterLesParametres(
                                commande, NaturesMouvementWU.CleDeJournal(nature), ancienne, valeur)
                        End If

                        commande.ExecuteNonQuery()
                    End Using

                    transaction.Commit()
                End Using
            End Using

        Catch ex As SqlException
            messageErreur = If(ex.Number = ERREUR_TABLE_ABSENTE,
                               MESSAGE_TABLE_ABSENTE,
                               $"Enregistrement du libellé de « {NaturesMouvementWU.Intitule(nature)} » " &
                               $"impossible : {ex.Message}")
            Return False

        Catch ex As InvalidOperationException
            messageErreur = $"Connexion SQL Server indisponible : {ex.Message}"
            Return False
        End Try

        Oublier()
        Return True
    End Function

    ''' <summary>
    ''' Le modèle actuel d'une nature, lu DANS la transaction de l'appelant et la ligne
    ''' verrouillée jusqu'à son terme. Chaîne vide si la nature n'est pas personnalisée.
    ''' </summary>
    Private Shared Function ModeleVerrouille(connexion As SqlConnection, transaction As SqlTransaction,
                                              code As String) As String

        Const lecture As String =
            "SELECT Modele FROM " & TABLE & " WITH (UPDLOCK, HOLDLOCK) WHERE Nature = @nature"

        Using commande As New SqlCommand(lecture, connexion, transaction)

            commande.Parameters.Add("@nature", SqlDbType.NVarChar, 40).Value = code

            Dim valeur As Object = commande.ExecuteScalar()
            If valeur Is Nothing OrElse valeur Is DBNull.Value Then Return String.Empty

            Return Convert.ToString(valeur)
        End Using
    End Function

#End Region

End Class
