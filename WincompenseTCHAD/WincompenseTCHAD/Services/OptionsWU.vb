Option Strict On
Option Explicit On

Imports System.Data
Imports System.Data.SqlClient

''' <summary>
''' Les options de traitement de la banque (table T_ParametreWU).
'''
''' POURQUOI EN BASE, ET NON SUR LE POSTE
'''
''' Ces options décrivent la façon de travailler de la banque, et non la configuration d'un
''' poste. Si le visa est obligatoire avant de produire le fichier core banking, il l'est pour
''' tout le monde — un agent ne doit pas pouvoir s'en affranchir en décochant une case chez lui.
'''
''' UNE ABSENCE N'EST JAMAIS UNE INTERDICTION
'''
''' Table absente, base injoignable, option jamais créée : la valeur par défaut s'applique, et
''' elle est toujours la moins bloquante. Une règle de contrôle qui s'activerait toute seule
''' parce qu'une lecture a échoué arrêterait la compense du jour pour une raison que personne
''' ne comprendrait.
'''
''' LA VALEUR EST RELUE, PAS DEVINÉE
'''
''' Elle est mise en cache le temps de la session pour ne pas interroger la base à chaque clic,
''' et Oublier() la fait relire — après une modification, ou quand la connexion change.
''' </summary>
Public NotInheritable Class OptionsWU

    Private Sub New()
    End Sub

    Private Const TABLE As String = "T_ParametreWU"

    ''' <summary>Le visa d'une journée bloque-t-il la production du fichier core banking ?</summary>
    Public Const CLE_VISA_AVANT_CORE_BANKING As String = "VISA_AVANT_CORE_BANKING"

    Public Const LIBELLE_VISA_AVANT_CORE_BANKING As String =
        "OUI : le fichier core banking ne peut pas être produit tant que la journée n'est pas " &
        "visée. NON : l'application avertit seulement."

    ''' <summary>Code d'erreur SQL Server signalant une table absente (« Invalid object name »).</summary>
    Private Const ERREUR_TABLE_ABSENTE As Integer = 208

    Public Const MESSAGE_TABLE_ABSENTE As String =
        "La table T_ParametreWU n'existe pas encore dans la base." & vbCrLf & vbCrLf &
        "Exécutez le script Scripts\18_OptionsTraitement.sql : il la crée." & vbCrLf &
        "Tant qu'elle est absente, l'application se comporte comme si aucune option n'était " &
        "activée — elle avertit, elle ne bloque pas."

#Region "Cache"

    ''' <summary>
    ''' Valeurs déjà lues. Nothing tant qu'aucune lecture n'a eu lieu : la distinction compte,
    ''' un dictionnaire vide signifierait « lu, et il n'y a rien ».
    ''' </summary>
    Private Shared _valeurs As Dictionary(Of String, String)

    ''' <summary>
    ''' Fait relire les options à la prochaine demande. À appeler après une modification, et
    ''' quand la connexion change de base — les options de l'ancienne ne valent pas pour la
    ''' nouvelle.
    ''' </summary>
    Public Shared Sub Oublier()
        _valeurs = Nothing
    End Sub

    Private Shared Function Valeurs() As Dictionary(Of String, String)

        If _valeurs IsNot Nothing Then Return _valeurs

        Dim lues As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

        Const lecture As String = "SELECT Cle, Valeur FROM " & TABLE

        Try
            Using connexion As SqlConnection = WURepository.CreerConnexion()
                connexion.Open()

                Using commande As New SqlCommand(lecture, connexion)
                    Using lecteur As SqlDataReader = commande.ExecuteReader()
                        While lecteur.Read()

                            Dim cle As String = Convert.ToString(lecteur.GetValue(0)).Trim()
                            If cle.Length = 0 Then Continue While

                            lues(cle) = If(lecteur.IsDBNull(1), String.Empty,
                                           Convert.ToString(lecteur.GetValue(1)).Trim())
                        End While
                    End Using
                End Using
            End Using

        Catch ex As SqlException
            ' Table absente ou base injoignable : les valeurs par défaut s'appliquent, et
            ' elles ne bloquent rien. On ne met PAS ce résultat en cache : la base peut
            ' redevenir joignable, et une option lue une fois à vide vaudrait pour la session.
            Return lues

        Catch ex As InvalidOperationException
            Return lues
        End Try

        _valeurs = lues
        Return _valeurs
    End Function

#End Region

#Region "Lecture"

    ''' <summary>
    ''' Vrai si la production du fichier core banking exige que la journée soit visée.
    ''' Faux par défaut, et faux en cas de doute.
    ''' </summary>
    Public Shared ReadOnly Property VisaAvantCoreBanking As Boolean
        Get
            Return EstOui(Lire(CLE_VISA_AVANT_CORE_BANKING))
        End Get
    End Property

    ''' <summary>La valeur brute d'une option, ou une chaîne vide.</summary>
    Public Shared Function Lire(cle As String) As String

        Dim lues As Dictionary(Of String, String) = Valeurs()
        If Not lues.ContainsKey(cle) Then Return String.Empty
        Return lues(cle)
    End Function

    ''' <summary>
    ''' Interprète une valeur en oui / non. Tout ce qui n'est pas franchement affirmatif vaut
    ''' NON : une option de blocage ne s'active pas sur un malentendu.
    ''' </summary>
    Public Shared Function EstOui(valeur As String) As Boolean

        Select Case If(valeur, String.Empty).Trim().ToUpperInvariant()
            Case "OUI", "O", "VRAI", "TRUE", "1" : Return True
            Case Else : Return False
        End Select
    End Function

    ''' <summary>La valeur telle qu'elle s'écrit en base.</summary>
    Public Shared Function Texte(actif As Boolean) As String
        Return If(actif, "OUI", "NON")
    End Function

#End Region

#Region "Écriture"

    ''' <summary>
    ''' Enregistre une option. Réservé à l'administrateur : une règle de procédure ne se change
    ''' pas depuis le guichet.
    '''
    ''' L'écriture crée la ligne si elle manque : une base où le script a été joué mais l'option
    ''' supprimée à la main doit pouvoir être remise d'aplomb depuis l'application.
    ''' </summary>
    Public Shared Function Enregistrer(cle As String, valeur As String, libelle As String,
                                       ByRef messageErreur As String) As Boolean

        messageErreur = String.Empty

        If Not SessionWU.PeutGererLesComptesSystemes Then
            messageErreur = "Seul l'administrateur peut modifier les options de traitement."
            Return False
        End If

        If String.IsNullOrWhiteSpace(cle) Then
            messageErreur = "Aucune option à enregistrer."
            Return False
        End If

        Const requete As String =
            "UPDATE " & TABLE & " SET Valeur = @valeur, Libelle = @libelle, " &
            "DateModification = GETDATE(), ModifiePar = @auteur WHERE Cle = @cle; " &
            "IF @@ROWCOUNT = 0 " &
            "INSERT INTO " & TABLE & " (Cle, Valeur, Libelle, DateModification, ModifiePar) " &
            "VALUES (@cle, @valeur, @libelle, GETDATE(), @auteur);"

        Try
            Using connexion As SqlConnection = WURepository.CreerConnexion()
                connexion.Open()

                Using commande As New SqlCommand(requete, connexion)

                    commande.Parameters.Add("@cle", SqlDbType.NVarChar, 50).Value = cle.Trim()
                    commande.Parameters.Add("@valeur", SqlDbType.NVarChar, 255).Value = If(valeur, String.Empty)
                    commande.Parameters.Add("@libelle", SqlDbType.NVarChar, 255).Value = If(libelle, String.Empty)
                    commande.Parameters.Add("@auteur", SqlDbType.NVarChar, 50).Value = SessionWU.Auteur

                    commande.ExecuteNonQuery()
                End Using
            End Using

        Catch ex As SqlException
            messageErreur = If(ex.Number = ERREUR_TABLE_ABSENTE,
                               MESSAGE_TABLE_ABSENTE,
                               $"Enregistrement de l'option impossible : {ex.Message}")
            Return False

        Catch ex As InvalidOperationException
            messageErreur = $"Connexion SQL Server indisponible : {ex.Message}"
            Return False
        End Try

        Oublier()
        Return True
    End Function

#End Region

End Class
