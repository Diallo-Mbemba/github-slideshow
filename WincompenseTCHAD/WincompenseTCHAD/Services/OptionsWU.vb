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

    ''' <summary>
    ''' Les envois non encore réglés (TxnStatus = "W") entrent-ils dans le calcul des écarts
    ''' de change ?
    '''
    ''' PAR DÉFAUT OUI, ET CETTE OPTION EST LA SEULE DU PROJET DONT L'ABSENCE VAUT « OUI ».
    ''' Les autres options commandent un BLOCAGE, et une absence ne doit jamais bloquer. Celle-ci
    ''' commande une EXCLUSION de lignes : une absence qui vaudrait « non » ferait disparaître
    ''' silencieusement 290 transactions et la moitié du gain de change. Le défaut le moins
    ''' dangereux n'est donc pas le même, et c'est pourquoi la lecture passe par EstNon et non
    ''' par EstOui.
    ''' </summary>
    Public Const CLE_CHANGE_ENVOIS_EN_ATTENTE As String = "CHANGE_ENVOIS_EN_ATTENTE"

    Public Const LIBELLE_CHANGE_ENVOIS_EN_ATTENTE As String =
        "OUI : les envois encore en attente de règlement (statut W) entrent dans le calcul des " &
        "écarts de change. NON : ils en sont écartés et leur change sera constaté plus tard."

    ''' <summary>
    ''' LE MODÈLE DE NARRATIVE : la phrase que porte chaque ligne de la pièce comptable, et
    ''' que le core banking reçoit dans sa colonne ADDLTEXT.
    '''
    ''' CE N'EST PAS UNE OPTION OUI/NON, et c'est la première de cette table qui ne l'est pas.
    ''' La valeur est le GABARIT lui-même, avec ses repères — « LD WU ACTIVITE {AGENCE}
    ''' {PERIODE} ». Elle est lue par ModeleNarrativeWU, qui seul sait la rendre ; cette classe
    ''' ne fait que l'apporter.
    '''
    ''' ELLE VIT ICI PLUTÔT QUE DANS UNE TABLE À ELLE, parce que T_ParametreWU est faite pour
    ''' ça : une clé, une valeur, son libellé, qui l'a changée et quand. Une table de plus
    ''' aurait demandé un script, un référentiel, des droits et un dépôt — pour une ligne.
    ''' </summary>
    Public Const CLE_NARRATIVE_MODELE As String = "NARRATIVE_MODELE"

    Public Const LIBELLE_NARRATIVE_MODELE As String =
        "Modèle de la narrative des lignes de la pièce comptable et de la colonne ADDLTEXT du " &
        "fichier core banking. Repères reconnus : {AGENCE}, {ACCOUNT}, {CODE_AGENCE}, {PERIODE}."

    ''' <summary>
    ''' COMMENT LES LIBELLÉS SONT CHOISIS : un seul modèle pour les douze lignes d'un point de
    ''' vente (GLOBAL), ou un modèle par nature de mouvement (PAR_NATURE).
    '''
    ''' GLOBAL PAR DÉFAUT, et par défaut en cas de doute : une valeur mal orthographiée, une
    ''' table absente, une base injoignable laissent le mode actuel. Basculer treize libellés
    ''' d'un coup sur un malentendu de lecture n'est pas une option.
    ''' </summary>
    Public Const CLE_NARRATIVE_MODE As String = "NARRATIVE_MODE"

    Public Const LIBELLE_NARRATIVE_MODE As String =
        "GLOBAL : les douze lignes d'un point de vente portent le même libellé. PAR_NATURE : " &
        "chaque ligne porte le libellé de sa nature de mouvement (table T_NarrativeNatureWU)."

    ''' <summary>Code d'erreur SQL Server signalant une table absente (« Invalid object name »).</summary>
    Private Const ERREUR_TABLE_ABSENTE As Integer = 208

    Public Const MESSAGE_TABLE_ABSENTE As String =
        "La table T_ParametreWU n'existe pas encore dans la base." & vbCrLf & vbCrLf &
        "Exécutez le script Scripts\00_InstallationComplete.sql : il la crée." & vbCrLf &
        "C'est le SEUL script à exécuter — il contient tous les autres — et il peut être " &
        "rejoué sans risque : il ne crée que ce qui manque." & vbCrLf &
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

    ''' <summary>
    ''' Vrai si les envois en attente de règlement entrent dans le calcul des écarts de change.
    ''' VRAI par défaut, et vrai en cas de doute : voir CLE_CHANGE_ENVOIS_EN_ATTENTE.
    ''' </summary>
    Public Shared ReadOnly Property ChangeInclureEnvoisEnAttente As Boolean
        Get
            Return Not EstNon(Lire(CLE_CHANGE_ENVOIS_EN_ATTENTE))
        End Get
    End Property

    ''' <summary>
    ''' Le modèle de narrative en vigueur, ou celui du code si la banque n'en a saisi aucun.
    '''
    ''' UNE ABSENCE REND LE COMPORTEMENT ACTUEL, et non une phrase vide : table absente, clé
    ''' jamais créée, base injoignable, valeur effacée à la main — dans tous ces cas la pièce
    ''' sort avec les narratives d'aujourd'hui. C'est la règle de ce fichier, et elle compte
    ''' doublement ici : une narrative vide ne bloquerait pas la compense, elle produirait
    ''' douze écritures sans libellé, qui partiraient au grand livre sans que rien n'avertisse.
    ''' </summary>
    Public Shared ReadOnly Property NarrativeModele As String
        Get
            Dim saisi As String = Lire(CLE_NARRATIVE_MODELE)
            If String.IsNullOrWhiteSpace(saisi) Then Return ModeleNarrativeWU.ModeleParDefaut
            Return saisi
        End Get
    End Property

    ''' <summary>
    ''' Le mode de narrative en vigueur. MODE UNIQUE par défaut, et en cas de doute : voir
    ''' CLE_NARRATIVE_MODE.
    ''' </summary>
    Public Shared ReadOnly Property NarrativeMode As ModeNarrativeWU
        Get
            Return NarrativesWU.ModeDepuisCode(Lire(CLE_NARRATIVE_MODE))
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

    ''' <summary>
    ''' Interprète une valeur en NON franc. Tout le reste — vide, inconnu, mal orthographié —
    ''' n'est pas un refus.
    '''
    ''' CE N'EST PAS LA NÉGATION D'EstOui, et la différence est le but. EstOui range l'inconnu
    ''' du côté du NON ; EstNon le range du côté du OUI. Chaque option choisit alors celle des
    ''' deux lectures dont le défaut est le moins dangereux pour elle : refuser de bloquer, ou
    ''' refuser d'exclure.
    ''' </summary>
    Public Shared Function EstNon(valeur As String) As Boolean

        Select Case If(valeur, String.Empty).Trim().ToUpperInvariant()
            Case "NON", "N", "FAUX", "FALSE", "0" : Return True
            Case Else : Return False
        End Select
    End Function

    ''' <summary>La valeur telle qu'elle s'écrit en base.</summary>
    Public Shared Function Texte(actif As Boolean) As String
        Return If(actif, "OUI", "NON")
    End Function

    ''' <summary>
    ''' Qui a changé cette option, et quand — tel que la base le porte, en une phrase prête à
    ''' afficher. Chaîne vide si la clé n'a jamais été enregistrée, ou si la lecture échoue.
    '''
    ''' POURQUOI C'EST LU ET NON DÉDUIT. L'écran pourrait annoncer « modifié par vous, à
    ''' l'instant » après un enregistrement, et c'est ce que fait l'écran des options. Mais le
    ''' modèle de narrative part dans le grand livre de la banque : celui qui l'ouvre doit
    ''' pouvoir constater si quelqu'un d'autre l'a changé avant lui, et non le supposer.
    '''
    ''' CETTE LECTURE NE PASSE PAS PAR LE CACHE : le cache ne retient que les valeurs, et ce
    ''' sont justement les deux colonnes qui n'y sont pas.
    ''' </summary>
    Public Shared Function DerniereModification(cle As String) As String

        If String.IsNullOrWhiteSpace(cle) Then Return String.Empty

        Const lecture As String =
            "SELECT ModifiePar, DateModification FROM " & TABLE & " WHERE Cle = @cle"

        Try
            Using connexion As SqlConnection = WURepository.CreerConnexion()
                connexion.Open()

                Using commande As New SqlCommand(lecture, connexion)
                    commande.Parameters.Add("@cle", SqlDbType.NVarChar, 50).Value = cle.Trim()

                    Using lecteur As SqlDataReader = commande.ExecuteReader()

                        If Not lecteur.Read() Then Return String.Empty

                        Dim auteur As String = If(lecteur.IsDBNull(0), String.Empty,
                                                  Convert.ToString(lecteur.GetValue(0)).Trim())

                        If lecteur.IsDBNull(1) Then
                            If auteur.Length = 0 Then Return String.Empty
                            Return $"Enregistré par {auteur}."
                        End If

                        Dim quand As Date = lecteur.GetDateTime(1)

                        If auteur.Length = 0 Then Return $"Enregistré le {quand:dd/MM/yyyy à HH:mm}."
                        Return $"Enregistré le {quand:dd/MM/yyyy à HH:mm} par {auteur}."
                    End Using
                End Using
            End Using

        Catch ex As SqlException
            Return String.Empty

        Catch ex As InvalidOperationException
            Return String.Empty
        End Try
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

        ' LA QUESTION EST POSÉE AVANT D'OUVRIR LA TRANSACTION : Disponible() ouvre sa propre
        ' connexion, et l'appeler transaction ouverte ferait tenir deux connexions à la fois
        ' pour une réponse mise en cache de toute façon.
        Dim journalDisponible As Boolean = JournalParametreWU.Disponible()

        Try
            Using connexion As SqlConnection = WURepository.CreerConnexion()
                connexion.Open()

                ' UNE SEULE TRANSACTION POUR LES DEUX ÉCRITURES. Le paramètre et sa trace
                ' partent ensemble ou ne partent pas : un changement sans trace laisserait le
                ' grand livre inexplicable, et une trace sans changement serait pire — elle
                ' accuserait quelqu'un d'une modification qui n'a pas eu lieu.
                Using transaction As SqlTransaction = connexion.BeginTransaction()

                    ' L'ancienne valeur est lue DANS la transaction, et la ligne est verrouillée
                    ' en lecture : sans cela, deux administrateurs enregistrant en même temps
                    ' journaliseraient tous deux la même « ancienne » valeur, et l'un des deux
                    ' changements n'aurait aucune trace de son point de départ.
                    Dim ancienne As String = ValeurVerrouillee(connexion, transaction, cle.Trim())

                    ' JOURNALISÉ SEULEMENT SI LA VALEUR CHANGE RÉELLEMENT. Réenregistrer le même
                    ' modèle est un geste d'écran — on ouvre, on regarde, on valide — et un
                    ' journal rempli de lignes identiques ne se lit plus.
                    Dim aJournaliser As Boolean =
                        journalDisponible AndAlso
                        Not String.Equals(ancienne.Trim(), If(valeur, String.Empty).Trim(),
                                          StringComparison.Ordinal)

                    Dim sql As String = requete
                    If aJournaliser Then sql &= " " & JournalParametreWU.INSERTION

                    Using commande As New SqlCommand(sql, connexion, transaction)

                        commande.Parameters.Add("@cle", SqlDbType.NVarChar, 50).Value = cle.Trim()
                        commande.Parameters.Add("@valeur", SqlDbType.NVarChar, 255).Value = If(valeur, String.Empty)
                        commande.Parameters.Add("@libelle", SqlDbType.NVarChar, 255).Value = If(libelle, String.Empty)
                        commande.Parameters.Add("@auteur", SqlDbType.NVarChar, 50).Value = SessionWU.Auteur

                        If aJournaliser Then
                            JournalParametreWU.AjouterLesParametres(commande, cle, ancienne, valeur)
                        End If

                        commande.ExecuteNonQuery()
                    End Using

                    transaction.Commit()
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

    ''' <summary>
    ''' La valeur actuelle d'une clé, lue DANS la transaction de l'appelant et la ligne
    ''' verrouillée jusqu'à son terme. Chaîne vide si la clé n'existe pas encore — ce qui est
    ''' une information, et non une absence de réponse : le journal l'écrira NULL.
    ''' </summary>
    Private Shared Function ValeurVerrouillee(connexion As SqlConnection, transaction As SqlTransaction,
                                              cle As String) As String

        Const lecture As String =
            "SELECT Valeur FROM " & TABLE & " WITH (UPDLOCK, HOLDLOCK) WHERE Cle = @cle"

        Using commande As New SqlCommand(lecture, connexion, transaction)

            commande.Parameters.Add("@cle", SqlDbType.NVarChar, 50).Value = cle

            Dim valeur As Object = commande.ExecuteScalar()
            If valeur Is Nothing OrElse valeur Is DBNull.Value Then Return String.Empty

            Return Convert.ToString(valeur)
        End Using
    End Function

#End Region

End Class
