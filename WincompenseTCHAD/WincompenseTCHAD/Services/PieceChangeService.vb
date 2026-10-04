Option Strict On
Option Explicit On

Imports System.Data

''' <summary>
''' Le découpage des pièces de change : combien de pièces pour un rapport.
'''
''' POURQUOI C'EST UN RÉGLAGE ET NON UNE CONSTANTE. La spécification demande une pièce par
''' journée de règlement, par sens et par code produit. Mesuré sur le rapport de référence —
''' trois journées — cela donne TRENTE-SEPT PIÈCES, dont une vingtaine portant moins de vingt
''' francs : une pièce comptable de 2 F, une autre de 1 F. Par journée seule, il en reste
''' trois. L'information est la même dans les trois cas, les totaux aussi ; seul le nombre de
''' pièces à passer change. Ce n'est pas au développeur de décider combien de pièces la
''' Direction Comptable veut saisir, et ce n'était pas une raison pour ne rien livrer.
''' </summary>
Public Enum DecoupageChangeWU

    ''' <summary>Une pièce par journée de règlement. Trois pièces sur le rapport de référence.</summary>
    ParJournee = 1

    ''' <summary>Une pièce par journée et par sens (envoi, paiement). Six pièces.</summary>
    ParJourneeEtSens = 2

    ''' <summary>
    ''' Une pièce par journée, par sens et par code produit : le découpage de la
    ''' spécification. Trente-sept pièces sur le rapport de référence.
    ''' </summary>
    ParJourneeSensEtProduit = 3

End Enum

''' <summary>
''' Une pièce comptable de change : ses lignes, et de quoi elles rendent compte.
'''
''' SES LIGNES SONT UN DataTable DE MÊME FORME QUE LA PIÈCE PRINCIPALE — Compte, Libelle,
''' Debit, Credit, CodeAgence. Ce n'est pas une imitation : c'est ce qui fait que l'export
''' Excel, le fichier core banking et le contrôle d'équilibre déjà écrits la traitent sans
''' une ligne de plus. Une forme propre aurait exigé de réécrire les trois.
''' </summary>
Public Class PieceChangeWU

    ''' <summary>
    ''' Clé du groupe, telle qu'elle identifie la pièce en base : « 2026-03-27 »,
    ''' « 2026-03-27|S » ou « 2026-03-27|S|IMTR » selon le découpage. Composée une fois ici,
    ''' elle reste unique dans les trois cas — voir T_PieceChangeWU.CleGroupe.
    ''' </summary>
    Public Property CleGroupe As String = String.Empty

    Public Property DateReglement As Date

    ''' <summary>"S", "P", ou vide si le découpage ne distingue pas le sens.</summary>
    Public Property Sens As String = String.Empty

    ''' <summary>Code produit, ou vide si le découpage ne le distingue pas.</summary>
    Public Property CodeProduit As String = String.Empty

    Public Property NombreTransactions As Integer
    Public Property Gains As Decimal
    Public Property Pertes As Decimal
    Public Property Parite As Decimal
    Public Property FichierSource As String = String.Empty

    ''' <summary>Le libellé porté par toutes les lignes de la pièce.</summary>
    Public Property Libelle As String = String.Empty

    ''' <summary>Les écritures. Voir la remarque sur la forme, en tête de cette classe.</summary>
    Public Property Lignes As DataTable

    Public ReadOnly Property Net As Decimal
        Get
            Return Gains - Pertes
        End Get
    End Property

    ''' <summary>
    ''' Total des débits de la pièce. Il égale toujours le total des crédits : chacune des deux
    ''' écritures de gain et chacune des deux écritures de perte est posée par paire.
    ''' </summary>
    Public ReadOnly Property TotalDebit As Long
        Get
            Return Totaliser("Debit")
        End Get
    End Property

    Public ReadOnly Property TotalCredit As Long
        Get
            Return Totaliser("Credit")
        End Get
    End Property

    ''' <summary>
    ''' Vrai si débits et crédits s'égalent. C'est une tautologie vu la façon dont la pièce est
    ''' construite — et c'est précisément pour cela qu'on le vérifie : le jour où une écriture
    ''' serait ajoutée sans sa contrepartie, le contrôle tombe au lieu de passer inaperçu.
    ''' </summary>
    Public ReadOnly Property EstEquilibree As Boolean
        Get
            Return TotalDebit = TotalCredit
        End Get
    End Property

    Private Function Totaliser(colonne As String) As Long

        If Lignes Is Nothing OrElse Not Lignes.Columns.Contains(colonne) Then Return 0L

        Dim total As Long = 0L
        For Each ligne As DataRow In Lignes.Rows
            total += Convert.ToInt64(ligne(colonne))
        Next
        Return total
    End Function

End Class

''' <summary>
''' Fabrique les pièces comptables de change à partir du résultat du calcul.
'''
''' PUR, COMME ChangeService, ET POUR LA MÊME RAISON. Il reçoit un résultat de calcul, un
''' paramétrage de comptes et un découpage ; il rend des pièces. Aucune base, aucun fichier,
''' aucun affichage. L'écran de contrôle peut donc produire exactement les pièces que la
''' comptabilisation produira, et le prouver avant que quoi que ce soit ne soit enregistré.
'''
''' LES ÉCRITURES, EN BRUT ET SANS COMPENSATION
'''
'''     Débit  compte de liaison       = total des gains
'''     Crédit compte de gain          = total des gains
'''     Débit  compte de perte         = total des pertes
'''     Crédit compte de liaison       = total des pertes
'''
''' LE COMPTE DE LIAISON APPARAÎT DONC DEUX FOIS, une fois au débit et une fois au crédit, et
''' ce n'est pas une maladresse : c'est ce que « présentation brute » veut dire. Compenser les
''' deux aurait donné une pièce plus courte et un gain net impossible à rapprocher du rapport,
''' où gains et pertes sont deux populations distinctes de transactions.
'''
''' AUCUNE LIGNE À ZÉRO n'est posée : une journée sans perte donne une pièce de deux lignes.
'''
''' CE SERVICE REFUSE DE PRODUIRE SANS SES COMPTES, et il le dit. C'est la règle posée par la
''' spécification, et c'est la seule tenable : inventer un numéro de compte pour une pièce qui
''' porte un demi-million de francs par semaine reviendrait à le faire comptabiliser.
''' </summary>
Public NotInheritable Class PieceChangeService

    Private Sub New()
    End Sub

#Region "Libellé"

    ''' <summary>
    ''' La mention qui identifie la pièce de change dans le journal, et dans la colonne
    ''' ADDLTEXT du fichier core banking.
    ''' </summary>
    Public Const MENTION As String = "écart de change"

    ''' <summary>
    ''' Compose le libellé de la pièce : « WU IMTR Envoi - écart de change - 27/03/2026 - 390 trx ».
    '''
    ''' Les parties que le découpage ne distingue pas disparaissent plutôt que de laisser un
    ''' trou : par journée seule, le libellé devient « WU - écart de change - 27/03/2026 -
    ''' 1 234 trx ». Le pire cas mesure 58 caractères, loin des 150 que le core banking
    ''' accepte dans ADDLTEXT.
    ''' </summary>
    Public Shared Function Narratif(sens As String, codeProduit As String,
                                    jour As Date, nombreTransactions As Integer) As String

        Dim morceaux As New List(Of String)
        morceaux.Add(ProduitTransfert.CODE_WESTERN_UNION)

        If Not String.IsNullOrWhiteSpace(codeProduit) Then morceaux.Add(codeProduit.Trim())

        Dim sensLisible As String = LibelleDuSens(sens)
        If sensLisible.Length > 0 Then morceaux.Add(sensLisible)

        Return $"{String.Join(" ", morceaux)} - {MENTION} - {jour:dd/MM/yyyy} - {nombreTransactions} trx"
    End Function

    Private Shared Function LibelleDuSens(sens As String) As String

        Select Case If(sens, String.Empty).Trim().ToUpperInvariant()
            Case ConstantesWU.SENS_ENVOI : Return "Envoi"
            Case ConstantesWU.SENS_PAIEMENT : Return "Paiement"
            Case Else : Return String.Empty
        End Select
    End Function

#End Region

#Region "Génération"

    ''' <summary>
    ''' Construit les pièces de change d'un résultat de calcul.
    ''' </summary>
    ''' <param name="resultat">Le résultat rendu par ChangeService.</param>
    ''' <param name="comptes">Le paramétrage en service. Nothing prend celui chargé au démarrage.</param>
    ''' <param name="decoupage">Combien de pièces. Voir <see cref="DecoupageChangeWU"/>.</param>
    ''' <param name="messageErreur">Le motif du refus, le cas échéant.</param>
    ''' <returns>Les pièces, triées, ou Nothing si la génération est refusée.</returns>
    Public Shared Function Generer(resultat As ResultatChangeWU,
                                   comptes As ComptesSystemeWU,
                                   decoupage As DecoupageChangeWU,
                                   ByRef messageErreur As String) As List(Of PieceChangeWU)

        messageErreur = String.Empty

        If resultat Is Nothing OrElse resultat.Ecarts.Count = 0 Then
            messageErreur = "Aucune transaction retenue : il n'y a aucun écart de change à comptabiliser."
            Return Nothing
        End If

        ' UN CALCUL INCOHÉRENT NE DEVIENT PAS UNE PIÈCE. Le contrôle arithmétique du résultat
        ' est une identité : s'il tombe, c'est que le parcours a compté une ligne d'un côté et
        ' pas de l'autre, et la pièce porterait ce défaut sans le montrer.
        If Not resultat.EstCoherent Then
            messageErreur = "Le contrôle arithmétique du calcul est en échec : " &
                            $"écart de {resultat.EcartDeControle} entre les deux façons d'obtenir le net." &
                            Environment.NewLine &
                            "Aucune pièce n'est produite tant que ce défaut n'est pas élucidé."
            Return Nothing
        End If

        Dim parametrage As ComptesSystemeWU = If(comptes, ComptesSystemeWU.Actuels)

        Dim manquants As List(Of String) = parametrage.ComptesDeChangeManquants()
        If manquants.Count > 0 Then
            messageErreur = MessageDesComptesManquants(manquants)
            Return Nothing
        End If

        Dim pieces As New List(Of PieceChangeWU)

        For Each groupe As IGrouping(Of String, EcartChangeWU) In
            resultat.Ecarts.Where(Function(ligne) ligne.DateReglement.HasValue).
                            GroupBy(Function(ligne) CleDuGroupe(ligne, decoupage)).
                            OrderBy(Function(parGroupe) parGroupe.Key)

            Dim piece As PieceChangeWU = ConstruireUnePiece(groupe, decoupage, parametrage, resultat)
            If piece IsNot Nothing Then pieces.Add(piece)
        Next

        ' LES TRANSACTIONS SANS DATE DE RÈGLEMENT NE SONT PAS PERDUES EN SILENCE. Une pièce se
        ' rattache à une journée : sans date, elle n'a pas de journée. Le cas ne se présente
        ' pas sur les rapports connus — les trois dates de règlement y sont toujours posées —
        ' mais s'il se présentait, le taire reviendrait à perdre un écart sans le dire.
        Dim sansDate As Integer = resultat.Ecarts.Where(Function(ligne) Not ligne.DateReglement.HasValue).Count()

        If pieces.Count = 0 Then
            messageErreur = If(sansDate > 0,
                               $"Aucune pièce : les {sansDate} transactions retenues ne portent pas de date " &
                               "de règlement exploitable, et une pièce se rattache à une journée.",
                               "Aucune pièce : tous les écarts de change sont nuls sur ce rapport.")
            Return Nothing
        End If

        If sansDate > 0 Then
            messageErreur = $"AVERTISSEMENT : {sansDate} transaction(s) retenue(s) sans date de règlement " &
                            "exploitable n'entrent dans aucune pièce. Leur écart n'est pas comptabilisé."
        End If

        Return pieces
    End Function

    ''' <summary>
    ''' Le message qui dit quoi faire, et non seulement que ça ne marche pas.
    '''
    ''' L'agent qui le lira n'est pas celui qui paramètre les comptes : il doit pouvoir le
    ''' transmettre tel quel à la Direction Comptable, et celle-ci doit y trouver le chemin
    ''' exact de l'écran et le nom exact des deux champs.
    ''' </summary>
    Private Shared Function MessageDesComptesManquants(manquants As List(Of String)) As String

        Dim liste As String = String.Join(Environment.NewLine & "    - ", manquants)

        Return "La pièce de change ne peut pas être produite : les comptes suivants ne sont pas " &
               "paramétrés." & Environment.NewLine & Environment.NewLine &
               "    - " & liste & Environment.NewLine & Environment.NewLine &
               "Renseignez-les dans Paramétrage > Comptes systèmes." & Environment.NewLine &
               "Si les champs « Gain de change » et « Perte de change » n'y figurent pas, c'est que " &
               "le script Scripts\21_EcartsDeChange.sql n'a pas encore été exécuté sur cette base." &
               Environment.NewLine & Environment.NewLine &
               "La pièce principale n'est pas concernée : elle continue de se produire normalement."
    End Function

    ''' <summary>
    ''' La clé qui regroupe les transactions d'une même pièce. La date est écrite en
    ''' aaaa-mm-jj pour que le tri alphabétique des clés soit le tri chronologique des pièces.
    ''' </summary>
    Private Shared Function CleDuGroupe(ligne As EcartChangeWU, decoupage As DecoupageChangeWU) As String

        Dim jour As String = ligne.DateReglement.Value.ToString("yyyy-MM-dd",
                                                                Globalization.CultureInfo.InvariantCulture)

        Select Case decoupage

            Case DecoupageChangeWU.ParJourneeEtSens
                Return $"{jour}|{ligne.Sens}"

            Case DecoupageChangeWU.ParJourneeSensEtProduit
                Return $"{jour}|{ligne.Sens}|{ligne.CodeProduit}"

            Case Else
                Return jour
        End Select
    End Function

    Private Shared Function ConstruireUnePiece(groupe As IGrouping(Of String, EcartChangeWU),
                                               decoupage As DecoupageChangeWU,
                                               comptes As ComptesSystemeWU,
                                               resultat As ResultatChangeWU) As PieceChangeWU

        Dim lignes As List(Of EcartChangeWU) = groupe.ToList()

        Dim gains As Long = Arrondir(lignes.Where(Function(l) l.Ecart > 0D).Sum(Function(l) l.Ecart))
        Dim pertes As Long = Arrondir(-lignes.Where(Function(l) l.Ecart < 0D).Sum(Function(l) l.Ecart))

        ' Un groupe dont tous les écarts sont nuls ne donne aucune écriture, et donc aucune
        ' pièce : une pièce vide n'a rien à faire dans un journal.
        If gains = 0L AndAlso pertes = 0L Then Return Nothing

        Dim premiere As EcartChangeWU = lignes(0)

        Dim sens As String = If(decoupage = DecoupageChangeWU.ParJournee, String.Empty, premiere.Sens)
        Dim produit As String = If(decoupage = DecoupageChangeWU.ParJourneeSensEtProduit,
                                   premiere.CodeProduit, String.Empty)

        Dim piece As New PieceChangeWU()
        piece.CleGroupe = groupe.Key
        piece.DateReglement = premiere.DateReglement.Value
        piece.Sens = sens
        piece.CodeProduit = produit
        piece.NombreTransactions = lignes.Count
        piece.Gains = gains
        piece.Pertes = pertes
        piece.Parite = premiere.Parite
        piece.FichierSource = resultat.FichierSource
        piece.Libelle = Narratif(sens, produit, piece.DateReglement, lignes.Count)

        piece.Lignes = TableVide()

        ' L'ordre : le gain d'abord, la perte ensuite, et dans chaque paire le débit avant le
        ' crédit. C'est l'ordre de lecture d'un comptable, et il est le même sur toutes les
        ' pièces du projet.
        If gains > 0L Then
            Ajouter(piece.Lignes, comptes.CompteCourant, piece.Libelle, gains, 0L)
            Ajouter(piece.Lignes, comptes.CompteGainDeChange, piece.Libelle, 0L, gains)
        End If

        If pertes > 0L Then
            Ajouter(piece.Lignes, comptes.ComptePerteDeChange, piece.Libelle, pertes, 0L)
            Ajouter(piece.Lignes, comptes.CompteCourant, piece.Libelle, 0L, pertes)
        End If

        Return piece
    End Function

    ''' <summary>
    ''' Arrondit un total au franc.
    '''
    ''' LES ÉCARTS SONT DÉJÀ DES ENTIERS : chaque transaction a été arrondie au franc par
    ''' EcartChangeWU.ContreValeur, et une somme d'entiers est entière. Cet arrondi ne corrige
    ''' donc rien — il CONVERTIT, et sert de garde-fou : si une évolution introduisait un
    ''' centime quelque part, la pièce resterait en francs entiers au lieu de refuser de
    ''' s'écrire sur une colonne BIGINT.
    ''' </summary>
    Private Shared Function Arrondir(montant As Decimal) As Long
        Return Convert.ToInt64(Math.Round(montant, 0, MidpointRounding.AwayFromZero))
    End Function

    Private Shared Sub Ajouter(table As DataTable, compte As String, libelle As String,
                               debit As Long, credit As Long)

        Dim ligne As DataRow = table.NewRow()
        ligne("Compte") = If(compte, String.Empty).Trim()
        ligne("Libelle") = libelle
        ligne("Debit") = debit
        ligne("Credit") = credit

        ' TOUTES LES LIGNES PORTENT LE CODE DU SIÈGE. L'écart de change est celui de la BANQUE :
        ' il naît de la différence entre ce qu'elle a encaissé en francs et ce que Western Union
        ' lui a réglé en euros. Aucun point de vente n'en est comptable, et le rattacher à l'un
        ' d'eux ferait apparaître dans ses états un résultat qui n'est pas le sien.
        ligne("CodeAgence") = ConstantesWU.CB_AGENCE_SIEGE

        table.Rows.Add(ligne)
    End Sub

    ''' <summary>
    ''' La table des écritures, vide. MÊME FORME QUE LA PIÈCE PRINCIPALE : c'est ce qui rend
    ''' l'export Excel, le fichier core banking et le contrôle d'équilibre réutilisables tels
    ''' quels. Voir PieceComptableService.GenererPieceComptable, qui pose les mêmes colonnes.
    ''' </summary>
    Public Shared Function TableVide() As DataTable

        Dim table As New DataTable("PieceChange")

        table.Columns.Add("Compte", GetType(String))
        table.Columns.Add("Libelle", GetType(String))
        table.Columns.Add("Debit", GetType(Long))
        table.Columns.Add("Credit", GetType(Long))
        table.Columns.Add("CodeAgence", GetType(String))

        Return table
    End Function

#End Region

#Region "Restitution"

    ''' <summary>
    ''' Toutes les écritures de toutes les pièces, dans une seule table.
    '''
    ''' POUR L'AFFICHAGE ET L'EXPORT, JAMAIS POUR LE CORE BANKING. Le fichier d'interface se
    ''' construit PIÈCE PAR PIÈCE : chaque pièce est une écriture comptable distincte, avec sa
    ''' journée et son numéro de lot. Les concaténer pour les charger d'un coup mélangerait
    ''' trois journées sous un seul lot, et le core banking n'y verrait qu'une pièce.
    ''' </summary>
    Public Shared Function ConsoliderLesLignes(pieces As IEnumerable(Of PieceChangeWU)) As DataTable

        Dim table As DataTable = TableVide()
        If pieces Is Nothing Then Return table

        For Each piece As PieceChangeWU In pieces

            If piece.Lignes Is Nothing Then Continue For

            For Each ligne As DataRow In piece.Lignes.Rows
                table.Rows.Add(ligne.ItemArray)
            Next
        Next

        Return table
    End Function

    ''' <summary>
    ''' La synthèse des pièces, telle qu'elle s'affiche et se colle dans un courriel.
    ''' Les totaux sont recalculés depuis les PIÈCES, et non repris du résultat du calcul :
    ''' c'est ainsi qu'on vérifie que rien n'a été perdu entre les deux.
    ''' </summary>
    Public Shared Function Synthese(pieces As List(Of PieceChangeWU),
                                    decoupage As DecoupageChangeWU) As String

        Dim fr As Globalization.CultureInfo = Globalization.CultureInfo.GetCultureInfo("fr-FR")
        Dim texte As New System.Text.StringBuilder()

        texte.AppendLine("PIÈCES COMPTABLES DE CHANGE")
        texte.AppendLine(New String("="c, 78))
        texte.AppendLine()
        texte.AppendLine($"Découpage : {LibelleDuDecoupage(decoupage)}")

        If pieces Is Nothing OrElse pieces.Count = 0 Then
            texte.AppendLine()
            texte.AppendLine("Aucune pièce.")
            Return texte.ToString()
        End If

        texte.AppendLine($"Pièces   : {pieces.Count:N0}")
        texte.AppendLine()
        texte.AppendLine("Journée      Sens  Prod.    Trx        Gains      Pertes         Net  Éq.")
        texte.AppendLine(New String("-"c, 78))

        ' Les colonnes sont composées AVANT d'être alignées. Une expression un peu longue
        ' glissée dans le trou d'une chaîne interpolée, avec en plus une largeur d'alignement,
        ' se compile mais ne se relit pas — et c'est précisément ce tableau que la banque
        ' lira pour rapprocher ses pièces.
        For Each piece As PieceChangeWU In pieces

            Dim sens As String = If(piece.Sens.Length = 0, "-", piece.Sens)
            Dim produit As String = If(piece.CodeProduit.Length = 0, "-", piece.CodeProduit)
            Dim equilibre As String = If(piece.EstEquilibree, "ok", "NON")

            texte.AppendLine(
                $"{piece.DateReglement:dd/MM/yyyy}   {sens,-4}  {produit,-6} " &
                $"{piece.NombreTransactions,6} " &
                $"{Montant(piece.Gains, fr),12} " &
                $"{Montant(piece.Pertes, fr),11} " &
                $"{Montant(piece.Net, fr),11}  " & equilibre)
        Next

        Dim gains As Decimal = pieces.Sum(Function(p) p.Gains)
        Dim pertes As Decimal = pieces.Sum(Function(p) p.Pertes)
        Dim debits As Long = pieces.Sum(Function(p) p.TotalDebit)
        Dim credits As Long = pieces.Sum(Function(p) p.TotalCredit)
        Dim transactions As Integer = pieces.Sum(Function(p) p.NombreTransactions)
        Dim ecritures As Integer = pieces.Where(Function(p) p.Lignes IsNot Nothing).
                                          Sum(Function(p) p.Lignes.Rows.Count)

        texte.AppendLine(New String("-"c, 78))
        texte.AppendLine($"TOTAL        {transactions,19} " &
                         $"{Montant(gains, fr),12} {Montant(pertes, fr),11} " &
                         $"{Montant(gains - pertes, fr),11}")
        texte.AppendLine()
        texte.AppendLine($"Écritures                : {ecritures:N0}")
        texte.AppendLine($"Total des débits         : {Montant(debits, fr)} {ConstantesWU.DEVISE_FCFA}")
        texte.AppendLine($"Total des crédits        : {Montant(credits, fr)} {ConstantesWU.DEVISE_FCFA}")

        Dim verdict As String = "OK — toutes les pièces sont équilibrées."
        If debits <> credits Then
            verdict = $"EN ÉCHEC : {Montant(debits - credits, fr)} d'écart entre débits et crédits."
        End If
        texte.AppendLine($"Contrôle d'équilibre     : {verdict}")

        Return texte.ToString()
    End Function

    ''' <summary>Un montant en francs entiers, séparateurs de milliers à la française.</summary>
    Private Shared Function Montant(valeur As Decimal, culture As Globalization.CultureInfo) As String
        Return valeur.ToString("N0", culture)
    End Function

    Public Shared Function LibelleDuDecoupage(decoupage As DecoupageChangeWU) As String

        Select Case decoupage
            Case DecoupageChangeWU.ParJourneeEtSens : Return "une pièce par journée et par sens"
            Case DecoupageChangeWU.ParJourneeSensEtProduit : Return "une pièce par journée, sens et code produit"
            Case Else : Return "une pièce par journée de règlement"
        End Select
    End Function

#End Region

End Class
