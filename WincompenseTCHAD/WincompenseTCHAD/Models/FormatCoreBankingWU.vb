Option Strict On
Option Explicit On

''' <summary>
''' LES CODES DU FORMAT CORE BANKING, tels que la banque les a fixés.
'''
''' POURQUOI UN PARAMÈTRE, ET NON SEPT CONSTANTES DE PLUS
'''
''' La banque en a déjà changé quatre d'un coup : BRN de N01 à S51, TXNCD débit de U24 à F03,
''' TXNCD crédit de F15 à F57, et COST_CENTER retirée. Chacun a coûté un cycle complet —
''' modification du code, compilation, commit, pull, redéploiement sur tous les postes — pour
''' des codes que la banque connaît et que nous, non.
'''
''' ET L'ENJEU EST PLUS GRAND QUE CELUI DE LA NARRATIVE. Un libellé mal écrit fait une ligne
''' laide dans le grand livre ; un TXNCD mal écrit fait une écriture du MAUVAIS TYPE, ou un lot
''' rejeté le matin de la compense. D'où le contrôle à la saisie, le journal des modifications,
''' et l'aperçu d'une ligne avant d'enregistrer.
'''
''' CE QUI N'EST PAS ICI, ET N'Y SERA PAS
'''
''' DRCR (« D » / « C ») n'est pas un code de la banque : c'est la sémantique du débit et du
''' crédit. Le rendre modifiable offrirait d'inverser tous les sens d'un fichier en une saisie,
''' et rien ne pourrait dire que c'est faux. AMOUNT, ACNO, VALDT et ADDLTEXT sont des données,
''' calculées ligne par ligne. INSTR_NO est toujours vide. BATCHNO est une règle, pas une
''' valeur.
'''
''' CETTE CLASSE EST IMMUABLE, ET NE TOUCHE NI LA BASE NI L'ÉCRAN. Elle reçoit sept valeurs et
''' répond aux questions du format. C'est ce qui permet de la rejouer hors application, et ce
''' qui garantit qu'un fichier en cours de production ne changera pas de codes parce que
''' quelqu'un enregistre au même moment.
''' </summary>
Public NotInheritable Class FormatCoreBankingWU

#Region "Construction"

    ''' <summary>
    ''' CHAQUE VALEUR VIDE RETOMBE SUR SA CONSTANTE, et c'est la règle qui protège les bases
    ''' non mises à jour : tant que le script n'a pas été rejoué, l'application produit
    ''' exactement le fichier d'aujourd'hui.
    ''' </summary>
    Public Sub New(detbsjrnl As String, brn As String, srccode As String,
                   txncdDebit As String, txncdCredit As String, agenceSiege As String,
                   comptesDuSiege As IEnumerable(Of String))

        _detbsjrnl = Retenu(detbsjrnl, ConstantesWU.CB_DETBSJRNL)
        _brn = Retenu(brn, ConstantesWU.CB_BRN)
        _srccode = Retenu(srccode, ConstantesWU.CB_SRCCODE)
        _txncdDebit = Retenu(txncdDebit, ConstantesWU.CB_TXNCD_DEBIT)
        _txncdCredit = Retenu(txncdCredit, ConstantesWU.CB_TXNCD_CREDIT)
        _agenceSiege = Retenu(agenceSiege, ConstantesWU.CB_AGENCE_SIEGE)

        ' RECOPIÉE, et non référencée : l'appelant garde sa liste et peut la remplir encore,
        ' un fichier en cours de production ne doit pas en voir le contenu changer.
        _comptesDuSiege = New List(Of String)

        If comptesDuSiege Is Nothing Then
            _comptesDuSiege.AddRange(ConstantesWU.CB_COMPTES_SIEGE)
        Else
            For Each compte As String In comptesDuSiege
                Dim net As String = If(compte, String.Empty).Trim()
                If net.Length = 0 Then Continue For
                If _comptesDuSiege.Contains(net, StringComparer.OrdinalIgnoreCase) Then Continue For
                _comptesDuSiege.Add(net)
            Next
        End If
    End Sub

    ''' <summary>La valeur saisie si elle dit quelque chose, la constante du code sinon.</summary>
    Private Shared Function Retenu(saisi As String, parDefaut As String) As String

        Dim net As String = If(saisi, String.Empty).Trim()
        If net.Length = 0 Then Return parDefaut

        Return net
    End Function

    Private ReadOnly _detbsjrnl As String
    Private ReadOnly _brn As String
    Private ReadOnly _srccode As String
    Private ReadOnly _txncdDebit As String
    Private ReadOnly _txncdCredit As String
    Private ReadOnly _agenceSiege As String
    Private ReadOnly _comptesDuSiege As List(Of String)

#End Region

#Region "Les sept valeurs"

    ''' <summary>Colonne DETBSJRNL, identique sur toutes les lignes du fichier.</summary>
    Public ReadOnly Property Detbsjrnl As String
        Get
            Return _detbsjrnl
        End Get
    End Property

    ''' <summary>
    ''' Colonne BRN : l'agence DU LOT, portée à l'identique par toutes ses lignes.
    '''
    ''' ELLE N'EST PAS L'AGENCE DU SIÈGE. BRN désigne le lot, ACBRN désigne CHAQUE écriture —
    ''' et vaut AgenceSiege pour les comptes internes. Un fichier porte aujourd'hui S51 en tête
    ''' de chaque ligne et N01 dans ACBRN pour ces comptes-là. Les confondre les ferait bouger
    ''' ensemble, ce que la banque n'a jamais demandé.
    ''' </summary>
    Public ReadOnly Property Brn As String
        Get
            Return _brn
        End Get
    End Property

    ''' <summary>Colonne SRCCODE, identique sur toutes les lignes du fichier.</summary>
    Public ReadOnly Property Srccode As String
        Get
            Return _srccode
        End Get
    End Property

    ''' <summary>Code transaction d'un débit (colonne TXNCD).</summary>
    Public ReadOnly Property TxncdDebit As String
        Get
            Return _txncdDebit
        End Get
    End Property

    ''' <summary>Code transaction d'un crédit (colonne TXNCD).</summary>
    Public ReadOnly Property TxncdCredit As String
        Get
            Return _txncdCredit
        End Get
    End Property

    ''' <summary>
    ''' Agence portée en ACBRN par les comptes du siège, et par les lignes qui ne se rattachent
    ''' à aucun point de vente — l'écart d'arrondi, les écritures de change.
    ''' </summary>
    Public ReadOnly Property AgenceSiege As String
        Get
            Return _agenceSiege
        End Get
    End Property

    ''' <summary>
    ''' Les comptes rattachés d'office à l'agence du siège, sans doublon et dans l'ordre saisi.
    '''
    ''' CE N'EST PAS UN PLAN COMPTABLE, c'est une règle d'AIGUILLAGE : aucune écriture ne leur
    ''' est passée du fait de cette liste, elle impose seulement leur agence en ACBRN. C'est
    ''' pourquoi ils ne vivent pas dans SystemeWU, qui reproduit le formulaire « Comptes
    ''' Systèmes WU » de la Direction Comptable.
    ''' </summary>
    Public ReadOnly Property ComptesDuSiege As IList(Of String)
        Get
            Return _comptesDuSiege.AsReadOnly()
        End Get
    End Property

#End Region

#Region "Ce que le format répond"

    ''' <summary>Le repli employé tant que rien n'a pu être lu : le comportement d'aujourd'hui.</summary>
    Public Shared Function ParDefaut() As FormatCoreBankingWU
        Return New FormatCoreBankingWU(String.Empty, String.Empty, String.Empty,
                                       String.Empty, String.Empty, String.Empty, Nothing)
    End Function

    ''' <summary>Le code transaction du sens demandé.</summary>
    Public Function Txncd(auDebit As Boolean) As String
        Return If(auDebit, _txncdDebit, _txncdCredit)
    End Function

    ''' <summary>Vrai si ce compte est rattaché d'office à l'agence du siège.</summary>
    Public Function EstCompteDuSiege(compte As String) As Boolean

        Dim net As String = If(compte, String.Empty).Trim()
        If net.Length = 0 Then Return False

        Return _comptesDuSiege.Contains(net, StringComparer.OrdinalIgnoreCase)
    End Function

    ''' <summary>
    ''' Agence à porter en ACBRN, dans cet ordre :
    '''
    '''   — les comptes du siège y sont rattachés d'office, quel que soit le point de vente qui
    '''     les a mouvementés ;
    '''   — sinon, le code agence du point de vente, tel que la pièce le porte ;
    '''   — à défaut, le siège : c'est le cas de la ligne d'écart d'arrondi, et celui d'un point
    '''     de vente dont le code agence manque encore. Laisser la colonne vide ferait rejeter
    '''     le fichier entier.
    ''' </summary>
    Public Function AgenceDeLaLigne(compte As String, codeAgence As String) As String

        If EstCompteDuSiege(compte) Then Return _agenceSiege

        Dim agence As String = If(codeAgence, String.Empty).Trim()
        If agence.Length > 0 Then Return agence

        Return _agenceSiege
    End Function

#End Region

#Region "Le découpage de la liste"

    ''' <summary>
    ''' Séparateur de la liste des comptes du siège, telle qu'elle est stockée en base.
    '''
    ''' UNE SEULE OPTION PLUTÔT QU'UNE TABLE : deux numéros de compte sans attribut ne
    ''' justifient pas une table à eux. Le journal des modifications y gagne — il montre
    ''' « avant : 379100319,379200585 », ce qui se lit. Le jour où la banque voudrait y
    ''' attacher un libellé ou une date d'effet, il faudra une table ; pas avant.
    ''' </summary>
    Public Const SEPARATEUR_COMPTES As Char = ","c

    ''' <summary>Les comptes portés par une valeur d'option, sans vide ni doublon.</summary>
    Public Shared Function ComptesDepuisTexte(texte As String) As List(Of String)

        Dim comptes As New List(Of String)

        For Each brut As String In If(texte, String.Empty).Split(SEPARATEUR_COMPTES)

            Dim net As String = brut.Trim()
            If net.Length = 0 Then Continue For
            If comptes.Contains(net, StringComparer.OrdinalIgnoreCase) Then Continue For

            comptes.Add(net)
        Next

        Return comptes
    End Function

    ''' <summary>La valeur d'option correspondant à une liste de comptes.</summary>
    Public Shared Function TexteDepuisComptes(comptes As IEnumerable(Of String)) As String

        If comptes Is Nothing Then Return String.Empty

        Dim retenus As New List(Of String)

        For Each brut As String In comptes

            Dim net As String = If(brut, String.Empty).Trim()
            If net.Length = 0 Then Continue For
            If retenus.Contains(net, StringComparer.OrdinalIgnoreCase) Then Continue For

            retenus.Add(net)
        Next

        Return String.Join(SEPARATEUR_COMPTES.ToString(), retenus)
    End Function

#End Region

#Region "Contrôle"

    ''' <summary>
    ''' Dit si un code est enregistrable, et pourquoi il ne l'est pas.
    '''
    ''' LE CONTRÔLE EST VOLONTAIREMENT PRUDENT, PAS PERMISSIF. La banque n'a pas encore dit la
    ''' forme exacte de ses codes — longueur fixe ? jeu de caractères ? casse imposée ? Tant
    ''' qu'elle ne l'a pas dite, refuser un code au motif qu'il ne ressemble pas aux cinq que
    ''' nous connaissons reviendrait à deviner sa règle, et à bloquer une valeur légitime le
    ''' jour où elle en donnera une sixième.
    '''
    ''' CE QUI EST REFUSÉ, C'EST CE QUI EST FAUX DANS TOUS LES CAS : un code vide, un code qui
    ''' dépasse la colonne, un code qui porte le séparateur de liste ou un caractère de
    ''' contrôle À L'INTÉRIEUR du code. QUAND LA BANQUE AURA RÉPONDU, c'est ici — et nulle
    ''' part ailleurs — qu'on resserrera.
    '''
    ''' L'INTÉRIEUR, ET PAS LES BORDS : le Trim ci-dessous enlève les blancs qui encadrent la
    ''' saisie, tabulation et retour à la ligne compris. Un code collé depuis un courriel
    ''' arrive souvent avec une tabulation à la fin ; la rogner et accepter le code est le bon
    ''' geste, car la valeur enregistrée est le code net. Ce qui reste interdit, c'est le
    ''' caractère invisible qui coupe le code en deux : celui-là part dans le fichier.
    ''' </summary>
    Public Shared Function ControlerLeCode(valeur As String, intitule As String,
                                           ByRef messageErreur As String) As Boolean

        messageErreur = String.Empty

        Dim net As String = If(valeur, String.Empty).Trim()

        If net.Length = 0 Then
            messageErreur = $"{intitule} est vide. Le core banking attend une valeur sur " &
                            "chaque ligne du fichier ; une colonne vide fait rejeter le lot entier."
            Return False
        End If

        If net.Length > ConstantesWU.CB_CODE_LONGUEUR_MAX Then
            messageErreur = $"{intitule} fait {net.Length} caractères, pour un maximum de " &
                            $"{ConstantesWU.CB_CODE_LONGUEUR_MAX}."
            Return False
        End If

        If net.IndexOf(SEPARATEUR_COMPTES) >= 0 Then
            messageErreur = $"{intitule} contient une virgule. Elle sépare les comptes du siège " &
                            "et ne peut pas figurer dans un code."
            Return False
        End If

        For Each caractere As Char In net
            If Char.IsControl(caractere) Then
                messageErreur = $"{intitule} contient un caractère invisible. Il est probablement " &
                                "arrivé par un copier-coller ; ressaisissez le code à la main."
                Return False
            End If
        Next

        Return True
    End Function

    ''' <summary>
    ''' Dit si la liste des comptes du siège est enregistrable.
    '''
    ''' UNE LISTE VIDE EST ACCEPTÉE, et ce n'est pas un oubli : une banque peut n'avoir aucun
    ''' compte à rattacher d'office au siège. Chaque ligne, en revanche, suit la règle des
    ''' codes — et le doublon est signalé plutôt que silencieusement absorbé, pour que celui
    ''' qui a saisi deux fois le même compte le sache.
    ''' </summary>
    Public Shared Function ControlerLesComptes(comptes As IEnumerable(Of String),
                                               ByRef messageErreur As String) As Boolean

        messageErreur = String.Empty

        If comptes Is Nothing Then Return True

        Dim vus As New List(Of String)

        For Each brut As String In comptes

            Dim net As String = If(brut, String.Empty).Trim()
            If net.Length = 0 Then Continue For

            If Not ControlerLeCode(net, $"Le compte du siège « {net} »", messageErreur) Then Return False

            If vus.Contains(net, StringComparer.OrdinalIgnoreCase) Then
                messageErreur = $"Le compte {net} figure deux fois dans la liste des comptes du siège."
                Return False
            End If

            vus.Add(net)
        Next

        Return True
    End Function

#End Region

End Class
