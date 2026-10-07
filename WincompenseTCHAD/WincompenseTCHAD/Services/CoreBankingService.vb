Option Strict On
Option Explicit On

Imports System.Data
Imports System.Globalization

''' <summary>
''' Fabrique le fichier d'interface chargé dans le core banking, à partir de la pièce comptable.
'''
''' La pièce reste l'objet de contrôle, lisible par un comptable. Ce fichier-ci est fait pour
''' être avalé par une machine, et il impactera réellement les comptes : il n'est donc produit
''' qu'à partir d'une pièce équilibrée, et jamais autrement.
''' </summary>
Public NotInheritable Class CoreBankingService

    Private Sub New()
    End Sub

    ''' <summary>
    ''' Les douze en-têtes, dans l'ordre exact attendu par le core banking.
    '''
    ''' ELLES ÉTAIENT TREIZE. COST_CENTER, qui portait 10000 sur toutes les lignes, a été
    ''' retirée sur demande écrite de la banque. L'ordre des douze restantes est inchangé.
    ''' </summary>
    Public Shared ReadOnly COLONNES As String() = New String() {
        "DETBSJRNL", "BRN", "BATCHNO", "SRCCODE", "AMOUNT", "ACNO", "DRCR",
        "ACBRN", "TXNCD", "VALDT", "INSTR_NO", "ADDLTEXT"
    }

    ''' <summary>
    ''' Alphabet du numéro de lot : chiffres puis MAJUSCULES, comme « 07P1 ».
    '''
    ''' IL ÉTAIT EN MINUSCULES, et la banque a demandé par écrit que le numéro de lot passe
    ''' en majuscules. La conséquence doit être connue : le numéro de lot est le SEUL élément
    ''' par lequel le core banking reconnaît un lot déjà chargé. Une journée chargée hier
    ''' sous « 07ob » se présentera désormais sous « 07OB » — si leur système distingue la
    ''' casse, il ne la reconnaîtra pas comme la même, et acceptera un second chargement.
    ''' La banque a confirmé le changement en connaissance de ce risque.
    '''
    ''' Le numéro de PIÈCE en hérite, lui aussi : PieceExcelWU.NumeroDePiece le compose à
    ''' partir du lot, et une pièce du 27/03 réimprimée portera « 07OB-001 » là où l'original
    ''' porte « 07ob-001 ». Même journée, même pièce, deux écritures du même numéro.
    ''' </summary>
    Private Const ALPHABET_LOT As String = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"

#Region "Numéro de lot"

    ''' <summary>
    ''' Numéro de lot d'une journée : quatre caractères, chiffres et minuscules.
    '''
    ''' Il n'est PAS tiré au hasard, et c'est délibéré. Le numéro de lot est le seul élément par
    ''' lequel le core banking peut reconnaître qu'on lui présente deux fois la même journée ; un
    ''' tirage aléatoire produirait deux numéros différents pour un même fichier réexporté, et
    ''' les comptes seraient impactés en double sans que rien ne le signale.
    '''
    ''' Il est donc dérivé de la date elle-même : le nombre de jours écoulés depuis une origine
    ''' fixe, écrit en base 36. Deux propriétés en découlent — une journée donne toujours le même
    ''' numéro, et deux journées n'en partagent jamais un, pendant plus de quatre mille ans.
    ''' L'origine est choisie pour que les numéros aient aujourd'hui la forme de ceux de la banque.
    '''
    ''' C'est la JOURNÉE D'ACTIVITÉ qui le détermine, et non la date de valeur. Celle-ci est le
    ''' jour de la compense : rattraper le lundi les journées du vendredi, du samedi et du
    ''' dimanche leur donnerait la même date de valeur. Trois fichiers distincts auraient alors
    ''' partagé un même numéro de lot, et le core banking les aurait pris pour trois chargements
    ''' du même.
    ''' </summary>
    Public Shared Function NumeroDeLot(dateActivite As Date) As String

        Dim jours As Integer = CInt((dateActivite.Date - ConstantesWU.CB_ORIGINE_LOT.Date).TotalDays)

        ' Une date antérieure à l'origine n'a pas de sens ici, mais ne doit pas faire échouer
        ' l'export : elle repart de zéro plutôt que de produire un numéro négatif.
        If jours < 0 Then jours = 0

        Dim caracteres(ConstantesWU.CB_LONGUEUR_LOT - 1) As Char

        For position As Integer = ConstantesWU.CB_LONGUEUR_LOT - 1 To 0 Step -1
            caracteres(position) = ALPHABET_LOT(jours Mod ALPHABET_LOT.Length)
            jours \= ALPHABET_LOT.Length
        Next

        Return New String(caracteres)
    End Function

    ''' <summary>
    ''' Numéro de lot de la pièce des ÉCARTS DE CHANGE d'une journée.
    '''
    ''' LE PROBLÈME QU'IL RÉSOUT. Le numéro de lot d'une journée de compensation est dérivé de
    ''' cette journée : le 27/03/2026 donne « 07OB », toujours. La pièce de change porte sur la
    ''' MÊME journée ; produite avec la même règle, elle arriverait au core banking sous « 07OB »
    ''' elle aussi. Le core banking n'a que ce numéro pour reconnaître un lot déjà chargé : il
    ''' rejetterait le second fichier comme un doublon du premier, ou les confondrait.
    '''
    ''' LA RÈGLE. Le numéro garde la journée, mais dans un ESPACE DISTINCT : la lettre « C »
    ''' pour change, suivie de la journée sur trois caractères en base 36. Le 27/03/2026 donne
    ''' donc « C7OB » là où la compensation donne « 07OB » : même journée reconnaissable, deux
    ''' lots qui ne peuvent pas se confondre.
    '''
    ''' POURQUOI AUCUNE COLLISION N'EST POSSIBLE. Un lot de compensation commence par le
    ''' quotient de la journée par 46 656 : il vaut « 0 » aujourd'hui et n'atteindrait « C »
    ''' qu'au bout de 559 872 jours, soit en 3531. Les trois caractères du lot de change
    ''' couvrent de leur côté 46 656 jours à partir de 1999, c'est-à-dire jusqu'en 2126.
    '''
    ''' DEUX PROPRIÉTÉS CONSERVÉES, les mêmes que pour la compensation : une journée donne
    ''' toujours le même numéro — un fichier réexporté est reconnu comme le même — et deux
    ''' journées n'en partagent jamais un.
    ''' </summary>
    Public Shared Function NumeroDeLotDeChange(dateActivite As Date) As String

        Dim jours As Integer = CInt((dateActivite.Date - ConstantesWU.CB_ORIGINE_LOT.Date).TotalDays)
        If jours < 0 Then jours = 0

        ' Au-delà de la capacité des trois caractères, on repart de zéro plutôt que de produire
        ' un numéro plus long que ce que le core banking accepte. Le cas se présenterait en 2126.
        Dim capacite As Integer = ALPHABET_LOT.Length * ALPHABET_LOT.Length * ALPHABET_LOT.Length
        jours = jours Mod capacite

        Dim caracteres(ConstantesWU.CB_LONGUEUR_LOT - 2) As Char

        For position As Integer = caracteres.Length - 1 To 0 Step -1
            caracteres(position) = ALPHABET_LOT(jours Mod ALPHABET_LOT.Length)
            jours \= ALPHABET_LOT.Length
        Next

        Return ConstantesWU.CB_LOT_PREFIXE_CHANGE & New String(caracteres)
    End Function

#End Region

#Region "Construction du fichier"

    ''' <summary>
    ''' Transforme la pièce comptable en table à douze colonnes, sous le numéro de lot de sa
    ''' journée.
    ''' </summary>
    ''' <param name="dtPiece">Pièce comptable produite par PieceComptableService.</param>
    ''' <param name="dateActivite">Journée traitée. Elle détermine le numéro de lot.</param>
    ''' <param name="dateValeur">Date portée par toutes les lignes (VALDT) : le jour de la compense.</param>
    ''' <param name="messageErreur">Motif du refus, le cas échéant.</param>
    ''' <returns>La table à charger, ou Nothing si la pièce ne s'y prête pas.</returns>
    Public Shared Function Construire(dtPiece As DataTable, dateActivite As Date, dateValeur As Date,
                                      ByRef messageErreur As String) As DataTable

        Return ConstruireSousLot(dtPiece, NumeroDeLot(dateActivite), dateValeur, messageErreur)
    End Function

    ''' <summary>
    ''' Même construction, mais sous un numéro de lot IMPOSÉ.
    '''
    ''' POURQUOI CETTE PORTE D'ENTRÉE EXISTE. La pièce des écarts de change porte sur la même
    ''' journée que la pièce de compensation, et le numéro de lot est dérivé de la journée :
    ''' les deux fichiers arriveraient donc au core banking SOUS LE MÊME NUMÉRO. Or ce numéro
    ''' est précisément ce par quoi le core banking reconnaît qu'on lui présente deux fois le
    ''' même lot — le second serait rejeté comme doublon, ou pire, fondu dans le premier.
    '''
    ''' L'appelant qui impose un lot doit donc en garantir l'unicité. Pour le change, c'est
    ''' NumeroDeLotDeChange qui s'en charge.
    ''' </summary>
    Public Shared Function ConstruireSousLot(dtPiece As DataTable, numeroLotImpose As String,
                                             dateValeur As Date,
                                             ByRef messageErreur As String) As DataTable

        messageErreur = String.Empty

        If String.IsNullOrWhiteSpace(numeroLotImpose) Then
            messageErreur = "Aucun numéro de lot : le fichier ne peut pas être produit."
            Return Nothing
        End If

        If dtPiece Is Nothing OrElse dtPiece.Rows.Count = 0 Then
            messageErreur = "La pièce comptable est vide : il n'y a rien à charger."
            Return Nothing
        End If

        If Not ControlerEquilibre(dtPiece, messageErreur) Then Return Nothing
        If Not ControlerLesNarratifs(dtPiece, messageErreur) Then Return Nothing

        Dim numeroLot As String = numeroLotImpose.Trim()
        Dim table As DataTable = TableVide()

        For Each ligne As DataRow In dtPiece.Rows

            Dim compte As String = Convert.ToString(ligne("Compte"), CultureInfo.InvariantCulture)
            Dim narratif As String = NarratifDe(ligne)
            Dim codeAgence As String = CodeAgenceDe(ligne)

            Dim debit As Long = Convert.ToInt64(ligne("Debit"), CultureInfo.InvariantCulture)
            Dim credit As Long = Convert.ToInt64(ligne("Credit"), CultureInfo.InvariantCulture)

            ' Une ligne de pièce porte un débit OU un crédit. Les deux à zéro n'arrivent pas —
            ' les auxiliaires de la pièce écartent déjà ce cas — mais une telle ligne n'aurait
            ' rien à impacter et n'a pas à encombrer le fichier.
            If debit <> 0L Then
                Ajouter(table, compte, narratif, debit, True, codeAgence, dateValeur, numeroLot)
            End If

            If credit <> 0L Then
                Ajouter(table, compte, narratif, credit, False, codeAgence, dateValeur, numeroLot)
            End If
        Next

        If table.Rows.Count = 0 Then
            messageErreur = "Aucune ligne de la pièce ne porte de montant : il n'y a rien à charger."
            Return Nothing
        End If

        Return table
    End Function

    ''' <summary>
    ''' Refuse une pièce dont un narratif dépasse ce que le core banking accepte.
    '''
    ''' La banque a confirmé 150 caractères pour ADDLTEXT. Un narratif plus long serait
    ''' tronqué par son système, ou ferait rejeter le lot : dans les deux cas, personne ne
    ''' s'en apercevrait au moment de l'export, et la découverte se ferait au chargement,
    ''' c'est-à-dire chez eux.
    '''
    ''' LE FICHIER N'EST PAS PRODUIT, ET ON NE TRONQUE PAS. Couper à 150 laisserait partir
    ''' un narratif amputé à mi-mot, qui ne dirait plus de quel sous-agent il s'agit — un
    ''' défaut silencieux là où l'arrêt est bruyant. Le message nomme la ligne fautive et sa
    ''' longueur, pour que le modèle de narrative, ou la désignation du point de vente, soit
    ''' raccourci.
    '''
    ''' LE CONTRÔLE PORTE SUR CE QUI PART RÉELLEMENT, c'est-à-dire sur NarratifDe — la
    ''' narrative du point de vente, et non le libellé de la pièce, qui lui n'est contraint
    ''' que par les 255 caractères de la colonne.
    ''' </summary>
    Private Shared Function ControlerLesNarratifs(dtPiece As DataTable, ByRef messageErreur As String) As Boolean

        For Each ligne As DataRow In dtPiece.Rows

            Dim narratif As String = NarratifDe(ligne)
            If narratif Is Nothing OrElse narratif.Length <= ConstantesWU.CB_NARRATIF_LONGUEUR_MAX Then Continue For

            messageErreur =
                $"Un narratif dépasse ce que le core banking accepte : {narratif.Length} caractères " &
                $"pour un maximum de {ConstantesWU.CB_NARRATIF_LONGUEUR_MAX}." &
                Environment.NewLine & Environment.NewLine &
                narratif & Environment.NewLine & Environment.NewLine &
                "Le fichier n'est pas produit. Le narratif n'est pas tronqué non plus : coupé, il ne " &
                "dirait plus de quel point de vente il s'agit. Raccourcissez le modèle dans l'écran " &
                "« Narrative comptable », ou la désignation de ce point de vente dans le référentiel, " &
                "puis régénérez la pièce."

            Return False
        Next

        Return True
    End Function

    ''' <summary>
    ''' Refuse une pièce déséquilibrée.
    '''
    ''' Le contrôle est refait ici alors que la pièce a déjà été équilibrée à sa génération :
    ''' ce fichier impacte des comptes réels, et rien ne garantit qu'il soit produit dans la
    ''' foulée de la génération.
    ''' </summary>
    Private Shared Function ControlerEquilibre(dtPiece As DataTable, ByRef messageErreur As String) As Boolean

        Dim totalDebit As Long = 0L
        Dim totalCredit As Long = 0L

        For Each ligne As DataRow In dtPiece.Rows
            totalDebit += Convert.ToInt64(ligne("Debit"), CultureInfo.InvariantCulture)
            totalCredit += Convert.ToInt64(ligne("Credit"), CultureInfo.InvariantCulture)
        Next

        If totalDebit = totalCredit Then Return True

        messageErreur =
            $"La pièce n'est pas équilibrée : {totalDebit:N0} FCFA au débit contre " &
            $"{totalCredit:N0} FCFA au crédit, soit un écart de {totalDebit - totalCredit:N0} FCFA." &
            Environment.NewLine & Environment.NewLine &
            "Le fichier n'est pas produit : chargé tel quel, il déséquilibrerait la comptabilité " &
            "de la banque. Régénérez la pièce comptable."

        Return False
    End Function

    ''' <summary>
    ''' LE NARRATIF QUE LE CORE BANKING VERRA, pour cette ligne de pièce.
    '''
    ''' LA PIÈCE COMPTABLE ET LE FICHIER CORE BANKING NE DISENT PAS LA MÊME CHOSE, et c'est
    ''' la banque qui l'a voulu ainsi : sur la pièce, chaque ligne porte le libellé de SA
    ''' NATURE (« COMPTE COURANT WESTERN UNION ETD », « TVA COLLECTEES WESTERN UNION »…) ;
    ''' dans le fichier, les douze lignes d'un même point de vente portent UNE SEULE ET MÊME
    ''' narrative, celle du point de vente. La pièce transporte les deux : son libellé dans
    ''' la colonne Libelle, et la narrative du point de vente dans la colonne Narratif.
    '''
    ''' LA COLONNE PEUT MANQUER — pièce rechargée depuis une base antérieure à cette
    ''' colonne, ou ligne d'écart d'arrondi, qui n'appartient à aucun point de vente et n'a
    ''' donc pas de narrative à elle. On retombe alors sur le libellé, qui dit déjà ce que
    ''' la ligne est : le fichier reste produisible dans tous les cas.
    ''' </summary>
    Private Shared Function NarratifDe(ligne As DataRow) As String

        If ligne.Table.Columns.Contains("Narratif") AndAlso Not ligne.IsNull("Narratif") Then
            Dim narratif As String = Convert.ToString(ligne("Narratif"), CultureInfo.InvariantCulture)
            If Not String.IsNullOrWhiteSpace(narratif) Then Return narratif
        End If

        Return Convert.ToString(ligne("Libelle"), CultureInfo.InvariantCulture)
    End Function

    ''' <summary>
    ''' Code agence de la ligne. La colonne peut manquer sur une pièce construite par une
    ''' version antérieure : on renvoie alors une chaîne vide, et la règle d'aiguillage
    ''' rattachera la ligne au siège.
    ''' </summary>
    Private Shared Function CodeAgenceDe(ligne As DataRow) As String

        If Not ligne.Table.Columns.Contains("CodeAgence") Then Return String.Empty
        If ligne.IsNull("CodeAgence") Then Return String.Empty

        Return Convert.ToString(ligne("CodeAgence"), CultureInfo.InvariantCulture)
    End Function

    Private Shared Function TableVide() As DataTable

        Dim table As New DataTable("CoreBanking")

        For Each colonne As String In COLONNES
            table.Columns.Add(colonne, GetType(String))
        Next

        Return table
    End Function

    Private Shared Sub Ajouter(table As DataTable, compte As String, libelle As String,
                               montant As Long, auDebit As Boolean, codeAgence As String,
                               dateValeur As Date, numeroLot As String)

        Dim ligne As LigneCoreBankingWU = LigneCoreBankingWU.Construire(
            compte, libelle, montant, auDebit, codeAgence, dateValeur, numeroLot)

        table.Rows.Add(ligne.DETBSJRNL,
                       ligne.BRN,
                       ligne.BATCHNO,
                       ligne.SRCCODE,
                       ligne.AMOUNT.ToString(CultureInfo.InvariantCulture),
                       ligne.ACNO,
                       ligne.DRCR,
                       ligne.ACBRN,
                       ligne.TXNCD,
                       ligne.VALDT.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture),
                       ligne.INSTR_NO,
                       ligne.ADDLTEXT)
    End Sub

#End Region

#Region "Nom du fichier"

    ''' <summary>
    ''' Nom proposé pour le fichier. Le core banking n'en impose aucun ; celui-ci porte la
    ''' journée d'activité et son numéro de lot, de sorte qu'un fichier retrouvé dans un dossier
    ''' se rattache sans ambiguïté à la journée qu'il comptabilise.
    ''' </summary>
    Public Shared Function NomDeFichier(dateActivite As Date) As String

        Return $"WU_CORE_{dateActivite:yyyyMMdd}_{NumeroDeLot(dateActivite)}.xlsx"
    End Function

    ''' <summary>
    ''' Nom proposé pour le fichier des ÉCARTS DE CHANGE. Il porte CHANGE dans son nom et le
    ''' lot de change dans son suffixe : deux fichiers retrouvés dans le même dossier ne
    ''' peuvent pas être confondus, ni l'un chargé à la place de l'autre.
    ''' </summary>
    Public Shared Function NomDeFichierDeChange(dateActivite As Date) As String

        Return $"WU_CHANGE_CORE_{dateActivite:yyyyMMdd}_{NumeroDeLotDeChange(dateActivite)}.xlsx"
    End Function

#End Region

End Class
