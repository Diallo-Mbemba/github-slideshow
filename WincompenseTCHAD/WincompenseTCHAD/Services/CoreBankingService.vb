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

    ''' <summary>Les treize en-têtes, dans l'ordre exact attendu par le core banking.</summary>
    Public Shared ReadOnly COLONNES As String() = New String() {
        "DETBSJRNL", "BRN", "BATCHNO", "SRCCODE", "AMOUNT", "ACNO", "DRCR",
        "ACBRN", "TXNCD", "VALDT", "INSTR_NO", "ADDLTEXT", "COST_CENTER"
    }

    ''' <summary>Alphabet du numéro de lot : chiffres puis minuscules, comme « 07p1 ».</summary>
    Private Const ALPHABET_LOT As String = "0123456789abcdefghijklmnopqrstuvwxyz"

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
    ''' C'est la JOURNÉE D'ACTIVITÉ qui le détermine, et non la date de valeur. Depuis que celle-ci
    ''' est reportée au premier jour ouvré, l'activité du vendredi, du samedi et du dimanche porte
    ''' la même date de valeur — le lundi. Trois fichiers distincts auraient alors partagé un même
    ''' numéro de lot, et le core banking les aurait pris pour trois chargements du même.
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

#End Region

#Region "Construction du fichier"

    ''' <summary>
    ''' Transforme la pièce comptable en table à treize colonnes.
    ''' </summary>
    ''' <param name="dtPiece">Pièce comptable produite par PieceComptableService.</param>
    ''' <param name="dateActivite">Journée traitée. Elle détermine le numéro de lot.</param>
    ''' <param name="dateValeur">Date portée par toutes les lignes : le premier jour ouvré suivant.</param>
    ''' <param name="messageErreur">Motif du refus, le cas échéant.</param>
    ''' <returns>La table à charger, ou Nothing si la pièce ne s'y prête pas.</returns>
    Public Shared Function Construire(dtPiece As DataTable, dateActivite As Date, dateValeur As Date,
                                      ByRef messageErreur As String) As DataTable

        messageErreur = String.Empty

        If dtPiece Is Nothing OrElse dtPiece.Rows.Count = 0 Then
            messageErreur = "La pièce comptable est vide : il n'y a rien à charger."
            Return Nothing
        End If

        If Not ControlerEquilibre(dtPiece, messageErreur) Then Return Nothing

        Dim numeroLot As String = NumeroDeLot(dateActivite)
        Dim table As DataTable = TableVide()

        For Each ligne As DataRow In dtPiece.Rows

            Dim compte As String = Convert.ToString(ligne("Compte"), CultureInfo.InvariantCulture)
            Dim libelle As String = Convert.ToString(ligne("Libelle"), CultureInfo.InvariantCulture)
            Dim codeAgence As String = CodeAgenceDe(ligne)

            Dim debit As Long = Convert.ToInt64(ligne("Debit"), CultureInfo.InvariantCulture)
            Dim credit As Long = Convert.ToInt64(ligne("Credit"), CultureInfo.InvariantCulture)

            ' Une ligne de pièce porte un débit OU un crédit. Les deux à zéro n'arrivent pas —
            ' les auxiliaires de la pièce écartent déjà ce cas — mais une telle ligne n'aurait
            ' rien à impacter et n'a pas à encombrer le fichier.
            If debit <> 0L Then
                Ajouter(table, compte, libelle, debit, True, codeAgence, dateValeur, numeroLot)
            End If

            If credit <> 0L Then
                Ajouter(table, compte, libelle, credit, False, codeAgence, dateValeur, numeroLot)
            End If
        Next

        If table.Rows.Count = 0 Then
            messageErreur = "Aucune ligne de la pièce ne porte de montant : il n'y a rien à charger."
            Return Nothing
        End If

        Return table
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
                       ligne.ADDLTEXT,
                       ligne.COST_CENTER)
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

#End Region

End Class
