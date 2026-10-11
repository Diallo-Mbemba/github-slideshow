Option Strict On
Option Explicit On

''' <summary>
''' Une ligne du fichier d'interface chargé dans le core banking.
'''
''' Ce fichier n'est pas une pièce comptable de plus : c'est la forme sous laquelle le core
''' banking accepte d'impacter réellement les comptes des sous-agents et les comptes internes
''' de la banque. Il compte douze colonnes, dont huit sont constantes ou déduites — seuls le
''' montant, le compte, le sens et le libellé viennent vraiment de la pièce, plus la date.
'''
''' Les noms des propriétés reprennent exactement ceux des colonnes attendues, en majuscules :
''' c'est sous ces libellés que le fichier est lu, et les renommer au goût du français les
''' rendrait introuvables.
''' </summary>
Public Class LigneCoreBankingWU

#Region "Les douze colonnes, dans l'ordre du fichier"

    ' LES VALEURS PAR DÉFAUT RESTENT CELLES DU CODE, et c'est voulu : une ligne construite
    ' sans format — un test, un appel oublié — sort comme aujourd'hui plutôt que vide. Le
    ' format paramétré, lui, est posé par Construire.
    Public Property DETBSJRNL As String = ConstantesWU.CB_DETBSJRNL
    Public Property BRN As String = ConstantesWU.CB_BRN
    Public Property BATCHNO As String = String.Empty
    Public Property SRCCODE As String = ConstantesWU.CB_SRCCODE
    Public Property AMOUNT As Long = 0L
    Public Property ACNO As String = String.Empty
    Public Property DRCR As String = String.Empty
    Public Property ACBRN As String = String.Empty
    Public Property TXNCD As String = String.Empty
    Public Property VALDT As Date
    Public Property INSTR_NO As String = String.Empty
    Public Property ADDLTEXT As String = String.Empty
    ' COST_CENTER A ÉTÉ RETIRÉE, sur demande écrite de la banque. Elle portait la valeur
    ' constante 10000 sur toutes les lignes depuis l'origine. Ni la propriété ni la colonne
    ' ne subsistent : une colonne conservée « au cas où » finit par être réalimentée.

#End Region

#Region "Construction depuis une ligne de pièce"

    ''' <summary>
    ''' Construit la ligne d'interface correspondant à une ligne de la pièce comptable.
    ''' </summary>
    ''' <param name="compte">Compte mouvementé.</param>
    ''' <param name="libelle">Libellé de l'écriture.</param>
    ''' <param name="montant">Montant, toujours positif : le sens est porté par DRCR.</param>
    ''' <param name="auDebit">Vrai pour un débit, faux pour un crédit.</param>
    ''' <param name="codeAgence">Code agence du point de vente, ou vide s'il n'y en a pas.</param>
    ''' <param name="dateCompensation">Date de valeur : la date de compensation (J+1).</param>
    ''' <param name="numeroLot">Numéro de lot, commun à toutes les lignes du fichier.</param>
    ''' <param name="format">
    ''' Les codes du format, lus UNE FOIS pour tout le fichier. Nothing retombe sur ceux du
    ''' code : un appelant qui ne les fournit pas obtient le fichier d'aujourd'hui, pas un
    ''' fichier aux colonnes vides.
    ''' </param>
    Public Shared Function Construire(compte As String, libelle As String, montant As Long,
                                      auDebit As Boolean, codeAgence As String,
                                      dateCompensation As Date, numeroLot As String,
                                      format As FormatCoreBankingWU) As LigneCoreBankingWU

        Dim codes As FormatCoreBankingWU = If(format, FormatCoreBankingWU.ParDefaut())

        ' DRCR N'EST PAS PARAMÉTRABLE, et ne le sera pas : « D » et « C » ne sont pas des codes
        ' de la banque, c'est la sémantique du débit et du crédit. Les rendre modifiables
        ' offrirait d'inverser tous les sens d'un fichier en une saisie.
        Return New LigneCoreBankingWU() With {
            .DETBSJRNL = codes.Detbsjrnl,
            .BRN = codes.Brn,
            .BATCHNO = numeroLot,
            .SRCCODE = codes.Srccode,
            .AMOUNT = Math.Abs(montant),
            .ACNO = If(compte, String.Empty).Trim(),
            .DRCR = If(auDebit, ConstantesWU.CB_SENS_DEBIT, ConstantesWU.CB_SENS_CREDIT),
            .ACBRN = codes.AgenceDeLaLigne(compte, codeAgence),
            .TXNCD = codes.Txncd(auDebit),
            .VALDT = dateCompensation.Date,
            .ADDLTEXT = If(libelle, String.Empty).Trim()
        }
    End Function


#End Region

End Class
