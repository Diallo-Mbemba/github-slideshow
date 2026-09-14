Option Strict On
Option Explicit On

''' <summary>
''' Une ligne du fichier d'interface chargé dans le core banking.
'''
''' Ce fichier n'est pas une pièce comptable de plus : c'est la forme sous laquelle le core
''' banking accepte d'impacter réellement les comptes des sous-agents et les comptes internes
''' de la banque. Il compte treize colonnes, dont neuf sont constantes ou déduites — seuls le
''' montant, le compte, le sens et le libellé viennent vraiment de la pièce, plus la date.
'''
''' Les noms des propriétés reprennent exactement ceux des colonnes attendues, en majuscules :
''' c'est sous ces libellés que le fichier est lu, et les renommer au goût du français les
''' rendrait introuvables.
''' </summary>
Public Class LigneCoreBankingWU

#Region "Les treize colonnes, dans l'ordre du fichier"

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
    Public Property COST_CENTER As String = ConstantesWU.CB_COST_CENTER

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
    Public Shared Function Construire(compte As String, libelle As String, montant As Long,
                                      auDebit As Boolean, codeAgence As String,
                                      dateCompensation As Date, numeroLot As String) As LigneCoreBankingWU

        Return New LigneCoreBankingWU() With {
            .BATCHNO = numeroLot,
            .AMOUNT = Math.Abs(montant),
            .ACNO = If(compte, String.Empty).Trim(),
            .DRCR = If(auDebit, ConstantesWU.CB_SENS_DEBIT, ConstantesWU.CB_SENS_CREDIT),
            .ACBRN = AgenceDeLaLigne(compte, codeAgence),
            .TXNCD = If(auDebit, ConstantesWU.CB_TXNCD_DEBIT, ConstantesWU.CB_TXNCD_CREDIT),
            .VALDT = dateCompensation.Date,
            .ADDLTEXT = If(libelle, String.Empty).Trim()
        }
    End Function

    ''' <summary>
    ''' Agence à porter dans ACBRN.
    '''
    ''' Trois cas, dans cet ordre :
    '''   — les comptes du siège y sont rattachés d'office, quel que soit le point de vente qui
    '''     les a mouvementés ;
    '''   — sinon, le code agence du point de vente, lu dans la base selon l'Account ;
    '''   — à défaut, le siège : c'est le cas de la ligne d'écart d'arrondi, qui n'appartient à
    '''     aucun point de vente, et celui d'un point de vente dont le code agence manque encore
    '''     dans la base. Laisser la colonne vide ferait rejeter le fichier entier.
    ''' </summary>
    Public Shared Function AgenceDeLaLigne(compte As String, codeAgence As String) As String

        Dim compteNettoye As String = If(compte, String.Empty).Trim()

        For Each compteSiege As String In ConstantesWU.CB_COMPTES_SIEGE
            If String.Equals(compteNettoye, compteSiege, StringComparison.OrdinalIgnoreCase) Then
                Return ConstantesWU.CB_AGENCE_SIEGE
            End If
        Next

        Dim agence As String = If(codeAgence, String.Empty).Trim()
        If agence.Length > 0 Then Return agence

        Return ConstantesWU.CB_AGENCE_SIEGE
    End Function

#End Region

End Class
