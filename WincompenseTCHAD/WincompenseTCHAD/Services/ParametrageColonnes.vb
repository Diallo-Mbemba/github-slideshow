Option Strict On
Option Explicit On

''' <summary>
''' Correspondance entre les comptes comptables de l'application et les colonnes de la table
''' SystemeWU.
'''
''' POURQUOI LES NOMS DE COLONNES SQL, ET NON DES NOMS INVENTÉS
'''
''' Le fichier exporté sera relu par un humain au moins autant que par l'application : un
''' auditeur, un administrateur qui compare deux installations, quelqu'un qui cherche pourquoi
''' une écriture part sur le mauvais compte. Écrire « Cpte_PositionNette » lui permet de
''' retrouver la ligne dans la base ; écrire « compteCourant » l'obligerait à connaître le
''' vocabulaire interne de l'application.
'''
''' LE COMPTE INTER BANCAIRE N'A QU'UNE LIGNE
'''
''' La table le porte dans deux colonnes, Cpte_attenteDEBIT et Cpte_attenteCREDIT, que la
''' ligne de paramétrage tient identiques et que l'écran présente comme un seul champ. Le
''' fichier en porte donc une seule ligne, sous le nom de la colonne DÉBIT, qui fait foi à la
''' lecture. Écrire deux lignes égales inviterait à les rendre différentes.
''' </summary>
Public NotInheritable Class ParametrageColonnes

    Private Sub New()
    End Sub

#Region "Noms des lignes du fichier"

    Public Const COMPTE_COURANT As String = "Cpte_PositionNette"
    Public Const INTER_BANCAIRE As String = "Cpte_attenteDEBIT"
    Public Const COMMISSION_TRANSFERT As String = "Cpte_Produit"
    Public Const COMMISSION_ENVOI As String = "Cpte_Produit_Envoi"
    Public Const COMMISSION_PAIEMENT As String = "Cpte_Produit_Paiement"
    Public Const IMPOTS_TAXE_ENVOI As String = "Tthu"
    Public Const TVA_COLLECTEE As String = "Tob"
    Public Const TTA_ENVOI As String = "Cpte_Envoi"
    Public Const TTA_RECEPTION As String = "Cpte_Paiement"

#End Region

#Region "Vers le fichier"

    ''' <summary>
    ''' Les comptes, ligne par ligne : nom de colonne, valeur, et ce que la ligne veut dire.
    ''' La troisième colonne n'est pas relue — elle est là pour celui qui ouvre le fichier.
    ''' </summary>
    Public Shared Function Paires(source As ComptesSystemeWU) As List(Of String())

        Dim lignes As New List(Of String())()
        If source Is Nothing Then Return lignes

        lignes.Add(New String() {COMPTE_COURANT, source.CompteCourant,
                                 "Compte courant Western Union"})
        lignes.Add(New String() {INTER_BANCAIRE, source.CompteInterBancaire,
                                 "Compte inter bancaire (écart d'arrondi)"})
        lignes.Add(New String() {COMMISSION_TRANSFERT, source.CommissionTransfertBanque,
                                 "Commission sur transfert revenant à la banque"})
        lignes.Add(New String() {COMMISSION_ENVOI, source.CommissionEnvoiBanque,
                                 "Commission sur envoi revenant à la banque"})
        lignes.Add(New String() {COMMISSION_PAIEMENT, source.CommissionPaiementBanque,
                                 "Commission sur paiement revenant à la banque"})
        lignes.Add(New String() {IMPOTS_TAXE_ENVOI, source.ImpotsTaxeEnvoi,
                                 "Impôts et taxe sur envoi"})
        lignes.Add(New String() {TVA_COLLECTEE, source.TVACollectee,
                                 "TVA collectée"})
        lignes.Add(New String() {TTA_ENVOI, source.TTAEnvoi,
                                 "TTA sur envoi"})
        lignes.Add(New String() {TTA_RECEPTION, source.TTAReception,
                                 "TTA sur paiement"})

        Return lignes
    End Function

#End Region

#Region "Depuis le fichier"

    ''' <summary>
    ''' Reconstitue les comptes à partir des lignes lues.
    '''
    ''' Une ligne absente laisse la valeur par défaut de l'application, comme le fait déjà la
    ''' lecture de la base : on ne remplace jamais un compte par une chaîne vide, ce qui
    ''' produirait une écriture sans compte.
    '''
    ''' CodeParametrage n'est volontairement PAS repris : il désigne la ligne de SystemeWU
    ''' dans la base D'ORIGINE, et la base d'arrivée a la sienne. C'est l'import qui rattache
    ''' les comptes à la ligne locale.
    ''' </summary>
    Public Shared Function Comptes(lignes As Dictionary(Of String, String)) As ComptesSystemeWU

        Dim resultat As New ComptesSystemeWU()
        If lignes Is Nothing Then Return resultat

        Affecter(lignes, COMPTE_COURANT, Sub(v) resultat.CompteCourant = v)
        Affecter(lignes, INTER_BANCAIRE, Sub(v) resultat.CompteInterBancaire = v)
        Affecter(lignes, COMMISSION_TRANSFERT, Sub(v) resultat.CommissionTransfertBanque = v)
        Affecter(lignes, COMMISSION_ENVOI, Sub(v) resultat.CommissionEnvoiBanque = v)
        Affecter(lignes, COMMISSION_PAIEMENT, Sub(v) resultat.CommissionPaiementBanque = v)
        Affecter(lignes, IMPOTS_TAXE_ENVOI, Sub(v) resultat.ImpotsTaxeEnvoi = v)
        Affecter(lignes, TVA_COLLECTEE, Sub(v) resultat.TVACollectee = v)
        Affecter(lignes, TTA_ENVOI, Sub(v) resultat.TTAEnvoi = v)
        Affecter(lignes, TTA_RECEPTION, Sub(v) resultat.TTAReception = v)

        Return resultat
    End Function

    Private Shared Sub Affecter(lignes As Dictionary(Of String, String), cle As String,
                                poser As Action(Of String))

        If Not lignes.ContainsKey(cle) Then Return

        Dim valeur As String = If(lignes(cle), String.Empty).Trim()
        If valeur.Length = 0 Then Return

        poser(valeur)
    End Sub

#End Region

#Region "Comparaison"

    ''' <summary>
    ''' Une différence entre deux paramétrages de comptes : ce que la base porte, ce que le
    ''' fichier propose. Montrée AVANT de remplacer — les neuf comptes déterminent toute la
    ''' pièce comptable, on ne les change pas à l'aveugle.
    ''' </summary>
    Public NotInheritable Class Difference

        Public Property Ligne As String = String.Empty
        Public Property Signification As String = String.Empty
        Public Property Actuel As String = String.Empty
        Public Property Propose As String = String.Empty

        Public ReadOnly Property Change As Boolean
            Get
                Return Not String.Equals(Actuel, Propose, StringComparison.OrdinalIgnoreCase)
            End Get
        End Property
    End Class

    ''' <summary>Compare les comptes de la base à ceux du fichier, ligne à ligne.</summary>
    Public Shared Function Comparer(actuels As ComptesSystemeWU,
                                    proposes As ComptesSystemeWU) As List(Of Difference)

        Dim differences As New List(Of Difference)()
        If proposes Is Nothing Then Return differences

        Dim avant As List(Of String()) = Paires(If(actuels, New ComptesSystemeWU()))
        Dim apres As List(Of String()) = Paires(proposes)

        For rang As Integer = 0 To apres.Count - 1

            Dim valeurAvant As String = If(rang < avant.Count, avant(rang)(1), String.Empty)

            differences.Add(New Difference() With {
                .Ligne = apres(rang)(0),
                .Signification = apres(rang)(2),
                .Actuel = If(valeurAvant, String.Empty),
                .Propose = If(apres(rang)(1), String.Empty)
            })
        Next

        Return differences
    End Function

#End Region

End Class
