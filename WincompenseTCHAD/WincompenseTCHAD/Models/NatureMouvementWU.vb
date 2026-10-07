Option Strict On
Option Explicit On

''' <summary>
''' LES QUATORZE NATURES DE MOUVEMENT D'UNE PIÈCE COMPTABLE.
'''
''' Une pièce porte douze lignes par point de vente, plus une treizième posée en fin de pièce
''' pour absorber l'écart d'arrondi global. Ce que chaque ligne EST se lit aujourd'hui dans
''' son numéro de compte ; cette énumération le nomme, pour que la banque puisse donner à
''' chacune le libellé qu'elle veut.
'''
''' POURQUOI UNE ÉNUMÉRATION, ET NON LE NUMÉRO DE COMPTE COMME CLÉ
'''
''' Deux natures PARTAGENT un compte : la commission de transfert et la commission d'envoi
''' part banque vont toutes deux sur 728300148 dans le paramétrage actuel, et les trois
''' commissions d'un sous-agent vont sur son unique CompteCommission. Une clé par compte
''' confondrait ce que la comptabilité distingue. Et les comptes CHANGENT — les deux comptes
''' de change viennent d'en faire la démonstration : une clé par compte serait orpheline au
''' premier changement de plan comptable.
''' </summary>
Public Enum NatureMouvementWU

    ''' <summary>Ligne de mouvement du point de vente (son compte de compensation, ou le compte inter bancaire).</summary>
    Mouvement = 1

    ''' <summary>Contrepartie sur le compte courant Western Union.</summary>
    CompteCourant = 2

    ''' <summary>Commission sur transfert, part banque.</summary>
    CommissionTransfertBanque = 3

    ''' <summary>Commission sur paiement, part banque.</summary>
    CommissionPaiementBanque = 4

    ''' <summary>Commission sur envoi, part banque.</summary>
    CommissionEnvoiBanque = 5

    ''' <summary>Commission sur transfert, part sous-agent.</summary>
    CommissionTransfertSousAgent = 6

    ''' <summary>Commission sur paiement, part sous-agent.</summary>
    CommissionPaiementSousAgent = 7

    ''' <summary>Commission sur envoi, part sous-agent.</summary>
    CommissionEnvoiSousAgent = 8

    ''' <summary>Impôts et taxe sur envoi.</summary>
    ImpotsTaxeEnvoi = 9

    ''' <summary>TVA collectée.</summary>
    TVA = 10

    ''' <summary>TTA sur envoi.</summary>
    TTAEnvoi = 11

    ''' <summary>TTA sur réception.</summary>
    TTAReception = 12

    ''' <summary>
    ''' Écart d'arrondi, posé en fin de pièce sur le compte inter bancaire.
    '''
    ''' ELLE N'EST PAS COMME LES AUTRES, et c'est écrit dans NarrativesWU : elle ne se
    ''' rattache à aucun point de vente, le repère principal du modèle global. Elle garde donc
    ''' TOUJOURS son libellé propre, dans les deux modes.
    ''' </summary>
    EcartArrondi = 13

    ''' <summary>
    ''' Ligne de mouvement d'une AGENCE PROPRE, posée sur le compte inter bancaire.
    '''
    ''' POURQUOI ELLE EST DISTINCTE DE Mouvement. Une agence propre n'a pas de compte de
    ''' compensation dans les livres de la banque : sa ligne va sur le compte inter bancaire,
    ''' que TOUTES les agences propres partagent. Le libellé de ce compte est celui du compte
    ''' — « VIREMENTS INTER-BANCAIRES EMIS » — et non le nom de l'agence, qui ferait croire à
    ''' un compte qui lui serait propre. La banque l'a relevé le 07/10/2026.
    '''
    ''' AJOUTÉE À LA FIN, ET NON INSÉRÉE : renuméroter l'énumération réaffecterait les
    ''' libellés déjà saisis par la banque à d'autres lignes de la pièce. Son rang dans
    ''' Toutes(), lui, la place à côté de Mouvement, où l'écran la lit.
    ''' </summary>
    MouvementInterBancaire = 14
End Enum

''' <summary>
''' Ce que chaque nature vaut en base, et ce qu'elle dit à l'écran.
'''
''' LE CODE STOCKÉ EST UNE CHAÎNE, PAS L'ENTIER DE L'ÉNUMÉRATION. Un entier rendrait la table
''' illisible dans Management Studio, et surtout : renuméroter l'énumération — en insérant une
''' nature au milieu, par exemple — réaffecterait silencieusement les libellés saisis par la
''' banque à d'autres lignes de la pièce. Le code, lui, ne bouge pas.
''' </summary>
Public NotInheritable Class NaturesMouvementWU

    Private Sub New()
    End Sub

    ''' <summary>Les quatorze natures, dans l'ordre où elles apparaissent sur la pièce.</summary>
    Public Shared Function Toutes() As List(Of NatureMouvementWU)

        Dim liste As New List(Of NatureMouvementWU)

        liste.Add(NatureMouvementWU.Mouvement)
        liste.Add(NatureMouvementWU.MouvementInterBancaire)
        liste.Add(NatureMouvementWU.CompteCourant)
        liste.Add(NatureMouvementWU.CommissionTransfertBanque)
        liste.Add(NatureMouvementWU.CommissionPaiementBanque)
        liste.Add(NatureMouvementWU.CommissionEnvoiBanque)
        liste.Add(NatureMouvementWU.CommissionTransfertSousAgent)
        liste.Add(NatureMouvementWU.CommissionPaiementSousAgent)
        liste.Add(NatureMouvementWU.CommissionEnvoiSousAgent)
        liste.Add(NatureMouvementWU.ImpotsTaxeEnvoi)
        liste.Add(NatureMouvementWU.TVA)
        liste.Add(NatureMouvementWU.TTAEnvoi)
        liste.Add(NatureMouvementWU.TTAReception)
        liste.Add(NatureMouvementWU.EcartArrondi)

        Return liste
    End Function

    ''' <summary>Le code tel qu'il est écrit dans T_NarrativeNatureWU.</summary>
    Public Shared Function Code(nature As NatureMouvementWU) As String

        Select Case nature
            Case NatureMouvementWU.Mouvement : Return "MOUVEMENT"
            Case NatureMouvementWU.MouvementInterBancaire : Return "MOUVEMENT_INTER_BANCAIRE"
            Case NatureMouvementWU.CompteCourant : Return "COMPTE_COURANT"
            Case NatureMouvementWU.CommissionTransfertBanque : Return "COMMISSION_TRANSFERT_BANQUE"
            Case NatureMouvementWU.CommissionPaiementBanque : Return "COMMISSION_PAIEMENT_BANQUE"
            Case NatureMouvementWU.CommissionEnvoiBanque : Return "COMMISSION_ENVOI_BANQUE"
            Case NatureMouvementWU.CommissionTransfertSousAgent : Return "COMMISSION_TRANSFERT_SA"
            Case NatureMouvementWU.CommissionPaiementSousAgent : Return "COMMISSION_PAIEMENT_SA"
            Case NatureMouvementWU.CommissionEnvoiSousAgent : Return "COMMISSION_ENVOI_SA"
            Case NatureMouvementWU.ImpotsTaxeEnvoi : Return "IMPOTS_TAXE_ENVOI"
            Case NatureMouvementWU.TVA : Return "TVA"
            Case NatureMouvementWU.TTAEnvoi : Return "TTA_ENVOI"
            Case NatureMouvementWU.TTAReception : Return "TTA_RECEPTION"
            Case NatureMouvementWU.EcartArrondi : Return "ECART_ARRONDI"
        End Select

        ' Une nature ajoutée à l'énumération et oubliée ici n'aurait pas de code, donc pas de
        ' ligne en base : son libellé retomberait silencieusement sur le modèle global. Le nom
        ' de l'énumération fait un code acceptable en attendant qu'on l'écrive au-dessus.
        Return nature.ToString().ToUpperInvariant()
    End Function

    ''' <summary>La nature portant ce code, ou Nothing si le code est inconnu.</summary>
    Public Shared Function DepuisCode(code As String) As NatureMouvementWU?

        Dim recherche As String = If(code, String.Empty).Trim()
        If recherche.Length = 0 Then Return Nothing

        For Each nature As NatureMouvementWU In Toutes()
            If String.Equals(NaturesMouvementWU.Code(nature), recherche,
                             StringComparison.OrdinalIgnoreCase) Then Return nature
        Next

        ' Un code inconnu — une nature retirée du code, ou une ligne saisie à la main en SQL —
        ' est IGNORÉ plutôt que rejeté : la pièce doit sortir, et elle sort avec le modèle
        ' global pour la nature que cette ligne prétendait décrire.
        Return Nothing
    End Function

    ''' <summary>Ce que la nature dit à l'écran, dans la grille des libellés.</summary>
    Public Shared Function Intitule(nature As NatureMouvementWU) As String

        Select Case nature
            Case NatureMouvementWU.Mouvement : Return "Mouvement d'un sous-agent"
            Case NatureMouvementWU.MouvementInterBancaire : Return "Mouvement d'une agence propre (compte inter bancaire)"
            Case NatureMouvementWU.CompteCourant : Return "Contrepartie compte courant WU"
            Case NatureMouvementWU.CommissionTransfertBanque : Return "Commission transfert — part banque"
            Case NatureMouvementWU.CommissionPaiementBanque : Return "Commission paiement — part banque"
            Case NatureMouvementWU.CommissionEnvoiBanque : Return "Commission envoi — part banque"
            Case NatureMouvementWU.CommissionTransfertSousAgent : Return "Commission transfert — part sous-agent"
            Case NatureMouvementWU.CommissionPaiementSousAgent : Return "Commission paiement — part sous-agent"
            Case NatureMouvementWU.CommissionEnvoiSousAgent : Return "Commission envoi — part sous-agent"
            Case NatureMouvementWU.ImpotsTaxeEnvoi : Return "Impôts et taxe sur envoi"
            Case NatureMouvementWU.TVA : Return "TVA collectée"
            Case NatureMouvementWU.TTAEnvoi : Return "TTA sur envoi"
            Case NatureMouvementWU.TTAReception : Return "TTA sur réception"
            Case NatureMouvementWU.EcartArrondi : Return "Écart d'arrondi (compte inter bancaire)"
        End Select

        Return Code(nature)
    End Function

    ''' <summary>
    ''' LE LIBELLÉ PAR DÉFAUT D'UNE NATURE, tel qu'il figurait sur la pièce avant que la
    ''' banque ne demande un texte unique — et tel qu'elle le redemande depuis son
    ''' rectificatif du 06/10/2026.
    '''
    ''' C'EST UN GABARIT, avec les mêmes repères que le modèle global. Une nature que la
    ''' banque n'a pas personnalisée porte celui-ci : elle n'a donc rien à saisir pour
    ''' retrouver la pièce d'avant.
    '''
    ''' CE N'EST PAS LE MODÈLE GLOBAL, et c'est tout le rectificatif. Le modèle global ne vaut
    ''' plus que pour la colonne ADDLTEXT du fichier core banking, où la banque veut une seule
    ''' phrase par point de vente ; la pièce comptable, elle, redit ce que chaque ligne est.
    ''' </summary>
    Public Shared Function ModeleParDefaut(nature As NatureMouvementWU) As String

        Select Case nature
            Case NatureMouvementWU.Mouvement : Return ConstantesWU.NARRATIVE_MOUVEMENT_DEFAUT
            Case NatureMouvementWU.MouvementInterBancaire : Return ConstantesWU.NARRATIVE_MOUVEMENT_INTER_BANCAIRE_DEFAUT
            Case NatureMouvementWU.CompteCourant : Return ConstantesWU.NARRATIVE_COMPTE_COURANT_DEFAUT
            Case NatureMouvementWU.CommissionTransfertBanque : Return ConstantesWU.NARRATIVE_COMMISSION_TRANSFERT_BANQUE_DEFAUT
            Case NatureMouvementWU.CommissionPaiementBanque : Return ConstantesWU.NARRATIVE_COMMISSION_PAIEMENT_BANQUE_DEFAUT
            Case NatureMouvementWU.CommissionEnvoiBanque : Return ConstantesWU.NARRATIVE_COMMISSION_ENVOI_BANQUE_DEFAUT
            Case NatureMouvementWU.CommissionTransfertSousAgent : Return ConstantesWU.NARRATIVE_COMMISSION_TRANSFERT_SA_DEFAUT
            Case NatureMouvementWU.CommissionPaiementSousAgent : Return ConstantesWU.NARRATIVE_COMMISSION_PAIEMENT_SA_DEFAUT
            Case NatureMouvementWU.CommissionEnvoiSousAgent : Return ConstantesWU.NARRATIVE_COMMISSION_ENVOI_SA_DEFAUT
            Case NatureMouvementWU.ImpotsTaxeEnvoi : Return ConstantesWU.NARRATIVE_IMPOTS_TAXE_ENVOI_DEFAUT
            Case NatureMouvementWU.TVA : Return ConstantesWU.NARRATIVE_TVA_DEFAUT
            Case NatureMouvementWU.TTAEnvoi : Return ConstantesWU.NARRATIVE_TTA_ENVOI_DEFAUT
            Case NatureMouvementWU.TTAReception : Return ConstantesWU.NARRATIVE_TTA_RECEPTION_DEFAUT
            Case NatureMouvementWU.EcartArrondi : Return ConstantesWU.NARRATIVE_ECART_DEFAUT
        End Select

        ' Une nature ajoutée à l'énumération et oubliée ici n'aurait pas de libellé : le modèle
        ' global est un repli acceptable, une ligne de pièce sans texte ne l'est pas.
        Return ConstantesWU.NARRATIVE_MODELE_DEFAUT
    End Function

#Region "Pour le journal"

    ''' <summary>Préfixe des clés de journal portant un libellé par nature.</summary>
    Public Const PREFIXE_CLE_JOURNAL As String = "NATURE_"

    ''' <summary>
    ''' La clé sous laquelle le journal enregistre le changement d'un libellé par nature.
    '''
    ''' Le journal est UNIQUE pour tout le paramétrage de narrative : le modèle global, le
    ''' mode, et les quatorze natures. Préfixer évite qu'une nature nommée « TVA » entre en
    ''' collision avec une option qui s'appellerait ainsi un jour.
    ''' </summary>
    Public Shared Function CleDeJournal(nature As NatureMouvementWU) As String
        Return PREFIXE_CLE_JOURNAL & Code(nature)
    End Function

    ''' <summary>
    ''' L'intitulé lisible d'une clé de journal, qu'elle désigne une nature ou une option.
    ''' Une clé inconnue est rendue telle quelle : mieux vaut un code brut qu'une ligne vide.
    ''' </summary>
    Public Shared Function IntituleDeCle(cle As String) As String

        Dim texte As String = If(cle, String.Empty).Trim()
        If texte.Length = 0 Then Return String.Empty

        If Not texte.StartsWith(PREFIXE_CLE_JOURNAL, StringComparison.OrdinalIgnoreCase) Then Return texte

        Dim nature As NatureMouvementWU? = DepuisCode(texte.Substring(PREFIXE_CLE_JOURNAL.Length))
        If Not nature.HasValue Then Return texte

        Return Intitule(nature.Value)
    End Function

#End Region

End Class
