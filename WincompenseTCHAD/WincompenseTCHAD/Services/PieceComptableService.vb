' Option Strict est désactivé UNIQUEMENT dans ce fichier car l'écriture du classeur utilise la
' liaison tardive (late binding) sur Microsoft Excel via Type.GetTypeFromProgID, afin de ne pas
' imposer de référence COM obligatoire au projet lorsque Excel n'est pas installé (section 1).
' Tout le reste du fichier (grille de contrôle, pièce comptable, équilibrage) reste fortement typé.
Option Strict Off
Option Explicit On

Imports System.Data
Imports System.Globalization
Imports System.Linq

''' <summary>
''' Construction de la grille de contrôle et génération de la pièce comptable définitive.
''' Contient également le contrôle d'équilibrage global (section 14) et l'export Excel optionnel.
''' </summary>
Public NotInheritable Class PieceComptableService

    Private Sub New()
    End Sub

#Region "Grille de contrôle (section 12)"

    ''' <summary>
    ''' Construit la DataTable de contrôle affichée dans dgvControle : une ligne par Account,
    ''' avec toutes les colonnes de calcul intermédiaires et finales.
    ''' Les colonnes ErreurSQL / DonneesManquantes sont incluses (masquables dans la grille)
    ''' pour permettre la mise en évidence visuelle des anomalies.
    ''' </summary>
    Public Shared Function CreerTableControle(listeCalculs As IEnumerable(Of CalculWU)) As DataTable

        Dim dt As New DataTable("dtControle")

        dt.Columns.Add("Account", GetType(String))
        dt.Columns.Add("Designation", GetType(String))
        dt.Columns.Add("Type", GetType(String))
        dt.Columns.Add("CodeAgence", GetType(String))
        dt.Columns.Add("CompteCompense", GetType(String))
        dt.Columns.Add("CompteCommission", GetType(String))
        dt.Columns.Add("TauxSA", GetType(Decimal))
        dt.Columns.Add("PrincipalEnvoi", GetType(Decimal))
        dt.Columns.Add("PrincipalPaye", GetType(Decimal))
        dt.Columns.Add("ChargeEnvoi", GetType(Decimal))
        dt.Columns.Add("Taxes", GetType(Decimal))
        dt.Columns.Add("CommissionTransfert", GetType(Decimal))
        dt.Columns.Add("CommissionPaiement", GetType(Decimal))
        dt.Columns.Add("CommissionEnvoi", GetType(Decimal))
        dt.Columns.Add("TVA", GetType(Decimal))
        dt.Columns.Add("TTAEnvoi", GetType(Decimal))
        dt.Columns.Add("TTAReception", GetType(Decimal))
        dt.Columns.Add("TaxeEnvoi", GetType(Decimal))
        dt.Columns.Add("CommissionTransfertBanque", GetType(Decimal))
        dt.Columns.Add("CommissionPaiementBanque", GetType(Decimal))
        dt.Columns.Add("CommissionEnvoiBanque", GetType(Decimal))
        dt.Columns.Add("CommissionTransfertSA", GetType(Decimal))
        dt.Columns.Add("CommissionPaiementSA", GetType(Decimal))
        dt.Columns.Add("CommissionEnvoiSA", GetType(Decimal))
        dt.Columns.Add("TotalDebit", GetType(Decimal))
        dt.Columns.Add("TotalCredit", GetType(Decimal))
        dt.Columns.Add("Solde", GetType(Decimal))
        dt.Columns.Add("EcartArrondi", GetType(Long))
        dt.Columns.Add("ErreurSQL", GetType(Boolean))
        dt.Columns.Add("DonneesManquantes", GetType(Boolean))

        ' Colonne technique, masquée dans la grille : elle porte le fait qu'un Account soit
        ' écarté de la pièce comptable, ce qui est une conséquence plus lourde qu'une simple
        ' donnée manquante et mérite sa propre mise en évidence.
        dt.Columns.Add("NonComptabilise", GetType(Boolean))

        If listeCalculs Is Nothing Then Return dt

        For Each calc As CalculWU In listeCalculs
            Dim ligne As DataRow = dt.NewRow()

            ligne("Account") = calc.Account
            ligne("Designation") = calc.Designation
            ligne("Type") = calc.TypePdv
            ligne("CodeAgence") = calc.CodeAgence
            ligne("CompteCompense") = calc.CompteCompense
            ligne("CompteCommission") = calc.CompteCommission
            ligne("TauxSA") = calc.TauxSA
            ligne("PrincipalEnvoi") = calc.PrincipalEnvoi
            ligne("PrincipalPaye") = calc.PrincipalPaye
            ligne("ChargeEnvoi") = calc.ChargeEnvoi
            ligne("Taxes") = calc.Taxes
            ligne("CommissionTransfert") = calc.CommissionTransfert
            ligne("CommissionPaiement") = calc.CommissionPaiement
            ligne("CommissionEnvoi") = calc.CommissionEnvoi
            ligne("TVA") = calc.TVA
            ligne("TTAEnvoi") = calc.TTAEnvoi
            ligne("TTAReception") = calc.TTAReception
            ligne("TaxeEnvoi") = calc.TaxeEnvoi
            ligne("CommissionTransfertBanque") = calc.CommissionTransfertBanque
            ligne("CommissionPaiementBanque") = calc.CommissionPaiementBanque
            ligne("CommissionEnvoiBanque") = calc.CommissionEnvoiBanque
            ligne("CommissionTransfertSA") = calc.CommissionTransfertSA
            ligne("CommissionPaiementSA") = calc.CommissionPaiementSA
            ligne("CommissionEnvoiSA") = calc.CommissionEnvoiSA
            ligne("TotalDebit") = calc.TotalDebit
            ligne("TotalCredit") = calc.TotalCredit
            ligne("Solde") = calc.Solde
            ligne("EcartArrondi") = CalculerEcartArrondi(calc)
            ligne("ErreurSQL") = calc.ErreurSQL
            ligne("DonneesManquantes") = calc.DonneesManquantes
            ligne("NonComptabilise") = Not calc.EstComptabilisable

            dt.Rows.Add(ligne)
        Next

        Return dt
    End Function

#End Region

#Region "Génération de la pièce comptable (section 13)"

    ''' <summary>
    ''' Construit la DataTable dtPiece (Compte, Libelle, Debit, Credit) à partir de la liste
    ''' des CalculWU, en reproduisant la structure du classeur de référence
    ''' PieceComptabilsationTchad.xlsx (validée sur un exemple réel de sous-agent) :
    '''
    '''   - UNE SEULE ligne de mouvement (Débit si positif, Crédit si négatif) sur le compte
    '''     de compensation du point de vente : CompteCompense pour un sous-agent, le compte
    '''     inter bancaire 381000101 « VIREMENTS INTERBANCAIRES ÉMISES » pour une agence propre
    '''     — celle-ci n'a pas de compte de compensation propre dans les livres de la banque
    '''     (confirmé par la banque). Pour le montant net :
    '''         NetMouvement = (PrincipalEnvoi + ChargeEnvoi + Taxes) − PrincipalPaye
    '''                        + TTAReception
    '''     La formule vit dans CalculWU.NetMouvement, et nulle part ailleurs. La TTA sur
    '''     réception s'y ajoute parce que la plateforme ne la prélève pas : à l'envoi elle
    '''     est déjà dans Taxes, à la réception elle n'est nulle part, et il faut bien
    '''     l'encaisser pour pouvoir la créditer plus bas.
    '''   - UNE SEULE ligne en contrepartie sur le compte courant WU (32100003292), pour la
    '''     part nette revenant à la banque une fois les commissions et taxes affectées :
    '''         NetCompteCourant = NetMouvement − (toutes commissions + toutes taxes)
    '''   - Les commissions (banque et sous-agent) et les taxes sont des lignes de CRÉDIT
    '''     uniquement (aucune ligne de débit miroir individuelle) : leur contrepartie débit
    '''     est absorbée globalement par la ligne de mouvement ci-dessus.
    '''
    ''' Cette structure a été vérifiée par rapprochement algébrique avec un exemple réel du
    ''' classeur de référence (agence BOLOLO) : les lignes de commissions/taxes et les
    ''' comptes correspondent à l'unité près, aux arrondis près (écart résiduel ≤ 1 FCFA dans
    ''' l'exemple, absorbé par le mécanisme du compte d'attente, section 14).
    ''' </summary>
    ''' <param name="listeCalculs">Les points de vente à comptabiliser.</param>
    ''' <param name="debutPeriode">Premier jour couvert. Nothing si la période est inconnue.</param>
    ''' <param name="finPeriode">Dernier jour couvert. Nothing pour une pièce d'une seule journée.</param>
    Public Shared Function GenererPieceComptable(listeCalculs As IEnumerable(Of CalculWU),
                                                 debutPeriode As Date?,
                                                 finPeriode As Date?) As DataTable

        Dim dt As New DataTable("dtPiece")
        dt.Columns.Add("Compte", GetType(String))
        dt.Columns.Add("Libelle", GetType(String))
        dt.Columns.Add("Debit", GetType(Long))
        dt.Columns.Add("Credit", GetType(Long))

        ' Code agence du point de vente qui a produit la ligne : codeagence pour un sous-agent,
        ' CodeAgenc-Voyager pour une agence propre, tel que WURepository l'a lu selon l'Account.
        ' Il ne s'affiche pas dans la pièce — il alimente la colonne ACBRN du fichier destiné au
        ' core banking, qui n'a aucun autre moyen de savoir à quelle agence rattacher l'écriture.
        dt.Columns.Add("CodeAgence", GetType(String))

        ' LA NARRATIVE DU POINT DE VENTE, à côté du libellé et non à sa place.
        '
        ' LA PIÈCE ET LE FICHIER CORE BANKING NE DISENT PAS LA MÊME CHOSE, et la banque l'a
        ' voulu ainsi : sur la pièce, chaque ligne porte le libellé de SA NATURE ; dans le
        ' fichier, les douze lignes d'un même point de vente portent UNE SEULE et même
        ' narrative. Un seul texte par ligne ne pourrait pas servir les deux.
        '
        ' Elle ne s'affiche pas dans la pièce — elle alimente la colonne ADDLTEXT du fichier,
        ' comme le code agence juste au-dessus alimente ACBRN.
        dt.Columns.Add("Narratif", GetType(String))


        If listeCalculs Is Nothing Then Return dt

        ' Comptes comptables en service, lus une fois pour toute la pièce : ils proviennent de la
        ' table SystemeWU (voir ComptesSystemeWU) et non plus de constantes figées dans le code.
        ' Les prendre ici, et non à chaque écriture, garantit qu'une même pièce ne mélange jamais
        ' deux paramétrages si les comptes venaient à être modifiés pendant sa génération.
        Dim comptes As ComptesSystemeWU = ComptesSystemeWU.Actuels

        ' La période, calculée UNE FOIS pour toute la pièce. La recalculer par ligne coûterait
        ' peu, mais laisserait la possibilité que deux lignes de la même pièce n'annoncent pas
        ' la même période — et c'est précisément ce qu'un narratif est censé rendre impossible.
        Dim periode As String = SuffixeDePeriode(debutPeriode, finPeriode)

        ' LE PARAMÉTRAGE DE NARRATIVE, lu UNE FOIS pour toute la pièce, et pour la même raison
        ' que les comptes juste au-dessus : une pièce ne doit jamais mélanger deux paramétrages
        ' si quelqu'un enregistrait un libellé pendant sa génération. L'objet rendu est
        ' immuable, et la lecture ne coûte rien — le mode, le modèle global et les quatorze
        ' libellés par nature sont gardés le temps de la session.
        Dim narratives As NarrativesWU = NarrativeRepository.EnVigueur()

        For Each calc As CalculWU In listeCalculs

            ' Un Account dont la banque ne connaît pas les comptes n'est pas comptabilisé : il
            ' est purement et simplement absent de la pièce (voir CalculWU.EstComptabilisable).
            ' Sa ligne de mouvement irait sinon sur le compte courant WU, c'est-à-dire sur un
            ' compte qui n'est pas le sien, et la correction se ferait à la main après coup.
            If Not calc.EstComptabilisable Then
                calc.EcartArrondi = 0L
                Continue For
            End If

            ' COMPTE DE LA LIGNE DE MOUVEMENT.
            '
            ' Un sous-agent a son propre compte de compensation dans les livres de la banque
            ' (T_Pdv_SA.CompteCompense) : son mouvement y va.
            '
            ' Une agence propre n'en a pas. Son mouvement va sur le COMPTE INTER BANCAIRE
            ' (381000101, « VIREMENTS INTERBANCAIRES ÉMISES ») — et non sur le compte courant
            ' Western Union, comme c'était le cas jusqu'ici. L'écriture ne quitte pas la banque :
            ' elle passe d'une agence à une autre, ce qu'un compte de virements interbancaires
            ' décrit exactement, là où le compte courant WU aurait porté deux fois le même
            ' numéro — une fois pour le mouvement, une fois pour sa propre contrepartie.
            ' Confirmé par la pièce manuelle de la banque pour AHB020013 (Ecobank AGP Siège).
            '
            ' Le numéro n'est pas écrit en dur : c'est celui du compte inter bancaire paramétré
            ' dans SystemeWU (colonnes Cpte_attenteDEBIT / Cpte_attenteCREDIT), le même que
            ' celui qui absorbe l'écart d'arrondi global en fin de pièce. C'est bien le même
            ' compte dans les livres de la banque ; le jour où elle les distinguerait, il
            ' suffirait d'ajouter une propriété à ComptesSystemeWU — ici et nulle part ailleurs.
            '
            ' La branche Else n'est atteinte QUE par une agence propre : un sous-agent sans
            ' compte de compensation n'est pas comptabilisable et a déjà été écarté plus haut
            ' par EstComptabilisable, tout comme un Account absent du paramétrage.
            Dim compteMouvement As String
            If String.Equals(calc.TypePdv, "SA", StringComparison.OrdinalIgnoreCase) AndAlso
               Not String.IsNullOrWhiteSpace(calc.CompteCompense) Then
                compteMouvement = calc.CompteCompense
            Else
                compteMouvement = comptes.CompteInterBancaire
            End If

            ' SA NATURE SUIT SON COMPTE, ET RIEN D'AUTRE. C'est le COMPTE qui donne son libellé
            ' à la ligne : un compte de compensation n'appartient qu'à son sous-agent, et porte
            ' donc son nom ; le compte inter bancaire est partagé, et porte le nom DU COMPTE —
            ' « VIREMENTS INTER-BANCAIRES EMIS ». La banque l'a relevé le 07/10/2026, puis
            ' précisé le lendemain : CHEZ LES SOUS-AGENTS AUSSI.
            '
            ' D'OÙ UNE COMPARAISON DE COMPTES, et non une reprise du If ci-dessus. Le type du
            ' point de vente n'est pas la question : un sous-agent dont le référentiel porterait
            ' le compte inter bancaire comme compte de compensation tomberait dans la première
            ' branche, et sortirait avec son nom sur un compte qui n'est pas le sien. En lisant
            ' le compte RETENU, la règle tient quel que soit le chemin qui y a mené.
            Dim natureMouvement As NatureMouvementWU =
                If(String.Equals(compteMouvement, comptes.CompteInterBancaire,
                                 StringComparison.OrdinalIgnoreCase),
                   NatureMouvementWU.MouvementInterBancaire, NatureMouvementWU.Mouvement)

            ' LES LIBELLÉS DES LIGNES DE CE POINT DE VENTE, un par nature de mouvement.
            '
            ' CHAQUE LIGNE DE LA PIÈCE PORTE LE LIBELLÉ DE SA NATURE — « COMPTE COURANT
            ' WESTERN UNION ETD », « Commission sur Transfert_Ecobank », « TVA COLLECTEES
            ' WESTERN UNION »… C'est la pièce manuelle de la banque : ni préfixe « LD », ni
            ' période, en majuscules, et avec ses espaces à elle. C'est le mode par défaut.
            '
            ' SEUL LE FICHIER CORE BANKING porte une seule et même narrative pour les douze
            ' lignes du point de vente (narratifPdv, juste en dessous), et c'est elle aussi
            ' que la RAISON reprend en bas de pièce. Les libellés, eux, ne la voient pas.
            '
            ' La banque garde la main : dans l'écran « Narrative comptable » elle personnalise
            ' chacun des quatorze libellés, et peut au besoin basculer la pièce elle aussi sur un
            ' seul modèle. Le code n'en sait rien et n'a pas à le savoir — il demande le
            ' libellé de la nature qu'il pose.
            '
            ' CALCULÉS TOUS D'UN COUP, et non ligne à ligne : les douze appels ci-dessous se
            ' contentent d'indexer leur nature, ce qui les laisse lisibles, et une seule
            ' résolution de mode a lieu par point de vente.
            Dim libelles As Dictionary(Of NatureMouvementWU, String) =
                narratives.Libelles(calc.Designation, calc.Account, calc.CodeAgence, periode)

            ' LA NARRATIVE DU POINT DE VENTE, la même pour ses douze lignes, destinée au seul
            ' fichier core banking. Elle sort du MODÈLE GLOBAL, et non de la nature : c'est
            ' exactement ce que la banque veut voir dans ADDLTEXT, et elle l'y veut quel que
            ' soit le mode choisi pour la pièce.
            Dim narratifPdv As String = ModeleNarrativeWU.Appliquer(
                narratives.ModeleGlobal, calc.Designation, calc.Account, calc.CodeAgence, periode)

            Dim totalCommissionsEtTaxes As Decimal =
                calc.CommissionTransfertBanque + calc.CommissionEnvoiBanque + calc.CommissionPaiementBanque +
                calc.CommissionTransfertSA + calc.CommissionPaiementSA + calc.CommissionEnvoiSA +
                calc.TaxeEnvoi + calc.TVA + calc.TTAEnvoi + calc.TTAReception

            Dim netMouvement As Decimal = calc.NetMouvement
            Dim netCompteCourant As Decimal = netMouvement - totalCommissionsEtTaxes

            ' Mémorise l'écart d'arrondi apporté par cet Account (visible dans la grille de contrôle).
            CalculerEcartArrondi(calc)

            ' 1) Ligne de mouvement (compte de compensation du point de vente).
            AjouterLigneSigneAuto(dt, compteMouvement, libelles(natureMouvement),
                                  netMouvement, calc.CodeAgence, narratifPdv)

            ' 2) Contrepartie sur le compte courant WU (part nette revenant à la banque).
            AjouterLigneSigneAuto(dt, comptes.CompteCourant, libelles(NatureMouvementWU.CompteCourant),
                                  -netCompteCourant, calc.CodeAgence, narratifPdv)

            ' 3) Commissions part Banque (toujours créditées, quel que soit le type de PDV).
            ' Transfert et Envoi partagent le même compte dans le paramétrage actuel (728300148),
            ' mais la table SystemeWU les porte dans deux colonnes distinctes (Cpte_Produit et
            ' Cpte_Produit_Envoi) : ils sont donc désormais dissociables sans toucher au code.
            ' L'ORDRE EST TRANSFERT, PAIEMENT, ENVOI — le même que pour le sous-agent plus bas,
            ' et le même que sur la pièce manuelle de la banque. Il ne l'était pas : le bloc
            ' banque allait transfert/envoi/paiement et le bloc sous-agent transfert/paiement/
            ' envoi. Dans une comparaison ligne à ligne, deux ordres différents font perdre du
            ' temps au vérificateur, et lui font passer des écarts.
            AjouterLigneSiNonNul(dt, comptes.CommissionTransfertBanque,
                                 libelles(NatureMouvementWU.CommissionTransfertBanque),
                                 0D, calc.CommissionTransfertBanque, calc.CodeAgence, narratifPdv)
            AjouterLigneSiNonNul(dt, comptes.CommissionPaiementBanque,
                                 libelles(NatureMouvementWU.CommissionPaiementBanque),
                                 0D, calc.CommissionPaiementBanque, calc.CodeAgence, narratifPdv)
            AjouterLigneSiNonNul(dt, comptes.CommissionEnvoiBanque,
                                 libelles(NatureMouvementWU.CommissionEnvoiBanque),
                                 0D, calc.CommissionEnvoiBanque, calc.CodeAgence, narratifPdv)

            ' 4) Commissions part Sous-agent (uniquement pour les SA disposant d'un CompteCommission).
            If String.Equals(calc.TypePdv, "SA", StringComparison.OrdinalIgnoreCase) AndAlso
               Not String.IsNullOrWhiteSpace(calc.CompteCommission) Then

                AjouterLigneSiNonNul(dt, calc.CompteCommission,
                                     libelles(NatureMouvementWU.CommissionTransfertSousAgent),
                                     0D, calc.CommissionTransfertSA, calc.CodeAgence, narratifPdv)
                AjouterLigneSiNonNul(dt, calc.CompteCommission,
                                     libelles(NatureMouvementWU.CommissionPaiementSousAgent),
                                     0D, calc.CommissionPaiementSA, calc.CodeAgence, narratifPdv)
                AjouterLigneSiNonNul(dt, calc.CompteCommission,
                                     libelles(NatureMouvementWU.CommissionEnvoiSousAgent),
                                     0D, calc.CommissionEnvoiSA, calc.CodeAgence, narratifPdv)
            End If

            ' 5) Taxes (impôts, TVA, TTA) : toujours créditées, à la charge de la banque.
            AjouterLigneSiNonNul(dt, comptes.ImpotsTaxeEnvoi, libelles(NatureMouvementWU.ImpotsTaxeEnvoi),
                                 0D, calc.TaxeEnvoi, calc.CodeAgence, narratifPdv)
            AjouterLigneSiNonNul(dt, comptes.TVACollectee, libelles(NatureMouvementWU.TVA),
                                 0D, calc.TVA, calc.CodeAgence, narratifPdv)
            AjouterLigneSiNonNul(dt, comptes.TTAEnvoi, libelles(NatureMouvementWU.TTAEnvoi),
                                 0D, calc.TTAEnvoi, calc.CodeAgence, narratifPdv)
            AjouterLigneSiNonNul(dt, comptes.TTAReception, libelles(NatureMouvementWU.TTAReception),
                                 0D, calc.TTAReception, calc.CodeAgence, narratifPdv)
        Next

        Return dt
    End Function

    ''' <summary>
    ''' Calcule — et mémorise sur le CalculWU — l'écart d'arrondi que cet Account apporte à la
    ''' pièce comptable : montant arrondi de sa ligne de mouvement MOINS la somme des montants
    ''' arrondis de ses contreparties (compte courant WU, commissions banque et sous-agent,
    ''' taxes). En valeurs exactes la différence est nulle par construction ; seul l'arrondi
    ''' FCFA ligne par ligne la rend non nulle (typiquement 0 ou ±1 FCFA par Account).
    '''
    ''' Par construction, la somme de ces écarts sur l'ensemble des Accounts est exactement
    ''' égale à l'écart global Débit − Crédit de la pièce, celui que VerifierEquilibrePiece
    ''' affecte au compte d'attente (section 14). Cela permet de tracer, ligne par ligne, d'où
    ''' provient l'écart global.
    '''
    ''' IMPORTANT : les arrondis reproduits ici doivent rester strictement alignés sur ceux
    ''' réellement posés par GenererPieceComptable (y compris la condition sur CompteCommission
    ''' pour les commissions sous-agent). Toute modification de l'une doit être répercutée ici.
    ''' </summary>
    ''' <returns>L'écart d'arrondi en FCFA (positif, négatif ou nul).</returns>
    Private Shared Function CalculerEcartArrondi(calc As CalculWU) As Long

        If calc Is Nothing Then Return 0L

        ' Un Account non comptabilisé n'apporte aucune ligne à la pièce : il ne peut donc pas
        ' en apporter l'écart d'arrondi. Afficher un écart pour lui laisserait croire qu'il
        ' pèse sur l'équilibre global alors qu'il n'y figure pas.
        If Not calc.EstComptabilisable Then
            calc.EcartArrondi = 0L
            Return 0L
        End If

        Dim totalCommissionsEtTaxes As Decimal =
            calc.CommissionTransfertBanque + calc.CommissionEnvoiBanque + calc.CommissionPaiementBanque +
            calc.CommissionTransfertSA + calc.CommissionPaiementSA + calc.CommissionEnvoiSA +
            calc.TaxeEnvoi + calc.TVA + calc.TTAEnvoi + calc.TTAReception

        Dim netMouvement As Decimal = calc.NetMouvement
        Dim netCompteCourant As Decimal = netMouvement - totalCommissionsEtTaxes

        ' Commissions sous-agent : postées uniquement si le CompteCommission est renseigné
        ' (même condition que dans GenererPieceComptable).
        Dim commissionsSA As Long = 0L
        If String.Equals(calc.TypePdv, "SA", StringComparison.OrdinalIgnoreCase) AndAlso
           Not String.IsNullOrWhiteSpace(calc.CompteCommission) Then
            commissionsSA = WUCalculationService.ArrondiFCFA(calc.CommissionTransfertSA) +
                            WUCalculationService.ArrondiFCFA(calc.CommissionPaiementSA) +
                            WUCalculationService.ArrondiFCFA(calc.CommissionEnvoiSA)
        End If

        Dim contrepartiesArrondies As Long =
            WUCalculationService.ArrondiFCFA(netCompteCourant) +
            WUCalculationService.ArrondiFCFA(calc.CommissionTransfertBanque) +
            WUCalculationService.ArrondiFCFA(calc.CommissionEnvoiBanque) +
            WUCalculationService.ArrondiFCFA(calc.CommissionPaiementBanque) +
            commissionsSA +
            WUCalculationService.ArrondiFCFA(calc.TaxeEnvoi) +
            WUCalculationService.ArrondiFCFA(calc.TVA) +
            WUCalculationService.ArrondiFCFA(calc.TTAEnvoi) +
            WUCalculationService.ArrondiFCFA(calc.TTAReception)

        calc.EcartArrondi = WUCalculationService.ArrondiFCFA(netMouvement) - contrepartiesArrondies
        Return calc.EcartArrondi
    End Function

    ''' <summary>
    ''' Dit la période couverte, sous la forme qu'un comptable lit sans hésiter.
    '''
    '''     une seule journée ........ DU 09 09 2026
    '''     un même mois ............. DU 08 AU 14 09 2026
    '''     à cheval sur deux mois ... DU 28/09/2026 AU 04/10/2026
    '''
    ''' UNE PÉRIODE ABSENTE NE REND RIEN, et surtout pas la date du jour. Un narratif
    ''' portant une date fausse est pire que muet : il désigne une semaine qui n'existe pas,
    ''' et rien dans la pièce ne permettrait de le redresser ensuite.
    '''
    ''' Un intervalle donné à l'envers est REMIS DANS L'ORDRE plutôt que refusé. La pièce
    ''' doit sortir, et « DU 14 AU 08 » ne serait relevé par personne avant l'inspection.
    ''' </summary>
    Public Shared Function SuffixeDePeriode(debut As Date?, fin As Date?) As String

        If Not debut.HasValue AndAlso Not fin.HasValue Then Return String.Empty

        Dim premier As Date = If(debut.HasValue, debut.Value.Date, fin.Value.Date)
        Dim dernier As Date = If(fin.HasValue, fin.Value.Date, premier)

        If dernier < premier Then
            Dim echange As Date = premier
            premier = dernier
            dernier = echange
        End If

        If premier = dernier Then
            Return String.Format(CultureInfo.InvariantCulture, ConstantesWU.PIECE_JOURNEE_FORMAT,
                                 premier.Day, premier.Month, premier.Year)
        End If

        If premier.Year = dernier.Year AndAlso premier.Month = dernier.Month Then
            Return String.Format(CultureInfo.InvariantCulture, ConstantesWU.PIECE_PERIODE_FORMAT,
                                 premier.Day, dernier.Day, dernier.Month, dernier.Year)
        End If

        Return String.Format(CultureInfo.InvariantCulture, ConstantesWU.PIECE_PERIODE_LONGUE_FORMAT,
                             premier, dernier)
    End Function

    ''' <summary>
    ''' Ajoute une ligne unique dont le sens (Débit/Crédit) est déterminé automatiquement par
    ''' le signe du montant : Débit si positif, Crédit si négatif. Rien n'est ajouté si le
    ''' montant arrondi est nul.
    ''' </summary>
    ''' <param name="narratif">
    ''' Ce que le FICHIER CORE BANKING lira dans ADDLTEXT, quand ce n'est pas le libellé. Vide
    ''' pour une ligne qui n'appartient à aucun point de vente — l'écart d'arrondi — et le
    ''' fichier retombe alors sur le libellé.
    ''' </param>
    Private Shared Sub AjouterLigneSigneAuto(dt As DataTable, compte As String, libelle As String,
                                             montant As Decimal, codeAgence As String,
                                             Optional narratif As String = "")
        Dim montantArrondi As Long = WUCalculationService.ArrondiFCFA(montant)
        If montantArrondi = 0L Then Return

        If montantArrondi > 0L Then
            AjouterLigne(dt, compte, libelle, montantArrondi, 0L, codeAgence, narratif)
        Else
            AjouterLigne(dt, compte, libelle, 0L, -montantArrondi, codeAgence, narratif)
        End If
    End Sub

    ''' <summary>Ajoute une ligne simple (Debit ou Credit) si le montant n'est pas nul une fois arrondi.</summary>
    Private Shared Sub AjouterLigneSiNonNul(dt As DataTable, compte As String, libelle As String,
                                            debit As Decimal, credit As Decimal, codeAgence As String,
                                            Optional narratif As String = "")
        Dim debitArrondi As Long = WUCalculationService.ArrondiFCFA(debit)
        Dim creditArrondi As Long = WUCalculationService.ArrondiFCFA(credit)
        If debitArrondi = 0L AndAlso creditArrondi = 0L Then Return
        AjouterLigne(dt, compte, libelle, debitArrondi, creditArrondi, codeAgence, narratif)
    End Sub

    Private Shared Sub AjouterLigne(dt As DataTable, compte As String, libelle As String,
                                    debit As Long, credit As Long, codeAgence As String,
                                    Optional narratif As String = "")
        Dim ligne As DataRow = dt.NewRow()
        ligne("Compte") = If(String.IsNullOrWhiteSpace(compte), ComptesSystemeWU.Actuels.CompteInterBancaire, compte)
        ligne("Libelle") = libelle
        ligne("Debit") = debit
        ligne("Credit") = credit

        ' La colonne peut manquer sur une pièce construite avant l'ajout du code agence : on ne
        ' la renseigne que si elle existe, plutôt que de faire échouer la génération.
        If dt.Columns.Contains("CodeAgence") Then
            ligne("CodeAgence") = If(codeAgence, String.Empty).Trim()
        End If

        ' MÊME PRUDENCE POUR LA NARRATIVE, et pour la même raison : une pièce de change ou une
        ' table construite par du code plus ancien peut ne pas porter la colonne.
        If dt.Columns.Contains("Narratif") Then
            ligne("Narratif") = If(narratif, String.Empty).Trim()
        End If

        dt.Rows.Add(ligne)
    End Sub

#End Region

#Region "Contrôle d'équilibre global (section 14)"

    ''' <summary>
    ''' Vérifie l'équilibre global de la pièce comptable (jamais Account par Account) et applique
    ''' la règle du compte d'attente :
    '''   - écart = 0 : pièce équilibrée, rien à faire ;
    '''   - 0 &lt; écart &lt;= 1000 : écart ajouté au CRÉDIT du compte d'attente ;
    '''   - -1000 &lt;= écart &lt; 0 : |écart| ajouté au DÉBIT du compte d'attente ;
    '''   - |écart| &gt; 1000 : anomalie, génération bloquée (retourne False).
    ''' </summary>
    ''' <param name="dtPiece">Pièce comptable générée par GenererPieceComptable (modifiée en place si un écart tolérable est absorbé).</param>
    ''' <param name="messageControle">Message explicite décrivant le résultat du contrôle.</param>
    ''' <param name="ecartAbsorbe">
    ''' L'écart effectivement POSÉ dans la pièce, signé : positif au crédit du compte d'attente,
    ''' négatif à son débit, zéro si la pièce s'équilibrait d'elle-même. C'est exactement
    ''' Débit − Crédit de la pièce avant absorption.
    '''
    ''' POURQUOI CETTE FONCTION LE DIT, AU LIEU DE LAISSER L'APPELANT LE RETROUVER. L'écran de
    ''' traitement en a besoin pour les deux chiffres du bordereau, et il les tirait jusqu'ici
    ''' de la pièce elle-même, EN RECONNAISSANT LA LIGNE À SON LIBELLÉ. C'était la dernière
    ''' comparaison de texte du projet, et elle interdisait de rendre ce libellé paramétrable :
    ''' la banque aurait changé un mot, la comparaison aurait échoué, et l'écart aurait disparu
    ''' du bordereau sans qu'aucun message ne le dise. Celui qui POSE la ligne est celui qui
    ''' sait ce qu'il a posé ; il le dit, et plus personne ne le devine.
    ''' </param>
    ''' <param name="compteEcart">
    ''' Le compte qui a absorbé l'écart, ou une chaîne vide si rien n'a été posé. C'est le
    ''' compte inter bancaire de SystemeWU, mais c'est à la pièce de le dire et non à
    ''' l'appelant de le relire.
    ''' </param>
    ''' <returns>True si la pièce est utilisable (équilibrée ou écart absorbé), False si l'anomalie bloque la génération.</returns>
    Public Shared Function VerifierEquilibrePiece(dtPiece As DataTable, ByRef messageControle As String,
                                                  ByRef ecartAbsorbe As Long,
                                                  ByRef compteEcart As String,
                                                  Optional debutPeriode As Date? = Nothing,
                                                  Optional finPeriode As Date? = Nothing) As Boolean

        ' Posés AVANT tout contrôle : une pièce vide, ou un écart hors seuil, ne doit pas
        ' laisser l'appelant avec les valeurs d'une génération précédente.
        ecartAbsorbe = 0L
        compteEcart = String.Empty

        If dtPiece Is Nothing OrElse dtPiece.Rows.Count = 0 Then
            messageControle = "La pièce comptable est vide : aucun contrôle d'équilibre possible."
            Return False
        End If

        Dim totalDebit As Decimal = dtPiece.AsEnumerable().Sum(Function(r) Convert.ToDecimal(r("Debit"), CultureInfo.InvariantCulture))
        Dim totalCredit As Decimal = dtPiece.AsEnumerable().Sum(Function(r) Convert.ToDecimal(r("Credit"), CultureInfo.InvariantCulture))
        Dim differenceGlobale As Decimal = totalDebit - totalCredit

        If differenceGlobale = 0D Then
            messageControle = $"Pièce comptable équilibrée (Débit = Crédit = {totalDebit:N0} FCFA)."
            Return True
        End If

        ' Compte encaissant l'écart d'arrondi résiduel : le compte inter bancaire paramétré
        ' dans SystemeWU (colonnes Cpte_attenteDEBIT / Cpte_attenteCREDIT).
        Dim compteDAttente As String = ComptesSystemeWU.Actuels.CompteInterBancaire

        ' SON LIBELLÉ VIENT DU PARAMÉTRAGE, COMME LES DOUZE AUTRES — mais de SA NATURE, et
        ' dans les DEUX modes. Cette ligne est posée après la pièce et ne se rattache à aucun
        ' point de vente : le modèle global la réduirait à « LD WU ACTIVITE ». La banque peut
        ' l'éditer, dans l'onglet des libellés par nature.
        '
        ' LA PÉRIODE LUI EST PASSÉE, bien qu'aucun libellé ne la porte par défaut : le repère
        ' {PERIODE} reste OFFERT à la banque dans l'écran de paramétrage, et un repère qu'on
        ' propose mais que la pièce rendrait vide serait un piège silencieux. Les trois autres
        ' sont vides, et c'est exact : il n'y a ici ni désignation, ni Account, ni code agence.
        '
        ' PÉRIODE ABSENTE : l'appelant qui ne la fournit pas obtient le libellé sans elle,
        ' plutôt qu'une pièce qui ne sort pas. C'est le cas des deux écrans qui n'appellent
        ' pas cette fonction pour poser un écart, mais pour contrôler.
        Dim libelleEcart As String = NarrativeRepository.EnVigueur().Libelle(
            NatureMouvementWU.EcartArrondi, String.Empty, String.Empty, String.Empty,
            SuffixeDePeriode(debutPeriode, finPeriode))

        If differenceGlobale > 0D AndAlso differenceGlobale <= ConstantesWU.SEUIL_ECART_TOLERE Then
            ' AUCUNE NARRATIVE DE POINT DE VENTE ICI, et c'est voulu : cette ligne
            ' n'appartient à aucun point de vente. Le fichier core banking retombe alors sur
            ' son libellé, qui dit déjà ce qu'elle est.
            AjouterLigne(dtPiece, compteDAttente, libelleEcart, 0L,
                         CLng(differenceGlobale), ConstantesWU.CB_AGENCE_SIEGE)

            ecartAbsorbe = CLng(differenceGlobale)
            compteEcart = compteDAttente

            messageControle = $"Écart de {differenceGlobale:N0} FCFA affecté au CRÉDIT du compte inter bancaire {compteDAttente}."
            Return True
        End If

        If differenceGlobale < 0D AndAlso differenceGlobale >= -ConstantesWU.SEUIL_ECART_TOLERE Then
            AjouterLigne(dtPiece, compteDAttente, libelleEcart,
                         CLng(Math.Abs(differenceGlobale)), 0L, ConstantesWU.CB_AGENCE_SIEGE)

            ' SIGNÉ, et négatif ici : un écart au débit et un écart au crédit ne se compensent
            ' pas dans la tête de celui qui relit le bordereau.
            ecartAbsorbe = CLng(differenceGlobale)
            compteEcart = compteDAttente

            messageControle = $"Écart de {Math.Abs(differenceGlobale):N0} FCFA affecté au DÉBIT du compte inter bancaire {compteDAttente}."
            Return True
        End If

        messageControle = $"ANOMALIE : l'écart global de {differenceGlobale:N0} FCFA dépasse le seuil toléré de " &
                           $"{ConstantesWU.SEUIL_ECART_TOLERE:N0} FCFA. Génération de la pièce bloquée."
        Return False
    End Function

#End Region

#Region "Export Excel (optionnel, section 1)"

    ''' <summary>
    ''' Nom de fichier proposé pour l'export de la pièce comptable d'une journée, sur le modèle
    ''' du fichier destiné au core banking : PieceWU_aaaammjj.xlsx. La journée figure dans le
    ''' nom pour qu'un dossier d'exports reste lisible sans ouvrir les classeurs.
    ''' </summary>
    Public Shared Function NomDeFichier(dateActivite As Date) As String
        Return $"PieceWU_{dateActivite:yyyyMMdd}.xlsx"
    End Function

    ''' <summary>
    ''' Exporte la pièce comptable à l'emplacement choisi par l'utilisateur PUIS laisse le
    ''' classeur ouvert sous ses yeux, comme le fait le fichier destiné au core banking :
    ''' un export qu'on ne voit pas est un export dont on doute.
    ''' </summary>
    ''' <returns>Le chemin du fichier produit.</returns>
    Public Shared Function ExporterEtOuvrirPieceExcel(dtPiece As DataTable, cheminFichier As String) As String
        Return EcrireClasseurPiece(dtPiece, cheminFichier)
    End Function

    ''' <summary>
    ''' Exporte la pièce DANS LE FORMULAIRE DE LA BANQUE, avec une feuille par point de vente
    ''' en plus de la pièce globale.
    '''
    ''' C'est la forme attendue par la comptabilité : le classeur à quatre colonnes que produit
    ''' la surcharge ci-dessus reste disponible pour un contrôle rapide, mais ce n'est pas un
    ''' document qui se vise et se signe.
    ''' </summary>
    ''' <param name="dtPiece">Pièce comptable de la journée.</param>
    ''' <param name="listeCalculs">Points de vente à détailler, un onglet chacun. Nothing pour n'écrire que la première feuille.</param>
    ''' <param name="dateActivite">Journée comptabilisée — celle que portera l'en-tête.</param>
    ''' <param name="cheminFichier">Chemin complet du .xlsx.</param>
    ''' <param name="nomPremiereFeuille">Onglet de la première feuille.</param>
    ''' <param name="intitulePremiereFeuille">Ligne d'identification de la première feuille.</param>
    ''' <param name="pointDeVente">
    ''' Le point de vente que porte la PREMIÈRE feuille, quand elle n'en porte qu'un. Sa RAISON
    ''' le nomme alors, comme les douze lignes qu'elle résume. Nothing pour une pièce globale,
    ''' qui les rassemble tous et n'en nomme aucun.
    ''' </param>
    ''' <param name="derniereJournee">
    ''' Dernière journée couverte, quand le rapport en portait plusieurs. La pièce s'intitule
    ''' alors « activité du X au Y » au lieu d'annoncer une seule journée pour une semaine.
    ''' </param>
    ''' <returns>Le chemin du fichier produit.</returns>
    Public Shared Function ExporterEtOuvrirPieceExcel(dtPiece As DataTable,
                                                      listeCalculs As IEnumerable(Of CalculWU),
                                                      dateActivite As Date,
                                                      cheminFichier As String,
                                                      Optional nomPremiereFeuille As String = "PIECE GLOBALE",
                                                      Optional intitulePremiereFeuille As String = "",
                                                      Optional agencePremiereFeuille As String = "",
                                                      Optional progression As ProgressionWU = Nothing,
                                                      Optional derniereJournee As Date? = Nothing,
                                                      Optional pointDeVente As CalculWU = Nothing) As String

        Return PieceExcelWU.Ecrire(dtPiece, listeCalculs, dateActivite, cheminFichier,
                                   nomPremiereFeuille, intitulePremiereFeuille, agencePremiereFeuille,
                                   progression, derniereJournee, pointDeVente)
    End Function

    ''' <summary>
    ''' Ouvre la pièce comptable dans Excel sans rien demander : le classeur est écrit dans un
    ''' fichier temporaire, puis laissé ouvert pour consultation, impression ou enregistrement
    ''' manuel immédiat.
    '''
    ''' L'écran de traitement présente désormais la pièce dans sa propre fenêtre avant tout
    ''' export, et passe par ExporterEtOuvrirPieceExcel. Cette méthode-ci reste en service pour
    ''' l'ouverture immédiate, sans choix d'emplacement — et parce qu'une version antérieure de
    ''' FrmCompensationWU l'appelle : la retirer casserait la compilation d'un poste dont tous
    ''' les fichiers ne seraient pas encore à jour.
    ''' </summary>
    ''' <param name="dtPiece">Pièce comptable à ouvrir.</param>
    ''' <returns>Chemin du fichier temporaire dans lequel le classeur a été sauvegardé.</returns>
    Public Shared Function OuvrirPieceComptableExcel(dtPiece As DataTable) As String

        Dim cheminTemp As String = IO.Path.Combine(IO.Path.GetTempPath(),
                                                   $"PieceWU_{Date.Now:yyyyMMdd_HHmmss}.xlsx")
        Return EcrireClasseurPiece(dtPiece, cheminTemp)
    End Function

    ''' <summary>
    ''' Écrit la pièce comptable dans un classeur Excel, par liaison tardive (voir la remarque
    ''' sur Option Strict Off en tête de fichier), et laisse le classeur OUVERT sous les yeux
    ''' de l'utilisateur.
    '''
    ''' Excel n'est jamais quitté en cas de succès : c'est tout l'intérêt de la méthode. Seule
    ''' une erreur referme ce qui a pu être ouvert, pour ne pas laisser un classeur incomplet
    ''' à l'écran.
    ''' </summary>
    ''' <param name="dtPiece">Pièce comptable à écrire.</param>
    ''' <param name="cheminFichier">Chemin complet du fichier .xlsx.</param>
    ''' <returns>Le chemin du fichier écrit.</returns>
    Private Shared Function EcrireClasseurPiece(dtPiece As DataTable, cheminFichier As String) As String

        If dtPiece Is Nothing OrElse dtPiece.Rows.Count = 0 Then
            Throw New InvalidOperationException("Aucune donnée à exporter : générez la pièce comptable au préalable.")
        End If

        If String.IsNullOrWhiteSpace(cheminFichier) Then
            Throw New ArgumentException("Chemin de fichier non renseigné.", NameOf(cheminFichier))
        End If

        Dim excelApp As Object = Nothing
        Dim classeur As Object = Nothing
        Dim feuille As Object = Nothing

        Try
            Dim typeExcel As Type = Type.GetTypeFromProgID("Excel.Application")
            If typeExcel Is Nothing Then
                Throw New InvalidOperationException("Microsoft Excel n'est pas installé sur ce poste.")
            End If

            excelApp = Activator.CreateInstance(typeExcel)
            excelApp.Visible = True
            excelApp.DisplayAlerts = False

            classeur = excelApp.Workbooks.Add()
            feuille = classeur.Worksheets(1)
            RemplirFeuillePiece(feuille, dtPiece)
            classeur.SaveAs(cheminFichier)

            ' Mise au premier plan de la fenêtre Excel : purement cosmétique, et volontairement
            ' protégée par un Try/Catch silencieux. L'objet Application n'expose pas Activate()
            ' (ce membre appartient à Workbook / Window) et Windows peut refuser à un processus
            ' d'arrière-plan de prendre le focus : dans les deux cas la pièce est déjà ouverte,
            ' l'échec ne doit donc pas remonter comme une erreur à l'utilisateur.
            Try
                excelApp.WindowState = -4137 ' xlMaximized
                classeur.Activate()
            Catch
                ' Ignoré volontairement : Excel est ouvert, seule la mise au premier plan a échoué.
            End Try

            Return cheminFichier

        Catch
            ' En cas d'échec, Excel ne doit pas rester ouvert sur un classeur incomplet : on
            ' referme ce qui a pu être ouvert avant de relancer l'erreur à l'appelant.
            Try
                If classeur IsNot Nothing Then classeur.Close(False)
            Catch
            End Try
            Try
                If excelApp IsNot Nothing Then excelApp.Quit()
            Catch
            End Try
            Throw

        Finally
            ' Seules les références COM intermédiaires sont libérées : excelApp reste actif tant
            ' que sa fenêtre est affichée à l'écran.
            If feuille IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(feuille)
            If classeur IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(classeur)
        End Try
    End Function

    ''' <summary>Écrit les en-têtes et les lignes de dtPiece dans une feuille Excel (late binding).</summary>
    Private Shared Sub RemplirFeuillePiece(feuille As Object, dtPiece As DataTable)
        feuille.Name = "Piece WU"

        feuille.Cells(1, 1).Value = "Compte"
        feuille.Cells(1, 2).Value = "Libelle"
        feuille.Cells(1, 3).Value = "Debit"
        feuille.Cells(1, 4).Value = "Credit"
        feuille.Range("A1:D1").Font.Bold = True

        Dim ligneExcel As Integer = 2
        For Each ligne As DataRow In dtPiece.Rows
            feuille.Cells(ligneExcel, 1).Value = Convert.ToString(ligne("Compte"))
            feuille.Cells(ligneExcel, 2).Value = Convert.ToString(ligne("Libelle"))
            feuille.Cells(ligneExcel, 3).Value = Convert.ToInt64(ligne("Debit"))
            feuille.Cells(ligneExcel, 4).Value = Convert.ToInt64(ligne("Credit"))
            ligneExcel += 1
        Next

        feuille.Columns.AutoFit()
    End Sub

#End Region

End Class
