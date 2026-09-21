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
    '''     courant WU pour une agence propre — celle-ci n'a pas de compte de compensation
    '''     propre dans les livres de la banque (confirmé par la banque). Pour le montant net :
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
    Public Shared Function GenererPieceComptable(listeCalculs As IEnumerable(Of CalculWU)) As DataTable

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

        If listeCalculs Is Nothing Then Return dt

        ' Comptes comptables en service, lus une fois pour toute la pièce : ils proviennent de la
        ' table SystemeWU (voir ComptesSystemeWU) et non plus de constantes figées dans le code.
        ' Les prendre ici, et non à chaque écriture, garantit qu'une même pièce ne mélange jamais
        ' deux paramétrages si les comptes venaient à être modifiés pendant sa génération.
        Dim comptes As ComptesSystemeWU = ComptesSystemeWU.Actuels

        For Each calc As CalculWU In listeCalculs

            ' Un Account dont la banque ne connaît pas les comptes n'est pas comptabilisé : il
            ' est purement et simplement absent de la pièce (voir CalculWU.EstComptabilisable).
            ' Sa ligne de mouvement irait sinon sur le compte courant WU, c'est-à-dire sur un
            ' compte qui n'est pas le sien, et la correction se ferait à la main après coup.
            If Not calc.EstComptabilisable Then
                calc.EcartArrondi = 0L
                Continue For
            End If

            Dim compteMouvement As String
            If String.Equals(calc.TypePdv, "SA", StringComparison.OrdinalIgnoreCase) AndAlso
               Not String.IsNullOrWhiteSpace(calc.CompteCompense) Then
                compteMouvement = calc.CompteCompense
            Else
                compteMouvement = comptes.CompteCourant
            End If

            Dim libelleMouvement As String = String.Format(ConstantesWU.LIB_MOUVEMENT_ACTIVITE_FORMAT, calc.Designation).Trim()

            Dim totalCommissionsEtTaxes As Decimal =
                calc.CommissionTransfertBanque + calc.CommissionEnvoiBanque + calc.CommissionPaiementBanque +
                calc.CommissionTransfertSA + calc.CommissionPaiementSA + calc.CommissionEnvoiSA +
                calc.TaxeEnvoi + calc.TVA + calc.TTAEnvoi + calc.TTAReception

            Dim netMouvement As Decimal = calc.NetMouvement
            Dim netCompteCourant As Decimal = netMouvement - totalCommissionsEtTaxes

            ' Mémorise l'écart d'arrondi apporté par cet Account (visible dans la grille de contrôle).
            CalculerEcartArrondi(calc)

            ' 1) Ligne de mouvement (compte de compensation du point de vente).
            AjouterLigneSigneAuto(dt, compteMouvement, libelleMouvement, netMouvement, calc.CodeAgence)

            ' 2) Contrepartie sur le compte courant WU (part nette revenant à la banque).
            AjouterLigneSigneAuto(dt, comptes.CompteCourant, ConstantesWU.LIB_COMPTE_COURANT, -netCompteCourant, calc.CodeAgence)

            ' 3) Commissions part Banque (toujours créditées, quel que soit le type de PDV).
            ' Transfert et Envoi partagent le même compte dans le paramétrage actuel (728300148),
            ' mais la table SystemeWU les porte dans deux colonnes distinctes (Cpte_Produit et
            ' Cpte_Produit_Envoi) : ils sont donc désormais dissociables sans toucher au code.
            AjouterLigneSiNonNul(dt, comptes.CommissionTransfertBanque,
                                 ConstantesWU.LIB_COMMISSION_TRANSFERT_BANQUE, 0D, calc.CommissionTransfertBanque, calc.CodeAgence)
            AjouterLigneSiNonNul(dt, comptes.CommissionEnvoiBanque,
                                 ConstantesWU.LIB_COMMISSION_ENVOI_BANQUE, 0D, calc.CommissionEnvoiBanque, calc.CodeAgence)
            AjouterLigneSiNonNul(dt, comptes.CommissionPaiementBanque,
                                 ConstantesWU.LIB_COMMISSION_PAIEMENT_BANQUE, 0D, calc.CommissionPaiementBanque, calc.CodeAgence)

            ' 4) Commissions part Sous-agent (uniquement pour les SA disposant d'un CompteCommission).
            If String.Equals(calc.TypePdv, "SA", StringComparison.OrdinalIgnoreCase) AndAlso
               Not String.IsNullOrWhiteSpace(calc.CompteCommission) Then

                AjouterLigneSiNonNul(dt, calc.CompteCommission,
                                     $"{ConstantesWU.LIB_COMMISSION_TRANSFERT_SA} {calc.Designation}".Trim(), 0D, calc.CommissionTransfertSA, calc.CodeAgence)
                AjouterLigneSiNonNul(dt, calc.CompteCommission,
                                     $"{ConstantesWU.LIB_COMMISSION_PAIEMENT_SA} {calc.Designation}".Trim(), 0D, calc.CommissionPaiementSA, calc.CodeAgence)
                AjouterLigneSiNonNul(dt, calc.CompteCommission,
                                     $"{ConstantesWU.LIB_COMMISSION_ENVOI_SA} {calc.Designation}".Trim(), 0D, calc.CommissionEnvoiSA, calc.CodeAgence)
            End If

            ' 5) Taxes (impôts, TVA, TTA) : toujours créditées, à la charge de la banque.
            AjouterLigneSiNonNul(dt, comptes.ImpotsTaxeEnvoi, ConstantesWU.LIB_IMPOTS_TAXE_ENVOI, 0D, calc.TaxeEnvoi, calc.CodeAgence)
            AjouterLigneSiNonNul(dt, comptes.TVACollectee, ConstantesWU.LIB_TVA, 0D, calc.TVA, calc.CodeAgence)
            AjouterLigneSiNonNul(dt, comptes.TTAEnvoi, ConstantesWU.LIB_TTA_ENVOI, 0D, calc.TTAEnvoi, calc.CodeAgence)
            AjouterLigneSiNonNul(dt, comptes.TTAReception, ConstantesWU.LIB_TTA_RECEPTION, 0D, calc.TTAReception, calc.CodeAgence)
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
    ''' Ajoute une ligne unique dont le sens (Débit/Crédit) est déterminé automatiquement par
    ''' le signe du montant : Débit si positif, Crédit si négatif. Rien n'est ajouté si le
    ''' montant arrondi est nul.
    ''' </summary>
    Private Shared Sub AjouterLigneSigneAuto(dt As DataTable, compte As String, libelle As String,
                                             montant As Decimal, codeAgence As String)
        Dim montantArrondi As Long = WUCalculationService.ArrondiFCFA(montant)
        If montantArrondi = 0L Then Return

        If montantArrondi > 0L Then
            AjouterLigne(dt, compte, libelle, montantArrondi, 0L, codeAgence)
        Else
            AjouterLigne(dt, compte, libelle, 0L, -montantArrondi, codeAgence)
        End If
    End Sub

    ''' <summary>Ajoute une ligne simple (Debit ou Credit) si le montant n'est pas nul une fois arrondi.</summary>
    Private Shared Sub AjouterLigneSiNonNul(dt As DataTable, compte As String, libelle As String,
                                            debit As Decimal, credit As Decimal, codeAgence As String)
        Dim debitArrondi As Long = WUCalculationService.ArrondiFCFA(debit)
        Dim creditArrondi As Long = WUCalculationService.ArrondiFCFA(credit)
        If debitArrondi = 0L AndAlso creditArrondi = 0L Then Return
        AjouterLigne(dt, compte, libelle, debitArrondi, creditArrondi, codeAgence)
    End Sub

    Private Shared Sub AjouterLigne(dt As DataTable, compte As String, libelle As String,
                                    debit As Long, credit As Long, codeAgence As String)
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
    ''' <returns>True si la pièce est utilisable (équilibrée ou écart absorbé), False si l'anomalie bloque la génération.</returns>
    Public Shared Function VerifierEquilibrePiece(dtPiece As DataTable, ByRef messageControle As String) As Boolean

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
        Dim compteEcart As String = ComptesSystemeWU.Actuels.CompteInterBancaire

        If differenceGlobale > 0D AndAlso differenceGlobale <= ConstantesWU.SEUIL_ECART_TOLERE Then
            AjouterLigne(dtPiece, compteEcart, ConstantesWU.LIB_ECART_ATTENTE, 0L,
                         CLng(differenceGlobale), ConstantesWU.CB_AGENCE_SIEGE)
            messageControle = $"Écart de {differenceGlobale:N0} FCFA affecté au CRÉDIT du compte inter bancaire {compteEcart}."
            Return True
        End If

        If differenceGlobale < 0D AndAlso differenceGlobale >= -ConstantesWU.SEUIL_ECART_TOLERE Then
            AjouterLigne(dtPiece, compteEcart, ConstantesWU.LIB_ECART_ATTENTE,
                         CLng(Math.Abs(differenceGlobale)), 0L, ConstantesWU.CB_AGENCE_SIEGE)
            messageControle = $"Écart de {Math.Abs(differenceGlobale):N0} FCFA affecté au DÉBIT du compte inter bancaire {compteEcart}."
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
    ''' <returns>Le chemin du fichier produit.</returns>
    Public Shared Function ExporterEtOuvrirPieceExcel(dtPiece As DataTable,
                                                      listeCalculs As IEnumerable(Of CalculWU),
                                                      dateActivite As Date,
                                                      cheminFichier As String,
                                                      Optional nomPremiereFeuille As String = "PIECE GLOBALE",
                                                      Optional intitulePremiereFeuille As String = "",
                                                      Optional agencePremiereFeuille As String = "",
                                                      Optional progression As ProgressionWU = Nothing) As String

        Return PieceExcelWU.Ecrire(dtPiece, listeCalculs, dateActivite, cheminFichier,
                                   nomPremiereFeuille, intitulePremiereFeuille, agencePremiereFeuille,
                                   progression)
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
