' Option Strict est désactivé UNIQUEMENT dans ce fichier car ExporterPieceExcel utilise la
' liaison tardive (late binding) sur Microsoft Excel via Type.GetTypeFromProgID, afin de ne pas
' imposer de référence COM obligatoire au projet lorsque Excel n'est pas installé (section 1).
' Tout le reste du fichier (grille de contrôle, pièce comptable, équilibrage) reste fortement typé.
Option Strict Off
Option Explicit On

Imports System.Data
Imports System.Globalization
Imports System.IO
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
    '''     de compensation du point de vente (CompteCompense du sous-agent, ou le compte
    '''     courant WU pour une agence propre / Account non paramétré), pour le montant net :
    '''         NetMouvement = (PrincipalEnvoi + ChargeEnvoi + Taxes) − PrincipalPaye
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

        If listeCalculs Is Nothing Then Return dt

        For Each calc As CalculWU In listeCalculs

            Dim compteMouvement As String
            If String.Equals(calc.TypePdv, "SA", StringComparison.OrdinalIgnoreCase) AndAlso
               Not String.IsNullOrWhiteSpace(calc.CompteCompense) Then
                compteMouvement = calc.CompteCompense
            Else
                compteMouvement = ConstantesWU.CPT_COMPTE_COURANT
            End If

            Dim libelleMouvement As String = String.Format(ConstantesWU.LIB_MOUVEMENT_ACTIVITE_FORMAT, calc.Designation).Trim()

            Dim totalCommissionsEtTaxes As Decimal =
                calc.CommissionTransfertBanque + calc.CommissionEnvoiBanque + calc.CommissionPaiementBanque +
                calc.CommissionTransfertSA + calc.CommissionPaiementSA + calc.CommissionEnvoiSA +
                calc.TaxeEnvoi + calc.TVA + calc.TTAEnvoi + calc.TTAReception

            Dim netMouvement As Decimal = (calc.PrincipalEnvoi + calc.ChargeEnvoi + calc.Taxes) - calc.PrincipalPaye
            Dim netCompteCourant As Decimal = netMouvement - totalCommissionsEtTaxes

            ' Mémorise l'écart d'arrondi apporté par cet Account (visible dans la grille de contrôle).
            CalculerEcartArrondi(calc)

            ' 1) Ligne de mouvement (compte de compensation du point de vente).
            AjouterLigneSigneAuto(dt, compteMouvement, libelleMouvement, netMouvement)

            ' 2) Contrepartie sur le compte courant WU (part nette revenant à la banque).
            AjouterLigneSigneAuto(dt, ConstantesWU.CPT_COMPTE_COURANT, ConstantesWU.LIB_COMPTE_COURANT, -netCompteCourant)

            ' 3) Commissions part Banque (toujours créditées, quel que soit le type de PDV).
            AjouterLigneSiNonNul(dt, ConstantesWU.CPT_COMMISSION_TRANSFERT_ENVOI_BANQUE,
                                 ConstantesWU.LIB_COMMISSION_TRANSFERT_BANQUE, 0D, calc.CommissionTransfertBanque)
            AjouterLigneSiNonNul(dt, ConstantesWU.CPT_COMMISSION_TRANSFERT_ENVOI_BANQUE,
                                 ConstantesWU.LIB_COMMISSION_ENVOI_BANQUE, 0D, calc.CommissionEnvoiBanque)
            AjouterLigneSiNonNul(dt, ConstantesWU.CPT_COMMISSION_PAIEMENT_BANQUE,
                                 ConstantesWU.LIB_COMMISSION_PAIEMENT_BANQUE, 0D, calc.CommissionPaiementBanque)

            ' 4) Commissions part Sous-agent (uniquement pour les SA disposant d'un CompteCommission).
            If String.Equals(calc.TypePdv, "SA", StringComparison.OrdinalIgnoreCase) AndAlso
               Not String.IsNullOrWhiteSpace(calc.CompteCommission) Then

                AjouterLigneSiNonNul(dt, calc.CompteCommission,
                                     $"{ConstantesWU.LIB_COMMISSION_TRANSFERT_SA} {calc.Designation}".Trim(), 0D, calc.CommissionTransfertSA)
                AjouterLigneSiNonNul(dt, calc.CompteCommission,
                                     $"{ConstantesWU.LIB_COMMISSION_PAIEMENT_SA} {calc.Designation}".Trim(), 0D, calc.CommissionPaiementSA)
                AjouterLigneSiNonNul(dt, calc.CompteCommission,
                                     $"{ConstantesWU.LIB_COMMISSION_ENVOI_SA} {calc.Designation}".Trim(), 0D, calc.CommissionEnvoiSA)
            End If

            ' 5) Taxes (impôts, TVA, TTA) : toujours créditées, à la charge de la banque.
            AjouterLigneSiNonNul(dt, ConstantesWU.CPT_IMPOTS_TAXE_ENVOI, ConstantesWU.LIB_IMPOTS_TAXE_ENVOI, 0D, calc.TaxeEnvoi)
            AjouterLigneSiNonNul(dt, ConstantesWU.CPT_TVA_COLLECTEE, ConstantesWU.LIB_TVA, 0D, calc.TVA)
            AjouterLigneSiNonNul(dt, ConstantesWU.CPT_TTA_ENVOI, ConstantesWU.LIB_TTA_ENVOI, 0D, calc.TTAEnvoi)
            AjouterLigneSiNonNul(dt, ConstantesWU.CPT_TTA_RECEPTION, ConstantesWU.LIB_TTA_RECEPTION, 0D, calc.TTAReception)
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

        Dim totalCommissionsEtTaxes As Decimal =
            calc.CommissionTransfertBanque + calc.CommissionEnvoiBanque + calc.CommissionPaiementBanque +
            calc.CommissionTransfertSA + calc.CommissionPaiementSA + calc.CommissionEnvoiSA +
            calc.TaxeEnvoi + calc.TVA + calc.TTAEnvoi + calc.TTAReception

        Dim netMouvement As Decimal = (calc.PrincipalEnvoi + calc.ChargeEnvoi + calc.Taxes) - calc.PrincipalPaye
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
    Private Shared Sub AjouterLigneSigneAuto(dt As DataTable, compte As String, libelle As String, montant As Decimal)
        Dim montantArrondi As Long = WUCalculationService.ArrondiFCFA(montant)
        If montantArrondi = 0L Then Return

        If montantArrondi > 0L Then
            AjouterLigne(dt, compte, libelle, montantArrondi, 0L)
        Else
            AjouterLigne(dt, compte, libelle, 0L, -montantArrondi)
        End If
    End Sub

    ''' <summary>Ajoute une ligne simple (Debit ou Credit) si le montant n'est pas nul une fois arrondi.</summary>
    Private Shared Sub AjouterLigneSiNonNul(dt As DataTable, compte As String, libelle As String, debit As Decimal, credit As Decimal)
        Dim debitArrondi As Long = WUCalculationService.ArrondiFCFA(debit)
        Dim creditArrondi As Long = WUCalculationService.ArrondiFCFA(credit)
        If debitArrondi = 0L AndAlso creditArrondi = 0L Then Return
        AjouterLigne(dt, compte, libelle, debitArrondi, creditArrondi)
    End Sub

    Private Shared Sub AjouterLigne(dt As DataTable, compte As String, libelle As String, debit As Long, credit As Long)
        Dim ligne As DataRow = dt.NewRow()
        ligne("Compte") = If(String.IsNullOrWhiteSpace(compte), ConstantesWU.CPT_ATTENTE, compte)
        ligne("Libelle") = libelle
        ligne("Debit") = debit
        ligne("Credit") = credit
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

        If differenceGlobale > 0D AndAlso differenceGlobale <= ConstantesWU.SEUIL_ECART_TOLERE Then
            AjouterLigne(dtPiece, ConstantesWU.CPT_ATTENTE, ConstantesWU.LIB_ECART_ATTENTE, 0L, CLng(differenceGlobale))
            messageControle = $"Écart de {differenceGlobale:N0} FCFA affecté au CRÉDIT du compte d'attente {ConstantesWU.CPT_ATTENTE}."
            Return True
        End If

        If differenceGlobale < 0D AndAlso differenceGlobale >= -ConstantesWU.SEUIL_ECART_TOLERE Then
            AjouterLigne(dtPiece, ConstantesWU.CPT_ATTENTE, ConstantesWU.LIB_ECART_ATTENTE, CLng(Math.Abs(differenceGlobale)), 0L)
            messageControle = $"Écart de {Math.Abs(differenceGlobale):N0} FCFA affecté au DÉBIT du compte d'attente {ConstantesWU.CPT_ATTENTE}."
            Return True
        End If

        messageControle = $"ANOMALIE : l'écart global de {differenceGlobale:N0} FCFA dépasse le seuil toléré de " &
                           $"{ConstantesWU.SEUIL_ECART_TOLERE:N0} FCFA. Génération de la pièce bloquée."
        Return False
    End Function

#End Region

#Region "Export Excel (optionnel, section 1)"

    ''' <summary>
    ''' Exporte la pièce comptable vers un classeur Excel via Microsoft Excel Interop.
    ''' Nécessite l'ajout de la référence COM "Microsoft Excel XX.0 Object Library" au projet.
    ''' Isolée dans sa propre fonction : toute la logique métier fonctionne sans Excel installé.
    ''' </summary>
    ''' <param name="dtPiece">Pièce comptable à exporter.</param>
    ''' <param name="cheminFichier">Chemin complet du fichier .xlsx à générer.</param>
    Public Shared Sub ExporterPieceExcel(dtPiece As DataTable, cheminFichier As String)

        If dtPiece Is Nothing OrElse dtPiece.Rows.Count = 0 Then
            Throw New InvalidOperationException("Aucune donnée à exporter : générez la pièce comptable avant l'export.")
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
            excelApp.Visible = False
            excelApp.DisplayAlerts = False

            classeur = excelApp.Workbooks.Add()
            feuille = classeur.Worksheets(1)
            RemplirFeuillePiece(feuille, dtPiece)
            classeur.SaveAs(cheminFichier)

        Finally
            Try
                If classeur IsNot Nothing Then classeur.Close(False)
            Catch
            End Try
            Try
                If excelApp IsNot Nothing Then excelApp.Quit()
            Catch
            End Try

            If feuille IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(feuille)
            If classeur IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(classeur)
            If excelApp IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp)
        End Try
    End Sub

    ''' <summary>
    ''' Génère la pièce comptable et l'ouvre DIRECTEMENT dans Microsoft Excel (fenêtre visible),
    ''' sans passer par une boîte de dialogue d'enregistrement. Le classeur est sauvegardé dans
    ''' un fichier temporaire (pour avoir un nom et être persisté sur disque) puis laissé OUVERT
    ''' pour consultation/impression/enregistrement manuel immédiat par l'utilisateur.
    ''' Nécessite Microsoft Excel installé sur le poste ; utilise la liaison tardive comme
    ''' ExporterPieceExcel (voir remarque sur Option Strict Off en tête de fichier).
    ''' </summary>
    ''' <param name="dtPiece">Pièce comptable à ouvrir (générée par GenererPieceComptable).</param>
    ''' <returns>Chemin du fichier temporaire dans lequel le classeur a été sauvegardé.</returns>
    Public Shared Function OuvrirPieceComptableExcel(dtPiece As DataTable) As String

        If dtPiece Is Nothing OrElse dtPiece.Rows.Count = 0 Then
            Throw New InvalidOperationException("Aucune donnée à afficher : générez la pièce comptable avant de l'ouvrir.")
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

            Dim cheminTemp As String = Path.Combine(Path.GetTempPath(), $"PieceWU_{Date.Now:yyyyMMdd_HHmmss}.xlsx")
            classeur.SaveAs(cheminTemp)

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

            Return cheminTemp

        Catch
            ' En cas d'échec, on ferme proprement ce qui a pu être ouvert avant de relancer l'erreur
            ' (contrairement au cas nominal, ici Excel ne doit pas rester ouvert sur un classeur en échec).
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
            ' NOTE : contrairement à ExporterPieceExcel, on ne ferme PAS le classeur et on ne quitte
            ' PAS Excel ici en cas de succès : c'est tout l'intérêt de cette méthode, laisser la
            ' pièce ouverte et visible pour l'utilisateur. Seules les références COM intermédiaires
            ' (feuille, classeur) sont libérées ; excelApp reste actif tant que sa fenêtre est ouverte.
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
