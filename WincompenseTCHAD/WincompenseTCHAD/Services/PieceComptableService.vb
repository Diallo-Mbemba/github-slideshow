' Option Strict est désactivé UNIQUEMENT dans ce fichier car ExporterPieceExcel utilise la
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
    ''' des CalculWU, en reproduisant la logique du modèle PieceComptabilsationTchad.xlsx.
    '''
    ''' HYPOTHÈSE (à valider contre le classeur de référence, non disponible à la rédaction) :
    ''' pour chaque Account, le "compte de mouvement" (CompteMouvement) est :
    '''   - le CompteCompense du sous-agent si TypePdv = "SA" et qu'il est renseigné,
    '''   - le compte courant WU (32100003292) sinon (agence propre "EC" ou Account "INCONNU").
    ''' Ce compte reçoit :
    '''   - un DÉBIT du PrincipalPaye (somme décaissée pour les paiements),
    '''   - un CRÉDIT du PrincipalEnvoi + ChargeEnvoi + Taxes (somme encaissée pour les envois).
    ''' Toutes les commissions et taxes sont ensuite débitées sur ce même compte de mouvement,
    ''' en contrepartie créditée sur le compte de destination approprié (compte de commission
    ''' banque, compte de commission du sous-agent, ou compte de taxe). Cette construction en
    ''' partie double garantit que seule la différence entre PrincipalPaye et
    ''' (PrincipalEnvoi + ChargeEnvoi + Taxes), cumulée sur tous les Accounts, peut générer un
    ''' écart global — ce qui est cohérent avec le mécanisme du compte d'attente (section 14).
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

            Dim libelleCompte As String = $"{ConstantesWU.LIB_COMPTE_COURANT} - {calc.Account} {calc.Designation}".Trim()

            ' 1) Mouvement principal : décaissement des paiements / encaissement des envois.
            AjouterLigneSiNonNul(dt, compteMouvement, libelleCompte, calc.PrincipalPaye, 0D)
            AjouterLigneSiNonNul(dt, compteMouvement, libelleCompte, 0D, calc.PrincipalEnvoi + calc.ChargeEnvoi + calc.Taxes)

            ' 2) Commissions part Banque (toujours postées, quel que soit le type de PDV).
            AjouterEcriture(dt, compteMouvement, ConstantesWU.CPT_COMMISSION_TRANSFERT_ENVOI_BANQUE,
                             ConstantesWU.LIB_COMMISSION_TRANSFERT_BANQUE, calc.CommissionTransfertBanque)
            AjouterEcriture(dt, compteMouvement, ConstantesWU.CPT_COMMISSION_TRANSFERT_ENVOI_BANQUE,
                             ConstantesWU.LIB_COMMISSION_ENVOI_BANQUE, calc.CommissionEnvoiBanque)
            AjouterEcriture(dt, compteMouvement, ConstantesWU.CPT_COMMISSION_PAIEMENT_BANQUE,
                             ConstantesWU.LIB_COMMISSION_PAIEMENT_BANQUE, calc.CommissionPaiementBanque)

            ' 3) Commissions part Sous-agent (uniquement pour les SA disposant d'un CompteCommission).
            If String.Equals(calc.TypePdv, "SA", StringComparison.OrdinalIgnoreCase) AndAlso
               Not String.IsNullOrWhiteSpace(calc.CompteCommission) Then

                AjouterEcriture(dt, compteMouvement, calc.CompteCommission,
                                 ConstantesWU.LIB_COMMISSION_TRANSFERT_SA, calc.CommissionTransfertSA)
                AjouterEcriture(dt, compteMouvement, calc.CompteCommission,
                                 ConstantesWU.LIB_COMMISSION_PAIEMENT_SA, calc.CommissionPaiementSA)
                AjouterEcriture(dt, compteMouvement, calc.CompteCommission,
                                 ConstantesWU.LIB_COMMISSION_ENVOI_SA, calc.CommissionEnvoiSA)
            End If

            ' 4) Taxes (impôts, TVA, TTA) : toujours à la charge de la banque.
            AjouterEcriture(dt, compteMouvement, ConstantesWU.CPT_IMPOTS_TAXE_ENVOI,
                             ConstantesWU.LIB_IMPOTS_TAXE_ENVOI, calc.TaxeEnvoi)
            AjouterEcriture(dt, compteMouvement, ConstantesWU.CPT_TVA_COLLECTEE,
                             ConstantesWU.LIB_TVA, calc.TVA)
            AjouterEcriture(dt, compteMouvement, ConstantesWU.CPT_TTA_ENVOI,
                             ConstantesWU.LIB_TTA_ENVOI, calc.TTAEnvoi)
            AjouterEcriture(dt, compteMouvement, ConstantesWU.CPT_TTA_RECEPTION,
                             ConstantesWU.LIB_TTA_RECEPTION, calc.TTAReception)
        Next

        Return dt
    End Function

    ''' <summary>Ajoute une écriture en partie double (Debit compteDebit / Credit compteCredit) si le montant n'est pas nul une fois arrondi.</summary>
    Private Shared Sub AjouterEcriture(dt As DataTable, compteDebit As String, compteCredit As String, libelle As String, montant As Decimal)
        Dim montantArrondi As Long = WUCalculationService.ArrondiFCFA(montant)
        If montantArrondi = 0L Then Return

        AjouterLigne(dt, compteDebit, libelle, montantArrondi, 0L)
        AjouterLigne(dt, compteCredit, libelle, 0L, montantArrondi)
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
            feuille.Name = "Piece WU"

            ' En-têtes.
            feuille.Cells(1, 1).Value = "Compte"
            feuille.Cells(1, 2).Value = "Libelle"
            feuille.Cells(1, 3).Value = "Debit"
            feuille.Cells(1, 4).Value = "Credit"

            Dim ligneExcel As Integer = 2
            For Each ligne As DataRow In dtPiece.Rows
                feuille.Cells(ligneExcel, 1).Value = Convert.ToString(ligne("Compte"))
                feuille.Cells(ligneExcel, 2).Value = Convert.ToString(ligne("Libelle"))
                feuille.Cells(ligneExcel, 3).Value = Convert.ToInt64(ligne("Debit"))
                feuille.Cells(ligneExcel, 4).Value = Convert.ToInt64(ligne("Credit"))
                ligneExcel += 1
            Next

            feuille.Columns.AutoFit()
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

#End Region

End Class
