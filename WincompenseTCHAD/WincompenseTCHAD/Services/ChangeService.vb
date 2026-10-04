Option Strict On
Option Explicit On

Imports System.Data

''' <summary>
''' Calcule les GAINS ET PERTES DE CHANGE, transaction par transaction, à partir du rapport
''' de règlement Western Union.
'''
''' CE SERVICE EST PUR, ET C'EST LA DÉCISION D'ARCHITECTURE DE TOUT LE CHANTIER. Il ne lit
''' aucun fichier, n'ouvre aucune connexion, n'affiche rien et n'écrit nulle part : il reçoit
''' un DataTable et des réglages, il rend un résultat. Trois conséquences, et chacune compte :
'''
'''   1. L'ÉCRAN DE CONTRÔLE EXERCE LE CODE LIVRÉ, et non une copie. La banque a refusé un
'''      projet de tests ; le contrôle se fait donc depuis l'application. Si le service lisait
'''      lui-même son fichier et sa base, l'écran ne pourrait vérifier que ce qu'il voit, pas
'''      ce qui sera comptabilisé.
'''   2. LE MÊME CALCUL SERT LES DEUX APPELANTS — l'écran de contrôle aujourd'hui, le
'''      traitement de la journée demain. Deux chemins de calcul auraient fini par diverger
'''      d'un franc, et ce franc aurait été introuvable.
'''   3. IL SE REJOUE À L'IDENTIQUE. Même rapport, mêmes réglages, même résultat : rien dans
'''      ce service ne dépend de l'heure, de l'utilisateur ni de l'état de la base.
'''
''' D'OÙ VIENT L'ÉCART. Western Union règle la banque en EUROS et le guichet encaisse en
''' FRANCS CFA. Le rapport porte les deux montants : le montant local réellement encaissé, et
''' le montant en devise réellement réglé. Converti à la parité fixe, le second ne retombe
''' pas exactement sur le premier — d'un ou deux francs le plus souvent, de quelques milliers
''' sur les gros envois. Cette différence est un gain ou une perte de change, et elle doit
''' être comptabilisée pour elle-même.
'''
''' LES RÈGLES, VALIDÉES AVEC LA BANQUE :
'''
'''   Filtrage   garder TransactionType = "T" ; écarter toutes les lignes "A" ; écarter les
'''              transactions remboursées en totalité ; écarter, en les journalisant, les
'''              lignes dont les colonnes LOC ne sont pas dans la devise attendue.
'''   Envoi (S)  montant local = RecPrincipalREC
'''              montant devise = ClearPrincipalLOC + ClearFXLOC
'''   Paiement   montant local = ClearPrincipalPAY
'''      (P)     montant devise = ClearPrincipalLOC    ET PAS ClearFXLOC : voir plus bas.
'''   Écart      montant local - arrondi(montant devise x parité), au franc.
'''
''' POURQUOI ClearFXLOC EST EXCLU DES PAIEMENTS. Ce n'est pas une asymétrie du rapport, c'est
''' une règle de notre propre comptabilisation : WUReportService.CalculerReglement ajoute
''' déjà Abs(ClearChargesLOC + ClearFXLOC) à la COMMISSION DE PAIEMENT de la banque. Le
''' reprendre ici le compterait DEUX FOIS — une fois en produit de commission, une fois en
''' gain de change — et la pièce de change gonflerait d'autant. La vérification sur le
''' rapport réel le confirme : les paiements y présentent un écart de change nul, ce qui est
''' le signe que leur part de change est intégralement captée ailleurs.
''' </summary>
Public NotInheritable Class ChangeService

    Private Sub New()
    End Sub

#Region "Calcul"

    ''' <summary>
    ''' Calcule les écarts de change d'un rapport de règlement.
    '''
    ''' Ne lève pas d'exception sur un rapport incomplet : une colonne manquante est signalée
    ''' par <see cref="ColonnesManquantes"/>, qu'il faut appeler AVANT, et un rapport dont
    ''' toutes les lignes sont écartées rend un résultat vide dont les exclusions disent
    ''' pourquoi. Un calcul comptable ne doit pas tomber, il doit se justifier.
    ''' </summary>
    ''' <param name="table">Le rapport de règlement, tel que WUReportService.LireRapportWU le rend.</param>
    ''' <param name="options">Les réglages. Nothing prend ceux de la base — voir OptionsChangeWU.Actuelles.</param>
    ''' <param name="nomDuFichier">Nom du rapport, conservé avec le résultat pour l'archivage.</param>
    Public Shared Function Calculer(table As DataTable,
                                    options As OptionsChangeWU,
                                    nomDuFichier As String) As ResultatChangeWU

        Dim reglages As OptionsChangeWU = If(options, OptionsChangeWU.Actuelles())

        Dim resultat As New ResultatChangeWU()
        resultat.Options = reglages
        resultat.FichierSource = If(nomDuFichier, String.Empty)

        If table Is Nothing Then Return resultat

        resultat.LignesLues = table.Rows.Count

        ' Les transactions remboursées en totalité sont connues AVANT le parcours : la ligne
        ' d'ajustement qui les annule peut figurer n'importe où dans le fichier, y compris
        ' après l'originale. Un parcours unique aurait donc retenu l'originale puis découvert
        ' trop tard qu'elle était annulée.
        Dim rembourses As HashSet(Of String) = MtcnRembourses(table)

        Dim numero As Integer = 1

        For Each row As DataRow In table.Rows

            numero += 1   ' la ligne 1 du fichier est l'en-tête : la première donnée est la 2

            Dim exclusion As ExclusionChangeWU = MotifDExclusion(row, rembourses, reglages, numero)
            If exclusion IsNot Nothing Then
                resultat.Exclusions.Add(exclusion)
                Continue For
            End If

            Dim ecart As EcartChangeWU = Convertir(row, reglages, resultat.FichierSource)
            If ecart Is Nothing Then
                resultat.Exclusions.Add(New ExclusionChangeWU() With {
                    .NumeroDeLigne = numero,
                    .Mtcn = LireTexte(row, ConstantesWU.COLONNE_MTCN),
                    .Motif = "sens de l'opération inconnu",
                    .Detail = $"{ConstantesWU.COLONNE_SENS} = " &
                              $"« {LireTexte(row, ConstantesWU.COLONNE_SENS)} »"})
                Continue For
            End If

            resultat.Ecarts.Add(ecart)
        Next

        Return resultat
    End Function

    ''' <summary>
    ''' Les colonnes attendues que le rapport ne porte pas. Vide quand il est complet.
    '''
    ''' À APPELER AVANT Calculer, ET C'EST IMPORTANT. Une colonne absente serait lue 0 par
    ''' ToDecimalSafe, sans aucune erreur : un rapport privé de ClearPrincipalLOC donnerait
    ''' une contre-valeur nulle et un « gain de change » égal à la totalité des montants
    ''' encaissés. Des centaines de millions de francs, et pas une exception pour le dire.
    ''' </summary>
    Public Shared Function ColonnesManquantes(table As DataTable) As List(Of String)

        Dim absentes As New List(Of String)

        If table Is Nothing Then
            absentes.AddRange(ConstantesWU.ColonnesEcartsDeChange)
            Return absentes
        End If

        For Each colonne As String In ConstantesWU.ColonnesEcartsDeChange
            If Not table.Columns.Contains(colonne) Then absentes.Add(colonne)
        Next

        Return absentes
    End Function

#End Region

#Region "Filtrage"

    ''' <summary>
    ''' Les MTCN que le rapport annule par un remboursement total.
    '''
    ''' DEUX COLONNES SONT LUES, ET LES DEUX SONT NÉCESSAIRES. Une ligne d'ajustement
    ''' « FOR FULL REFUND » porte son propre MTCN, et désigne parfois celui de la transaction
    ''' d'origine dans AdjustmentMTCN. Les deux formes se rencontrent dans le rapport réel ;
    ''' ne lire que la première laisserait passer des originales annulées.
    '''
    ''' CE QUE CE FILTRE NE PEUT PAS FAIRE, et qu'il faut savoir : il n'attrape que les
    ''' originales présentes DANS LE MÊME FICHIER. Sur le rapport étudié, treize MTCN sont
    ''' désignés par un remboursement et un seul a son originale dans le fichier — les douze
    ''' autres ont été réglés une semaine plus tôt, dans un rapport déjà comptabilisé. Leur
    ''' change a donc été constaté, et son annulation devra l'être à son tour : c'est une
    ''' régularisation, pas un oubli de filtrage, et elle sort du calcul d'une journée.
    ''' </summary>
    Private Shared Function MtcnRembourses(table As DataTable) As HashSet(Of String)

        Dim annules As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        For Each row As DataRow In table.Rows

            If Not EstAjustement(row) Then Continue For

            If Not String.Equals(LireTexte(row, ConstantesWU.COLONNE_TYPE_AJUSTEMENT),
                                 ConstantesWU.AJUSTEMENT_REMBOURSEMENT_TOTAL,
                                 StringComparison.OrdinalIgnoreCase) Then Continue For

            Dim propre As String = LireTexte(row, ConstantesWU.COLONNE_MTCN)
            If propre.Length > 0 Then annules.Add(propre)

            Dim designe As String = LireTexte(row, ConstantesWU.COLONNE_MTCN_AJUSTE)
            If designe.Length > 0 Then annules.Add(designe)
        Next

        Return annules
    End Function

    ''' <summary>
    ''' Le motif pour lequel une ligne est écartée, ou Nothing si elle est retenue.
    '''
    ''' L'ORDRE DES TESTS EST L'ORDRE DE LECTURE DU MOTIF. Une ligne d'ajustement libellée
    ''' dans une autre devise doit être journalisée comme un ajustement, et non comme un
    ''' problème de devise : le motif annoncé est celui qui explique vraiment l'exclusion.
    ''' </summary>
    Private Shared Function MotifDExclusion(row As DataRow,
                                            rembourses As HashSet(Of String),
                                            reglages As OptionsChangeWU,
                                            numero As Integer) As ExclusionChangeWU

        Dim mtcn As String = LireTexte(row, ConstantesWU.COLONNE_MTCN)
        Dim typeDeLigne As String = LireTexte(row, ConstantesWU.COLONNE_TYPE_TRANSACTION)

        If EstAjustement(row) Then
            Dim nature As String = LireTexte(row, ConstantesWU.COLONNE_TYPE_AJUSTEMENT)
            Return New ExclusionChangeWU() With {
                .NumeroDeLigne = numero, .Mtcn = mtcn,
                .Motif = "ligne d'ajustement : aucun change propre",
                .Detail = If(nature.Length = 0, "ajustement sans type", nature)}
        End If

        If Not String.Equals(typeDeLigne, ConstantesWU.TYPE_TRANSACTION_NORMALE,
                             StringComparison.OrdinalIgnoreCase) Then
            Return New ExclusionChangeWU() With {
                .NumeroDeLigne = numero, .Mtcn = mtcn,
                .Motif = "nature de ligne non traitée",
                .Detail = $"{ConstantesWU.COLONNE_TYPE_TRANSACTION} = « {typeDeLigne} »"}
        End If

        If mtcn.Length > 0 AndAlso rembourses.Contains(mtcn) Then
            Return New ExclusionChangeWU() With {
                .NumeroDeLigne = numero, .Mtcn = mtcn,
                .Motif = "transaction remboursée en totalité",
                .Detail = ConstantesWU.AJUSTEMENT_REMBOURSEMENT_TOTAL}
        End If

        Dim devise As String = LireTexte(row, ConstantesWU.COLONNE_DEVISE_LOC)
        If Not String.Equals(devise, reglages.DeviseAttendue, StringComparison.OrdinalIgnoreCase) Then
            Return New ExclusionChangeWU() With {
                .NumeroDeLigne = numero, .Mtcn = mtcn,
                .Motif = $"montants de règlement libellés en {If(devise.Length = 0, "(devise absente)", devise)} " &
                         $"et non en {reglages.DeviseAttendue} : aucune conversion",
                .Detail = $"{ConstantesWU.COLONNE_DEVISE_LOC} = « {devise} »"}
        End If

        If Not reglages.InclureEnvoisEnAttente Then

            Dim statut As String = LireTexte(row, ConstantesWU.COLONNE_STATUT_REGLEMENT)
            If String.Equals(statut, ConstantesWU.STATUT_REGLEMENT_EN_ATTENTE,
                             StringComparison.OrdinalIgnoreCase) Then
                Return New ExclusionChangeWU() With {
                    .NumeroDeLigne = numero, .Mtcn = mtcn,
                    .Motif = "envoi en attente de règlement, écarté sur option",
                    .Detail = $"{ConstantesWU.COLONNE_STATUT_REGLEMENT} = « {statut} »"}
            End If
        End If

        Return Nothing
    End Function

    Private Shared Function EstAjustement(row As DataRow) As Boolean

        Return String.Equals(LireTexte(row, ConstantesWU.COLONNE_TYPE_TRANSACTION),
                             ConstantesWU.TYPE_TRANSACTION_AJUSTEMENT,
                             StringComparison.OrdinalIgnoreCase)
    End Function

#End Region

#Region "Conversion d'une ligne"

    ''' <summary>
    ''' Construit l'écart d'une ligne retenue. Nothing si son sens n'est ni un envoi ni un
    ''' paiement : l'appelant la journalise alors plutôt que de l'inventer à zéro.
    ''' </summary>
    Private Shared Function Convertir(row As DataRow,
                                      reglages As OptionsChangeWU,
                                      nomDuFichier As String) As EcartChangeWU

        Dim sens As String = LireTexte(row, ConstantesWU.COLONNE_SENS).ToUpperInvariant()

        Dim principalDevise As Decimal = LireNombre(row, ConstantesWU.COLONNE_PRINCIPAL_DEVISE)
        Dim changeDevise As Decimal = LireNombre(row, ConstantesWU.COLONNE_CHANGE_DEVISE)

        Dim montantLocal As Decimal
        Dim montantEnDevise As Decimal

        Select Case sens

            Case ConstantesWU.SENS_ENVOI
                montantLocal = LireNombre(row, ConstantesWU.COLONNE_PRINCIPAL_ENVOI_LOCAL)
                montantEnDevise = principalDevise + changeDevise

            Case ConstantesWU.SENS_PAIEMENT
                montantLocal = LireNombre(row, ConstantesWU.COLONNE_PRINCIPAL_PAYE_LOCAL)
                ' ClearFXLOC volontairement absent : voir l'en-tête de cette classe.
                montantEnDevise = principalDevise

            Case Else
                Return Nothing
        End Select

        Return New EcartChangeWU() With {
            .Mtcn = LireTexte(row, ConstantesWU.COLONNE_MTCN),
            .DateReglement = DateDeReglement(row),
            .Sens = sens,
            .CodeProduit = LireTexte(row, ConstantesWU.COLONNE_CODE_PRODUIT),
            .Statut = LireTexte(row, ConstantesWU.COLONNE_STATUT_REGLEMENT),
            .DeviseLocale = LireTexte(row, ConstantesWU.COLONNE_DEVISE_LOC),
            .MontantLocal = montantLocal,
            .MontantEnDevise = montantEnDevise,
            .ClearPrincipalLoc = principalDevise,
            .ClearFxLoc = changeDevise,
            .Parite = reglages.Parite,
            .FichierSource = nomDuFichier}
    End Function

#End Region

#Region "Lecture d'une cellule"

    ''' <summary>
    ''' Le texte d'une cellule, jamais Nothing, toujours débarrassé de ses espaces de bordure.
    ''' Une colonne absente rend une chaîne vide : le défaut est alors signalé par
    ''' ColonnesManquantes, et non par une exception au milieu d'un parcours de 2 400 lignes.
    ''' </summary>
    Private Shared Function LireTexte(row As DataRow, colonne As String) As String

        If row Is Nothing OrElse Not row.Table.Columns.Contains(colonne) Then Return String.Empty

        Dim valeur As Object = row(colonne)
        If valeur Is Nothing OrElse valeur Is DBNull.Value Then Return String.Empty

        Return valeur.ToString().Trim()
    End Function

    ''' <summary>
    ''' Le nombre d'une cellule. La conversion est confiée à WUReportService.ToDecimalSafe,
    ''' qui sait lire les deux formats de rapport — séparateur décimal point ou virgule — et
    ''' qui est la fonction employée par le calcul de la compensation. Le change et la
    ''' compensation doivent lire les mêmes chiffres de la même façon.
    ''' </summary>
    Private Shared Function LireNombre(row As DataRow, colonne As String) As Decimal

        If row Is Nothing OrElse Not row.Table.Columns.Contains(colonne) Then Return 0D

        Return WUReportService.ToDecimalSafe(row(colonne))
    End Function

    ''' <summary>
    ''' La date de règlement de la LIGNE, reconstituée depuis son triplet Année/Mois/Jour.
    '''
    ''' Les préfixes sont ceux du rapport de règlement — SetDateLOC puis RepDate — essayés dans
    ''' cet ordre, comme le fait WUReportService pour la date du rapport entier. La différence
    ''' est qu'ici la date est lue LIGNE PAR LIGNE : une pièce de change couvrant une semaine
    ''' doit pouvoir dire de quel jour vient chaque écart.
    '''
    ''' Nothing quand aucun triplet n'est exploitable. La date ne participe à aucun calcul :
    ''' son absence n'empêche jamais un écart d'être comptabilisé.
    ''' </summary>
    Private Shared Function DateDeReglement(row As DataRow) As Date?

        For Each prefixe As String In ConstantesWU.PrefixesDateReglement

            Dim annee, mois, jour As Integer

            If Not Integer.TryParse(LireTexte(row, prefixe & ConstantesWU.SUFFIXE_DATE_ANNEE), annee) Then Continue For
            If Not Integer.TryParse(LireTexte(row, prefixe & ConstantesWU.SUFFIXE_DATE_MOIS), mois) Then Continue For
            If Not Integer.TryParse(LireTexte(row, prefixe & ConstantesWU.SUFFIXE_DATE_JOUR), jour) Then Continue For

            If annee <= 0 OrElse mois < 1 OrElse mois > 12 OrElse jour < 1 OrElse jour > 31 Then Continue For

            Try
                Return New Date(annee, mois, jour)
            Catch ex As ArgumentOutOfRangeException
                ' Triplet incohérent (31 février) : on essaie le préfixe suivant.
            End Try
        Next

        Return Nothing
    End Function

#End Region

End Class
