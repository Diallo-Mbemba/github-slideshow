Option Strict On
Option Explicit On

''' <summary>
''' Les réglages du calcul des écarts de change : la parité, la devise attendue, et le sort
''' des envois non encore réglés.
'''
''' POURQUOI UN OBJET PLUTÔT QUE TROIS PARAMÈTRES. Le service de calcul est appelé de deux
''' endroits — l'écran de contrôle et, demain, le traitement de la journée — et ces deux
''' appels doivent employer EXACTEMENT les mêmes règles, sinon l'écran de contrôle ne
''' contrôle rien. Les réunir en un objet, dont la fabrication par défaut est unique, est
''' ce qui le garantit : on ne peut pas oublier un réglage, on peut seulement le changer.
''' </summary>
Public Class OptionsChangeWU

    ''' <summary>
    ''' Parité employée pour convertir la devise de règlement en francs CFA. Par défaut la
    ''' parité fixe EUR/XAF, qui n'a pas bougé depuis 1999 — mais elle reste un réglage : le
    ''' jour où Western Union réglerait en dollars, la pièce ne doit pas se taire.
    ''' </summary>
    Public Property Parite As Decimal = ConstantesWU.TAUX_CONVERSION

    ''' <summary>
    ''' Devise dans laquelle les colonnes LOC du rapport doivent être libellées pour qu'une
    ''' conversion ait un sens. Une ligne libellée dans une autre devise est écartée et
    ''' journalisée, jamais convertie au hasard.
    '''
    ''' CE RÉGLAGE EST UN GARDE-FOU, PAS UNE PRÉFÉRENCE. Le rapport étudié porte EUR sur ses
    ''' 2 395 lignes ; un rapport où les colonnes LOC seraient déjà en FCFA ne produirait
    ''' aucun écart de change — et doit alors le dire, au lieu d'afficher des millions
    ''' d'écart nés d'une double conversion.
    ''' </summary>
    Public Property DeviseAttendue As String = ConstantesWU.DEVISE_EURO

    ''' <summary>
    ''' Les envois encore en attente de règlement (TxnStatus = "W") entrent-ils dans le calcul ?
    '''
    ''' VRAI PAR DÉFAUT, SUR DÉCISION DE LA BANQUE. Ce n'est pas un détail : dans le rapport
    ''' étudié, 290 envois en attente portent 277 070 F sur 555 849 F de gains — la MOITIÉ du
    ''' résultat. La banque a tranché « inclus par défaut » : la transaction est née, son
    ''' change l'est aussi, et la différer reviendrait à comptabiliser le gain dans un exercice
    ''' qui n'est pas le sien. Le réglage reste ouvert pour qu'elle puisse mesurer l'écart que
    ''' cela représente, pas pour qu'un poste en décide seul.
    ''' </summary>
    Public Property InclureEnvoisEnAttente As Boolean = True

    ''' <summary>
    ''' Les réglages en service : ceux de la base pour ce qui s'y conserve, les valeurs par
    ''' défaut pour le reste. C'est la SEULE fabrication employée par l'application ; les
    ''' appels de contrôle qui s'en écartent le font explicitement, réglage par réglage.
    ''' </summary>
    Public Shared Function Actuelles() As OptionsChangeWU

        Dim options As New OptionsChangeWU()
        options.InclureEnvoisEnAttente = OptionsWU.ChangeInclureEnvoisEnAttente
        Return options
    End Function

    ''' <summary>Les réglages, écrits en une phrase, pour l'en-tête d'un contrôle ou d'une pièce.</summary>
    Public Function Description() As String

        Return $"Parité {Parite.ToString("N3", Globalization.CultureInfo.GetCultureInfo("fr-FR"))} " &
               $"{DeviseAttendue}/{ConstantesWU.DEVISE_FCFA} — envois en attente " &
               If(InclureEnvoisEnAttente, "INCLUS", "exclus") & "."
    End Function

End Class

''' <summary>
''' Le résultat d'un calcul d'écarts de change : les transactions retenues, celles qui ont
''' été écartées, et les totaux.
'''
''' AUCUN TOTAL N'EST CONSERVÉ : TOUS SE CALCULENT À LA DEMANDE. Un total accumulé pendant
''' le parcours cesse d'être vrai dès qu'une ligne est ajoutée ou retirée, et rien ne le
''' signale. Ici le gain, la perte et le net se relisent depuis les lignes à chaque fois :
''' ils ne peuvent pas démentir le détail qui les accompagne.
'''
''' LE CONTRÔLE EST DANS L'OBJET, PAS DANS L'ÉCRAN. Somme(montants locaux) moins
''' Somme(contre-valeurs) doit faire exactement Gains moins Pertes. C'est une identité
''' arithmétique : elle tient toujours, sauf si une ligne a été comptée d'un côté et pas de
''' l'autre. La vérifier coûte une soustraction et attrape une faute de parcours ; la
''' vérifier DANS le résultat fait qu'aucun appelant ne peut l'oublier.
''' </summary>
Public Class ResultatChangeWU

#Region "Contenu"

    ''' <summary>Les transactions retenues, une par écart — y compris les écarts nuls.</summary>
    Public ReadOnly Property Ecarts As New List(Of EcartChangeWU)

    ''' <summary>Les lignes écartées et leur motif. Voir <see cref="ExclusionChangeWU"/>.</summary>
    Public ReadOnly Property Exclusions As New List(Of ExclusionChangeWU)

    ''' <summary>Les réglages employés. Conservés avec le résultat : un total sans sa parité ne veut rien dire.</summary>
    Public Property Options As OptionsChangeWU = New OptionsChangeWU()

    ''' <summary>Nom du rapport lu.</summary>
    Public Property FichierSource As String = String.Empty

    ''' <summary>Nombre de lignes du rapport, retenues et écartées confondues.</summary>
    Public Property LignesLues As Integer

    ''' <summary>
    ''' Période couverte par le rapport, telle que les dates de règlement des lignes retenues
    ''' la donnent. Nothing quand le rapport ne porte aucune date exploitable — et Nothing
    ''' plutôt qu'une date minimale, qui se serait affichée « 01/01/0001 » dans la synthèse.
    ''' </summary>
    Public ReadOnly Property PremiereDate As Date?
        Get
            Dim dates As List(Of Date) = DatesDeReglement()
            If dates.Count = 0 Then Return Nothing
            Return dates.Min()
        End Get
    End Property

    ''' <summary>Voir <see cref="PremiereDate"/>.</summary>
    Public ReadOnly Property DerniereDate As Date?
        Get
            Dim dates As List(Of Date) = DatesDeReglement()
            If dates.Count = 0 Then Return Nothing
            Return dates.Max()
        End Get
    End Property

    ''' <summary>Les dates de règlement présentes dans les lignes retenues, sans les absentes.</summary>
    Private Function DatesDeReglement() As List(Of Date)

        Return Ecarts.Where(Function(ligne) ligne.DateReglement.HasValue).
                      Select(Function(ligne) ligne.DateReglement.Value).
                      ToList()
    End Function

#End Region

#Region "Totaux"

    ''' <summary>Somme des écarts POSITIFS : ce que la banque gagne au change.</summary>
    Public ReadOnly Property Gains As Decimal
        Get
            Return Ecarts.Where(Function(ligne) ligne.Ecart > 0D).Sum(Function(ligne) ligne.Ecart)
        End Get
    End Property

    ''' <summary>
    ''' Somme des écarts NÉGATIFS, exprimée POSITIVEMENT : une perte de 1 544 F se lit 1 544,
    ''' et non -1 544. C'est ainsi qu'elle s'écrit sur une pièce comptable, au débit.
    ''' </summary>
    Public ReadOnly Property Pertes As Decimal
        Get
            Return -Ecarts.Where(Function(ligne) ligne.Ecart < 0D).Sum(Function(ligne) ligne.Ecart)
        End Get
    End Property

    ''' <summary>Gains moins pertes. Positif, la journée est gagnante.</summary>
    Public ReadOnly Property Net As Decimal
        Get
            Return Gains - Pertes
        End Get
    End Property

    Public ReadOnly Property TotalMontantLocal As Decimal
        Get
            Return Ecarts.Sum(Function(ligne) ligne.MontantLocal)
        End Get
    End Property

    Public ReadOnly Property TotalMontantEnDevise As Decimal
        Get
            Return Ecarts.Sum(Function(ligne) ligne.MontantEnDevise)
        End Get
    End Property

    Public ReadOnly Property TotalContreValeur As Decimal
        Get
            Return Ecarts.Sum(Function(ligne) ligne.ContreValeur)
        End Get
    End Property

    Public ReadOnly Property NombreDeGains As Integer
        Get
            Return Ecarts.Where(Function(ligne) ligne.Nature = NatureEcartWU.Gain).Count()
        End Get
    End Property

    Public ReadOnly Property NombreDePertes As Integer
        Get
            Return Ecarts.Where(Function(ligne) ligne.Nature = NatureEcartWU.Perte).Count()
        End Get
    End Property

    Public ReadOnly Property NombreDeNeutres As Integer
        Get
            Return Ecarts.Where(Function(ligne) ligne.Nature = NatureEcartWU.Neutre).Count()
        End Get
    End Property

    ''' <summary>Nombre d'envois en attente retenus, pour que leur poids se lise sans le chercher.</summary>
    Public ReadOnly Property NombreEnvoisEnAttente As Integer
        Get
            Return Ecarts.Where(Function(ligne) EstEnAttente(ligne)).Count()
        End Get
    End Property

    ''' <summary>Écart de change porté par les seuls envois en attente retenus.</summary>
    Public ReadOnly Property NetEnvoisEnAttente As Decimal
        Get
            Return Ecarts.Where(Function(ligne) EstEnAttente(ligne)).Sum(Function(ligne) ligne.Ecart)
        End Get
    End Property

    Private Shared Function EstEnAttente(ligne As EcartChangeWU) As Boolean
        Return String.Equals(If(ligne.Statut, String.Empty).Trim(),
                             ConstantesWU.STATUT_REGLEMENT_EN_ATTENTE,
                             StringComparison.OrdinalIgnoreCase)
    End Function

#End Region

#Region "Contrôle"

    ''' <summary>
    ''' La différence entre les deux façons de calculer le net. Elle vaut zéro, toujours ; une
    ''' autre valeur signale une faute de parcours dans le service, pas une anomalie du rapport.
    ''' </summary>
    Public ReadOnly Property EcartDeControle As Decimal
        Get
            Return (TotalMontantLocal - TotalContreValeur) - Net
        End Get
    End Property

    ''' <summary>Vrai quand le contrôle arithmétique passe. Voir <see cref="EcartDeControle"/>.</summary>
    Public ReadOnly Property EstCoherent As Boolean
        Get
            Return EcartDeControle = 0D
        End Get
    End Property

    ''' <summary>
    ''' Vrai quand le rapport n'a produit aucun écart ALORS QU'il portait des lignes : le cas
    ''' d'un rapport dont les colonnes LOC sont déjà en francs CFA. Ce n'est pas une erreur,
    ''' c'est un rapport sans change — et il faut le dire, pas afficher un total à zéro.
    ''' </summary>
    Public ReadOnly Property AucuneConversionPossible As Boolean
        Get
            Return Ecarts.Count = 0 AndAlso Exclusions.Count > 0
        End Get
    End Property

#End Region

#Region "Restitution"

    ''' <summary>
    ''' La synthèse, telle qu'elle s'affiche à l'écran de contrôle et se colle dans un courriel.
    '''
    ''' Elle est écrite ICI, et non dans le formulaire, pour une raison précise : c'est ce texte
    ''' que la banque recevra par courriel quand elle contestera un total. Il doit donc pouvoir
    ''' être produit sans ouvrir l'application — depuis le traitement de la journée, demain.
    ''' </summary>
    Public Function Synthese() As String

        Dim fr As Globalization.CultureInfo = Globalization.CultureInfo.GetCultureInfo("fr-FR")
        Dim texte As New System.Text.StringBuilder()

        texte.AppendLine("ÉCARTS DE CHANGE — SYNTHÈSE")
        texte.AppendLine(New String("="c, 72))
        texte.AppendLine()
        texte.AppendLine($"Rapport                  : {FichierSource}")
        texte.AppendLine($"Réglages                 : {Options.Description()}")

        If PremiereDate.HasValue AndAlso DerniereDate.HasValue Then
            texte.AppendLine($"Période de règlement     : du {PremiereDate.Value:dd/MM/yyyy} au {DerniereDate.Value:dd/MM/yyyy}")
        End If

        texte.AppendLine()
        texte.AppendLine($"Lignes lues              : {LignesLues:N0}")
        texte.AppendLine($"Transactions retenues    : {Ecarts.Count:N0}")
        texte.AppendLine($"Lignes écartées          : {Exclusions.Count:N0}")
        texte.AppendLine()
        texte.AppendLine($"Gains de change          : {Gains.ToString("N0", fr)} {ConstantesWU.DEVISE_FCFA}   ({NombreDeGains:N0} transactions)")
        texte.AppendLine($"Pertes de change         : {Pertes.ToString("N0", fr)} {ConstantesWU.DEVISE_FCFA}   ({NombreDePertes:N0} transactions)")
        texte.AppendLine($"Écart net                : {Net.ToString("N0", fr)} {ConstantesWU.DEVISE_FCFA}")
        texte.AppendLine($"Conversions exactes      : {NombreDeNeutres:N0} transactions")
        texte.AppendLine()
        texte.AppendLine($"Somme des montants locaux : {TotalMontantLocal.ToString("N0", fr)} {ConstantesWU.DEVISE_FCFA}")
        texte.AppendLine($"Somme des contre-valeurs  : {TotalContreValeur.ToString("N0", fr)} {ConstantesWU.DEVISE_FCFA}")
        texte.AppendLine($"Différence                : {(TotalMontantLocal - TotalContreValeur).ToString("N0", fr)} {ConstantesWU.DEVISE_FCFA}")

        ' Le verdict est composé AVANT d'être inséré : une chaîne interpolée imbriquée dans le
        ' trou d'une autre se compile, mais ne se relit pas.
        Dim verdict As String = "OK — la différence est égale au net."
        If Not EstCoherent Then
            verdict = $"ÉCART DE {EcartDeControle.ToString("N0", fr)} : ANOMALIE DE CALCUL."
        End If

        texte.AppendLine($"Contrôle                  : {verdict}")

        If NombreEnvoisEnAttente > 0 Then
            texte.AppendLine()
            texte.AppendLine($"Dont envois en attente (statut {ConstantesWU.STATUT_REGLEMENT_EN_ATTENTE}) : " &
                             $"{NombreEnvoisEnAttente:N0} transactions, {NetEnvoisEnAttente.ToString("N0", fr)} {ConstantesWU.DEVISE_FCFA} d'écart net.")
            texte.AppendLine("Les exclure ramènerait l'écart net à " &
                             $"{(Net - NetEnvoisEnAttente).ToString("N0", fr)} {ConstantesWU.DEVISE_FCFA}.")
        End If

        If AucuneConversionPossible Then
            texte.AppendLine()
            texte.AppendLine("AUCUNE TRANSACTION RETENUE.")
            texte.AppendLine("Toutes les lignes ont été écartées : si le motif est la devise, ce rapport exprime")
            texte.AppendLine($"déjà ses montants de règlement en {ConstantesWU.DEVISE_FCFA} et ne porte donc aucun écart de change.")
        End If

        texte.AppendLine()
        texte.AppendLine(DetailDesExclusions())

        Return texte.ToString()
    End Function

    ''' <summary>
    ''' Les exclusions, regroupées par motif puis détaillées ligne à ligne.
    '''
    ''' Le regroupement précède le détail parce que vingt-huit lignes se lisent, deux mille non :
    ''' le jour où un rapport écartera la moitié de ses lignes, le total par motif dira
    ''' immédiatement pourquoi, et le détail restera consultable en dessous.
    ''' </summary>
    Public Function DetailDesExclusions() As String

        Dim texte As New System.Text.StringBuilder()

        texte.AppendLine("LIGNES ÉCARTÉES")
        texte.AppendLine(New String("-"c, 72))

        If Exclusions.Count = 0 Then
            texte.AppendLine("Aucune : toutes les lignes du rapport sont entrées dans le calcul.")
            Return texte.ToString()
        End If

        For Each groupe As IGrouping(Of String, ExclusionChangeWU) In
            Exclusions.GroupBy(Function(ligne) ligne.Motif).
                       OrderByDescending(Function(parMotif) parMotif.Count())

            texte.AppendLine($"{groupe.Count(),6}  {groupe.Key}")
        Next

        texte.AppendLine()

        For Each exclusion As ExclusionChangeWU In Exclusions
            texte.AppendLine($"  ligne {exclusion.NumeroDeLigne,6}  MTCN {exclusion.Mtcn,-12}  " &
                             $"{exclusion.Motif}{If(exclusion.Detail.Length = 0, String.Empty, "  [" & exclusion.Detail & "]")}")
        Next

        Return texte.ToString()
    End Function

#End Region

End Class
