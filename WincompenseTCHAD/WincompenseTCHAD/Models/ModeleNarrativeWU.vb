Option Strict On
Option Explicit On

''' <summary>
''' LE MOTEUR DE RENDU DES TEXTES DE LA PIÈCE : un gabarit, des valeurs, une phrase.
'''
''' DEUX RENDUS, PARCE QUE DEUX DESTINATIONS. Appliquer produit LA NARRATIVE DU POINT DE
''' VENTE — « LD WU ACTIVITE <point de vente> <période> », en majuscules, espaces résorbés —
''' que le core banking reçoit dans ADDLTEXT et que la RAISON reprend en bas de pièce.
''' AppliquerAuLibelle produit LE LIBELLÉ D'UNE LIGNE : majuscules elles aussi — la banque
''' les veut partout — mais LES ESPACES INTACTS, parce que la pièce doit reproduire la pièce
''' manuelle, qui écrit « TTA (TAXE SUR RECEPTION  DE FONDS WU) » avec deux espaces.
'''
''' LES CONFONDRE, C'ÉTAIT IMPOSER À LA PIÈCE LES RÈGLES DE FORME DU CORE BANKING. Elles ne
''' se ressemblent que de loin : l'une est une phrase de système, l'autre un document que la
''' Direction Comptable rapproche ligne à ligne, et dont le moindre espace compte.
'''
''' POURQUOI UN MODÈLE, ET NON UNE CONSTANTE DE PLUS
'''
''' Cette phrase a changé QUATRE FOIS en quatre livraisons : « CCS_{0} ACTIVITE WU », puis le
''' préfixe LD, puis la forme dictée « LD WU ACTIVITE &lt;point de vente&gt; &lt;période&gt; »,
''' puis les majuscules. Chaque mot a coûté un cycle complet — modification, compilation,
''' commit, pull, redéploiement sur les postes de la banque. Le texte que la Direction
''' Comptable veut lire dans son grand livre n'est pas une règle de calcul : c'est un
''' paramètre, et il appartient à la banque.
'''
''' CE QU'IL NE REND PAS PARAMÉTRABLE, ET C'EST VOULU
'''
'''   — LES MAJUSCULES sont posées ici, pas laissées à la saisie, et sur LES DEUX rendus :
'''     une seule saisie en minuscules suffirait à faire sortir une ligne qui ne ressemble pas
'''     aux onze autres. Elles sont INVARIANTES — le « i » turc deviendrait « İ ».
'''   — LA PÉRIODE s'écrit toujours de la même façon (PieceComptableService.SuffixeDePeriode).
'''     La banque choisit OÙ elle se place dans la phrase, pas comment elle s'écrit : une date
'''     mal formée dans un narratif comptable ne se rattrape pas après coup.
'''   — LES JETONS sont une LISTE FERMÉE. Un jeton inconnu est refusé à l'enregistrement, et
'''     non recopié tel quel : « LD WU ACTIVITE {AGENCY} » partirait au grand livre avec ses
'''     accolades, sur douze lignes et autant de points de vente.
'''
''' CETTE CLASSE NE TOUCHE NI LA BASE NI L'ÉCRAN. Elle reçoit un gabarit et des valeurs, elle
''' rend une phrase. C'est ce qui permet de la rejouer sur les 2 395 transactions réelles et
''' sur tout le référentiel avant que la banque ne la voie.
'''
''' CE QUI N'A PAS BESOIN D'ÊTRE HISTORISÉ. Une pièce déjà produite ne bouge pas quand le
''' modèle change : son libellé est conservé LIGNE PAR LIGNE dans T_PieceWU — et dans la
''' table d'archive, et dans T_PieceChangeWU. Une pièce archivée ré-exportée, en Excel comme
''' en fichier core banking, ressort avec le texte qu'elle avait le jour de sa génération.
''' </summary>
Public NotInheritable Class ModeleNarrativeWU

    Private Sub New()
    End Sub

#Region "Les jetons"

    ''' <summary>Désignation du point de vente, telle que le référentiel la porte.</summary>
    Public Const JETON_AGENCE As String = "{AGENCE}"

    ''' <summary>Numéro Account du point de vente (AHB020013, …).</summary>
    Public Const JETON_ACCOUNT As String = "{ACCOUNT}"

    ''' <summary>Code agence, celui-là même qui part dans la colonne ACBRN du core banking.</summary>
    Public Const JETON_CODE_AGENCE As String = "{CODE_AGENCE}"

    ''' <summary>Période couverte par la pièce, dans la forme « DU 08 AU 14 09 2026 ».</summary>
    Public Const JETON_PERIODE As String = "{PERIODE}"

    ''' <summary>
    ''' Un jeton et ce qu'il veut dire. L'écran de paramétrage lit CETTE liste pour afficher
    ''' sa légende : s'il écrivait la sienne, les deux divergeraient au premier jeton ajouté,
    ''' et la banque saisirait un jeton que la légende annonce mais que la substitution ignore.
    ''' </summary>
    Public NotInheritable Class Jeton

        Public Sub New(nom As String, description As String)
            _nom = nom
            _description = description
        End Sub

        Private ReadOnly _nom As String
        Private ReadOnly _description As String

        Public ReadOnly Property Nom As String
            Get
                Return _nom
            End Get
        End Property

        Public ReadOnly Property Description As String
            Get
                Return _description
            End Get
        End Property
    End Class

    ''' <summary>Les jetons reconnus, dans l'ordre où la légende les présente.</summary>
    Public Shared Function Jetons() As List(Of Jeton)

        Dim liste As New List(Of Jeton)

        liste.Add(New Jeton(JETON_AGENCE, "désignation du point de vente (agence ou sous-agent)"))
        liste.Add(New Jeton(JETON_ACCOUNT, "numéro Account du point de vente"))
        liste.Add(New Jeton(JETON_CODE_AGENCE, "code agence, celui qui part en ACBRN"))
        liste.Add(New Jeton(JETON_PERIODE, "période couverte, « DU 08 AU 14 09 2026 »"))

        Return liste
    End Function

#End Region

#Region "Le gabarit"

    ''' <summary>
    ''' Le modèle appliqué tant que la banque n'en a saisi aucun : la forme actuelle, au
    ''' caractère près. Une base où le script 22 n'a pas été joué sort donc exactement les
    ''' narratives d'aujourd'hui.
    ''' </summary>
    Public Shared ReadOnly Property ModeleParDefaut As String
        Get
            Return ConstantesWU.NARRATIVE_MODELE_DEFAUT
        End Get
    End Property

#End Region

#Region "Application"

    ''' <summary>
    ''' LA NARRATIVE DU POINT DE VENTE : jetons remplacés, espaces résorbés, le tout en
    ''' MAJUSCULES. C'est elle que le fichier core banking reçoit dans ADDLTEXT, et c'est elle
    ''' que la RAISON porte en bas de la pièce.
    '''
    ''' CE N'EST PAS LE RENDU DES LIBELLÉS DE LIGNE : ceux-là passent par AppliquerAuLibelle,
    ''' qui rend ce que la banque a écrit, sans toucher ni à la casse ni aux espaces. Les deux
    ''' destinations n'ont jamais eu les mêmes règles de forme, et les confondre revenait à
    ''' imposer à la pièce comptable celles du core banking.
    '''
    ''' UN JETON VIDE NE LAISSE PAS DE TROU : la pièce globale ne porte aucun point de vente,
    ''' et « LD WU ACTIVITE  DU 08 AU 14 09 2026 » afficherait un double espace là où un
    ''' comptable lit une donnée manquante.
    ''' </summary>
    Public Shared Function Appliquer(modele As String, designation As String, account As String,
                                     codeAgence As String, periode As String) As String

        Return ResorberLesEspaces(
            Substituer(modele, designation, account, codeAgence, periode, False)).ToUpperInvariant()
    End Function

    ''' <summary>
    ''' LE LIBELLÉ D'UNE LIGNE DE PIÈCE : jetons remplacés, MAJUSCULES, et les espaces tels
    ''' que la banque les a écrits.
    '''
    ''' LES MAJUSCULES SONT UNE RÈGLE, PAS UNE SAISIE, et la banque les veut sur la pièce comme
    ''' sur la narrative : « Commission sur Transfert_Ecobank » sort « COMMISSION SUR
    ''' TRANSFERT_ECOBANK ». Les poser ici plutôt que de les attendre du gabarit évite qu'une
    ''' seule saisie en minuscules, dans l'écran de paramétrage, fasse sortir une ligne qui ne
    ''' ressemble pas aux onze autres.
    '''
    ''' INVARIANTES, et non celles du poste : le « i » turc deviendrait « İ », et un poste
    ''' réglé en turc produirait des libellés différents de ceux des autres postes.
    '''
    ''' LES ESPACES, EUX, NE SONT PAS TOUCHÉS. La pièce manuelle de la banque — celle que la
    ''' Direction Comptable rapproche ligne à ligne — écrit « TTA (TAXE SUR RECEPTION  DE FONDS
    ''' WU) » avec deux espaces. Les résorber ferait diverger les deux documents sur cette
    ''' ligne, et c'est le rapprochement qui en pâtirait. C'est là toute la différence avec
    ''' Appliquer, qui les résorbe parce que le core banking n'a pas de pièce à rapprocher.
    '''
    ''' UN JETON VIDE EMPORTE UN ESPACE AVEC LUI, celui qui le précède, à défaut celui qui le
    ''' suit. C'est ce qui évite le trou de « Commission sur Transfert_Sous-agence  » sans
    ''' toucher au double espace que la banque a, lui, VOULU écrire — un ResorberLesEspaces ne
    ''' sait pas distinguer les deux, celui-ci n'a pas à le faire.
    ''' </summary>
    Public Shared Function AppliquerAuLibelle(modele As String, designation As String,
                                              account As String, codeAgence As String,
                                              periode As String) As String

        Return Substituer(
            modele, designation, account, codeAgence, periode, True).Trim().ToUpperInvariant()
    End Function

    ''' <summary>
    ''' Le parcours unique du gabarit, partagé par les deux rendus.
    '''
    ''' UN SEUL PARCOURS, et non quatre Replace successifs. Deux raisons :
    '''
    '''   — {CODE_AGENCE} contient {AGENCE} à la lecture d'un Replace naïf, qui aurait produit
    '''     « {CODE_ECOBANK MOUNDOU} » ;
    '''   — String.Replace du Framework 4.8 ne connaît pas StringComparison, et un jeton saisi
    '''     « {agence} » doit être reconnu. Un comptable qui tape en minuscules a raison de
    '''     s'attendre à ce que ça marche.
    ''' </summary>
    ''' <param name="mangerLEspaceDunJetonVide">
    ''' Vrai pour un libellé de pièce : un jeton sans valeur disparaît AVEC un espace, et le
    ''' reste du texte est rendu tel quel. Faux pour la narrative, dont les espaces sont
    ''' résorbés ensuite de toute façon.
    ''' </param>
    Private Shared Function Substituer(modele As String, designation As String, account As String,
                                       codeAgence As String, periode As String,
                                       mangerLEspaceDunJetonVide As Boolean) As String

        Dim gabarit As String = If(String.IsNullOrWhiteSpace(modele), ModeleParDefaut, modele)

        Dim valeurs As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        valeurs.Add(JETON_AGENCE, If(designation, String.Empty).Trim())
        valeurs.Add(JETON_ACCOUNT, If(account, String.Empty).Trim())
        valeurs.Add(JETON_CODE_AGENCE, If(codeAgence, String.Empty).Trim())
        valeurs.Add(JETON_PERIODE, If(periode, String.Empty).Trim())

        Dim rendu As String = String.Empty
        Dim position As Integer = 0

        Do While position < gabarit.Length

            Dim debut As Integer = gabarit.IndexOf("{"c, position)

            If debut < 0 Then
                rendu &= gabarit.Substring(position)
                Exit Do
            End If

            rendu &= gabarit.Substring(position, debut - position)

            Dim fin As Integer = gabarit.IndexOf("}"c, debut + 1)

            ' Accolade jamais refermée : le reste du gabarit est recopié tel quel. Controler
            ' refuse ce cas à l'enregistrement ; si un tel modèle se trouve malgré tout en
            ' base — saisi en SQL direct — la pièce sort avec une phrase bizarre plutôt que
            ' de ne pas sortir du tout.
            If fin < 0 Then
                rendu &= gabarit.Substring(debut)
                Exit Do
            End If

            Dim jeton As String = gabarit.Substring(debut, fin - debut + 1)

            ' Un jeton inconnu est recopié avec ses accolades : il n'est jamais vide, et ne
            ' déclenche donc pas la règle de l'espace mangé.
            Dim valeur As String = jeton
            If valeurs.ContainsKey(jeton) Then valeur = valeurs(jeton)

            position = fin + 1

            If mangerLEspaceDunJetonVide AndAlso valeur.Length = 0 Then

                If rendu.EndsWith(" ", StringComparison.Ordinal) Then
                    rendu = rendu.Substring(0, rendu.Length - 1)
                ElseIf position < gabarit.Length AndAlso gabarit.Chars(position) = " "c Then
                    position += 1
                End If

                Continue Do
            End If

            rendu &= valeur
        Loop

        Return rendu
    End Function

    ''' <summary>
    ''' Ramène les suites d'espaces à un seul, et coupe ceux des extrémités. C'est ce qui rend
    ''' un jeton vide invisible au lieu de laisser un trou dans la phrase.
    ''' </summary>
    Private Shared Function ResorberLesEspaces(texte As String) As String

        Dim resultat As String = If(texte, String.Empty)

        Do While resultat.Contains("  ")
            resultat = resultat.Replace("  ", " ")
        Loop

        Return resultat.Trim()
    End Function

#End Region

#Region "Contrôle"

    ''' <summary>
    ''' Dit si un modèle est enregistrable, et pourquoi il ne l'est pas.
    '''
    ''' LE CONTRÔLE DES 150 CARACTÈRES SE FAIT ICI, SUR LE PIRE CAS RÉEL. L'appelant fournit
    ''' les valeurs les plus longues du référentiel — la désignation la plus longue, l'Account
    ''' le plus long, le code agence le plus long — et la période la plus longue que
    ''' SuffixeDePeriode sache produire. La phrase est alors RENDUE et mesurée : rien n'est
    ''' estimé, et un modèle validé ici ne fera pas échouer la production du fichier core
    ''' banking le matin de la compense.
    '''
    ''' CE CONTRÔLE NE REMPLACE PAS CELUI DE CoreBankingService.ControlerLesNarratifs, qui
    ''' reste en place et refuse toujours de produire le fichier au-delà de 150 caractères. Un
    ''' point de vente RENOMMÉ après l'enregistrement du modèle peut faire dépasser une phrase
    ''' que cet écran avait validée : il faut alors que quelque chose s'y oppose encore.
    '''
    ''' UN LIBELLÉ DE PIÈCE N'EST PAS BORNÉ PAR LES MÊMES 150 CARACTÈRES. Il ne part pas dans
    ''' ADDLTEXT : il va dans la colonne Libelle de T_PieceWU, un NVARCHAR(255). Lui imposer
    ''' la limite du core banking ferait refuser à l'écran un libellé que la pièce accepte
    ''' parfaitement. Le paramètre pourLaPiece dit laquelle des deux destinations on contrôle,
    ''' et il commande AUSSI le rendu : un libellé se mesure tel qu'il sortira, espaces non
    ''' résorbés.
    ''' </summary>
    ''' <param name="pourLaPiece">
    ''' Vrai pour un libellé de ligne (255 caractères, rendu tel quel), faux pour la narrative
    ''' du point de vente (150 caractères, espaces résorbés).
    ''' </param>
    Public Shared Function Controler(modele As String,
                                     pireDesignation As String, pireAccount As String,
                                     pireCodeAgence As String, pirePeriode As String,
                                     ByRef messageErreur As String,
                                     Optional pourLaPiece As Boolean = False) As Boolean

        messageErreur = String.Empty

        Dim gabarit As String = If(modele, String.Empty).Trim()

        If gabarit.Length = 0 Then
            messageErreur = "Le modèle de narrative est vide. Chaque ligne de la pièce doit porter " &
                            "un libellé : la comptabilité n'impute pas une écriture muette."
            Return False
        End If

        If gabarit.Length > ConstantesWU.NARRATIVE_MODELE_LONGUEUR_MAX Then
            messageErreur = $"Le modèle fait {gabarit.Length} caractères, pour un maximum de " &
                            $"{ConstantesWU.NARRATIVE_MODELE_LONGUEUR_MAX}. Au-delà, la base le " &
                            "tronquerait sans rien dire."
            Return False
        End If

        Dim inconnus As List(Of String) = JetonsInconnus(gabarit)

        If inconnus.Count > 0 Then
            messageErreur = "Le modèle emploie un repère que l'application ne connaît pas : " &
                            String.Join(", ", inconnus) & "." & Environment.NewLine & Environment.NewLine &
                            "Il partirait tel quel au grand livre, accolades comprises. " &
                            "Les repères reconnus sont : " &
                            String.Join(", ", Jetons().Select(Function(j) j.Nom)) & "."
            Return False
        End If

        Dim pireCas As String
        If pourLaPiece Then
            pireCas = AppliquerAuLibelle(gabarit, pireDesignation, pireAccount,
                                         pireCodeAgence, pirePeriode)
        Else
            pireCas = Appliquer(gabarit, pireDesignation, pireAccount,
                                pireCodeAgence, pirePeriode)
        End If

        If pireCas.Length = 0 Then
            messageErreur = "Ce modèle ne produit aucun texte : il ne contient que des repères, " &
                            "et le référentiel ne les renseigne pas."
            Return False
        End If

        Dim maximum As Integer = If(pourLaPiece, ConstantesWU.NARRATIVE_MODELE_LONGUEUR_MAX,
                                    ConstantesWU.CB_NARRATIF_LONGUEUR_MAX)

        If pireCas.Length > maximum Then
            messageErreur = $"Ce modèle produirait {pireCas.Length} caractères dans le cas le plus " &
                            If(pourLaPiece,
                               $"long, alors que la colonne de la pièce en accepte {maximum}.",
                               $"long, alors que le core banking en accepte {maximum}.") &
                            Environment.NewLine & Environment.NewLine & pireCas & Environment.NewLine &
                            Environment.NewLine &
                            If(pourLaPiece,
                               "Le libellé serait tronqué en base. ",
                               "Le fichier ne pourrait pas être produit pour ce point de vente. ") &
                            "Raccourcissez le modèle, ou retirez-en un repère."
            Return False
        End If

        Return True
    End Function

    ''' <summary>
    ''' Vrai si le modèle emploie au moins un repère.
    '''
    ''' UN MODÈLE SANS AUCUN REPÈRE N'EST PAS REFUSÉ, IL EST SIGNALÉ. C'est un choix que la
    ''' banque a le droit de faire, et c'en est un qu'elle a le droit de regretter : toutes
    ''' les écritures de tous les points de vente de toutes les journées porteraient alors le
    ''' même texte, et le grand livre ne dirait plus ni QUI ni QUAND. C'est exactement ce que
    ''' la banque nous reprochait quand les libellés ne nommaient pas la période.
    '''
    ''' Le refuser serait décider à sa place ; le taire serait la laisser se tromper sans
    ''' qu'elle le voie. L'écran l'écrit donc, avant l'enregistrement.
    ''' </summary>
    Public Shared Function EmploieUnRepere(modele As String) As Boolean

        Dim gabarit As String = If(modele, String.Empty)

        For Each reconnu As Jeton In Jetons()
            If gabarit.IndexOf(reconnu.Nom, StringComparison.OrdinalIgnoreCase) >= 0 Then Return True
        Next

        Return False
    End Function

    ''' <summary>
    ''' Les repères du modèle qui ne figurent pas dans la liste fermée, sans doublon et dans
    ''' l'ordre où ils apparaissent.
    '''
    ''' Une accolade ouvrante jamais refermée est signalée elle aussi : « {AGENCE » n'est pas
    ''' un jeton, et ce qui le suit serait recopié jusqu'au bout de la phrase.
    ''' </summary>
    Private Shared Function JetonsInconnus(modele As String) As List(Of String)

        Dim inconnus As New List(Of String)
        Dim gabarit As String = If(modele, String.Empty)
        Dim connus As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        For Each reconnu As Jeton In Jetons()
            connus.Add(reconnu.Nom)
        Next

        Dim position As Integer = 0

        Do
            Dim debut As Integer = gabarit.IndexOf("{"c, position)
            If debut < 0 Then Exit Do

            Dim fin As Integer = gabarit.IndexOf("}"c, debut + 1)

            If fin < 0 Then
                inconnus.Add(gabarit.Substring(debut))
                Exit Do
            End If

            Dim jeton As String = gabarit.Substring(debut, fin - debut + 1)

            If Not connus.Contains(jeton) AndAlso
               Not inconnus.Any(Function(deja) String.Equals(deja, jeton, StringComparison.OrdinalIgnoreCase)) Then
                inconnus.Add(jeton)
            End If

            position = fin + 1
        Loop

        Return inconnus
    End Function

#End Region

End Class
