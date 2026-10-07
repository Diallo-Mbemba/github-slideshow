Option Strict On
Option Explicit On

''' <summary>
''' Comment les libellés d'une pièce sont choisis. La banque bascule d'un mode à l'autre
''' depuis l'écran « Narrative comptable ».
''' </summary>
Public Enum ModeNarrativeWU

    ''' <summary>
    ''' UN SEUL MODÈLE pour les douze lignes d'un point de vente — « LD WU ACTIVITE <point de
    ''' vente> <période> ». C'est la forme que le FICHIER CORE BANKING porte toujours, quel
    ''' que soit le mode ; ce mode-ci l'étend à la PIÈCE COMPTABLE, où ce que chaque ligne EST
    ''' se lit alors dans son seul numéro de compte.
    ''' </summary>
    ModeleUnique = 1

    ''' <summary>
    ''' UN LIBELLÉ PAR NATURE DE MOUVEMENT : « COMPTE COURANT WESTERN UNION ETD »,
    ''' « Commission sur Transfert_Ecobank », « TVA COLLECTEES WESTERN UNION »… C'EST LE MODE
    ''' PAR DÉFAUT, celui de la pièce manuelle de la banque. Ni préfixe « LD » ni période :
    ''' tous deux appartiennent à la narrative, en bas de pièce et dans ADDLTEXT. Une nature
    ''' laissée vide retombe sur ce libellé-là.
    ''' </summary>
    ParNature = 2
End Enum

''' <summary>
''' LE PARAMÉTRAGE DE NARRATIVE EN VIGUEUR : le mode, le modèle global, et les quatorze modèles
''' par nature. C'est lui qui répond à la seule question qui compte au moment de poser une
''' écriture : quel texte cette ligne-ci doit-elle porter ?
'''
''' IL EST LU UNE FOIS PAR PIÈCE, ET IL EST IMMUABLE. Une pièce ne doit jamais mélanger deux
''' paramétrages parce que quelqu'un a enregistré un libellé pendant sa génération : les
''' valeurs sont figées à la construction, et rien ici ne les change ensuite.
'''
''' IL NE TOUCHE NI LA BASE NI L'ÉCRAN. C'est NarrativeRepository qui va le chercher ; cette
''' classe se contente de trancher, et se rejoue donc hors de l'application.
'''
''' LE MODE NE CONCERNE QUE LA PIÈCE. Le fichier core banking porte toujours la narrative
''' UNIQUE du point de vente, celle du modèle global : c'est le rectificatif de la banque du
''' 06/10/2026, et c'est pourquoi la pièce transporte DEUX textes par ligne — son libellé, et
''' la narrative du point de vente dans sa colonne Narratif.
''' </summary>
Public NotInheritable Class NarrativesWU

    ''' <summary>Valeur de la colonne Valeur de T_ParametreWU pour le mode « un seul modèle ».</summary>
    Public Const CODE_MODE_UNIQUE As String = "GLOBAL"

    ''' <summary>Valeur de la colonne Valeur de T_ParametreWU pour le mode « par nature ».</summary>
    Public Const CODE_MODE_PAR_NATURE As String = "PAR_NATURE"

    Public Sub New(mode As ModeNarrativeWU, modeleGlobal As String,
                   modelesParNature As Dictionary(Of NatureMouvementWU, String))

        _mode = mode
        _modeleGlobal = If(String.IsNullOrWhiteSpace(modeleGlobal),
                           ModeleNarrativeWU.ModeleParDefaut, modeleGlobal.Trim())

        ' RECOPIÉ, et non référencé : l'appelant garde son dictionnaire et peut le remplir
        ' encore, une pièce en cours de génération ne doit pas en voir le contenu changer.
        _modeles = New Dictionary(Of NatureMouvementWU, String)

        If modelesParNature Is Nothing Then Return

        For Each paire As KeyValuePair(Of NatureMouvementWU, String) In modelesParNature
            If String.IsNullOrWhiteSpace(paire.Value) Then Continue For
            _modeles(paire.Key) = paire.Value.Trim()
        Next
    End Sub

    Private ReadOnly _mode As ModeNarrativeWU
    Private ReadOnly _modeleGlobal As String
    Private ReadOnly _modeles As Dictionary(Of NatureMouvementWU, String)

    Public ReadOnly Property Mode As ModeNarrativeWU
        Get
            Return _mode
        End Get
    End Property

    Public ReadOnly Property ModeleGlobal As String
        Get
            Return _modeleGlobal
        End Get
    End Property

#Region "Les replis"

    ''' <summary>
    ''' Le paramétrage appliqué quand rien n'a pu être lu : table absente, base injoignable,
    ''' script jamais exécuté. La pièce sort alors avec ses libellés historiques, c'est-à-dire
    ''' telle que la banque la connaît — et jamais une pièce sans libellés.
    ''' </summary>
    Public Shared Function ParDefaut() As NarrativesWU
        Return New NarrativesWU(ModeNarrativeWU.ParNature, ModeleNarrativeWU.ModeleParDefaut, Nothing)
    End Function

    ''' <summary>Le mode que désigne la valeur lue en base. Tout ce qui n'est pas franchement
    ''' « GLOBAL » vaut PAR NATURE : une valeur mal orthographiée ne doit pas effacer les
    ''' douze libellés de la pièce sans que personne ne l'ait demandé.</summary>
    Public Shared Function ModeDepuisCode(valeur As String) As ModeNarrativeWU

        If String.Equals(If(valeur, String.Empty).Trim(), CODE_MODE_UNIQUE,
                         StringComparison.OrdinalIgnoreCase) Then
            Return ModeNarrativeWU.ModeleUnique
        End If

        Return ModeNarrativeWU.ParNature
    End Function

    ''' <summary>La valeur à écrire en base pour un mode.</summary>
    Public Shared Function CodeDeMode(mode As ModeNarrativeWU) As String
        Return If(mode = ModeNarrativeWU.ParNature, CODE_MODE_PAR_NATURE, CODE_MODE_UNIQUE)
    End Function

    ''' <summary>Ce que le mode dit à l'écran, une fois choisi.</summary>
    Public Shared Function IntituleDeMode(mode As ModeNarrativeWU) As String

        If mode = ModeNarrativeWU.ParNature Then
            Return "un libellé par nature de mouvement"
        End If

        Return "un seul libellé pour toutes les lignes d'un point de vente"
    End Function

#End Region

#Region "Le choix du modèle"

    ''' <summary>
    ''' Le modèle qui s'applique à cette nature, mode compris.
    '''
    ''' L'ÉCART D'ARRONDI EST TRAITÉ À PART, DANS LES DEUX MODES. Il est posé APRÈS la pièce,
    ''' pour absorber la différence globale, et ne se rattache à AUCUN point de vente — le
    ''' repère principal du modèle global. Lui appliquer ce modèle le réduirait à « LD WU
    ''' ACTIVITE », c'est-à-dire à une ligne qui ne dit plus ce qu'elle est. Il garde donc son
    ''' texte propre, période comprise, que la banque peut éditer comme les autres.
    '''
    ''' UNE NATURE LAISSÉE VIDE RETOMBE SUR SON LIBELLÉ HISTORIQUE, et non sur le modèle
    ''' global. C'est le rectificatif de la banque du 06/10/2026 : la narrative unique ne vaut
    ''' que pour la colonne ADDLTEXT du fichier core banking ; la pièce comptable reprend les
    ''' libellés détaillés, et la banque n'a rien à saisir pour les retrouver.
    ''' </summary>
    Public Function Modele(nature As NatureMouvementWU) As String

        ' Une nature que la banque a personnalisée l'emporte toujours, quel que soit le mode :
        ' elle a saisi ce texte-là pour cette ligne-là.
        If _modeles.ContainsKey(nature) Then Return _modeles(nature)

        ' L'ÉCART D'ARRONDI NE SUIT JAMAIS LE MODÈLE GLOBAL. Posé après la pièce, il ne se
        ' rattache à aucun point de vente : le modèle global le réduirait à « LD WU ACTIVITE ».
        If nature = NatureMouvementWU.EcartArrondi Then Return ConstantesWU.NARRATIVE_ECART_DEFAUT

        If _mode = ModeNarrativeWU.ModeleUnique Then Return _modeleGlobal

        Return NaturesMouvementWU.ModeleParDefaut(nature)
    End Function

    ''' <summary>Vrai si cette nature porte un modèle qui lui est propre, saisi par la banque.</summary>
    Public Function EstPersonnalisee(nature As NatureMouvementWU) As Boolean
        Return _modeles.ContainsKey(nature)
    End Function

    ''' <summary>
    ''' Le libellé d'une ligne de cette nature, prêt à être posé dans la pièce.
    '''
    ''' RENDU PAR AppliquerAuLibelle, ET NON PAR Appliquer. Les deux mettent en majuscules —
    ''' la banque les veut sur la pièce comme sur la narrative — mais celui-ci NE RÉSORBE PAS
    ''' LES ESPACES : la pièce manuelle écrit « TTA (TAXE SUR RECEPTION  DE FONDS WU) » avec
    ''' deux espaces, et c'est elle que la Direction Comptable rapproche ligne à ligne.
    ''' </summary>
    Public Function Libelle(nature As NatureMouvementWU, designation As String, account As String,
                            codeAgence As String, periode As String) As String

        Return ModeleNarrativeWU.AppliquerAuLibelle(Modele(nature), designation, account,
                                                    codeAgence, periode)
    End Function

    ''' <summary>
    ''' Les quatorze libellés d'un point de vente, en un seul appel.
    '''
    ''' POURQUOI TOUS D'UN COUP. Les douze lignes d'un même point de vente se posent à la
    ''' suite, et chacune a besoin du sien : les calculer ensemble évite douze résolutions de
    ''' mode et douze appels au moteur de substitution, et surtout rend LISIBLES les douze
    ''' appels de la pièce, qui se contentent d'indexer leur nature.
    ''' </summary>
    Public Function Libelles(designation As String, account As String, codeAgence As String,
                             periode As String) As Dictionary(Of NatureMouvementWU, String)

        Dim rendus As New Dictionary(Of NatureMouvementWU, String)

        For Each nature As NatureMouvementWU In NaturesMouvementWU.Toutes()
            rendus(nature) = Libelle(nature, designation, account, codeAgence, periode)
        Next

        Return rendus
    End Function

#End Region

End Class
