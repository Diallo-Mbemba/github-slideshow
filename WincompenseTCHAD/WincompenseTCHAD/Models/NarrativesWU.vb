Option Strict On
Option Explicit On

''' <summary>
''' Comment les libellés d'une pièce sont choisis. La banque bascule d'un mode à l'autre
''' depuis l'écran « Narrative comptable ».
''' </summary>
Public Enum ModeNarrativeWU

    ''' <summary>
    ''' UN SEUL MODÈLE pour les douze lignes d'un point de vente. C'est le mode par défaut, et
    ''' c'est la forme que la banque avait dictée : ce que chaque ligne EST se lit dans son
    ''' numéro de compte, ce qu'elle COUVRE se lit dans le libellé.
    ''' </summary>
    ModeleUnique = 1

    ''' <summary>
    ''' UN MODÈLE PAR NATURE DE MOUVEMENT. Chaque ligne porte le libellé de sa nature, avec
    ''' les mêmes repères — une nature laissée vide retombe sur le modèle global.
    ''' </summary>
    ParNature = 2
End Enum

''' <summary>
''' LE PARAMÉTRAGE DE NARRATIVE EN VIGUEUR : le mode, le modèle global, et les treize modèles
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
''' BASCULER DE MODE NE CHANGE RIEN, ET C'EST VOULU. Les treize natures sont amorcées avec le
''' modèle global : la banque ne voit une différence qu'après avoir édité une nature. Une
''' bascule qui réécrirait treize libellés d'un coup serait un piège — on ne découvre pas au
''' grand livre ce qu'une case à cocher a décidé.
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
    ''' script jamais exécuté. C'est le comportement ACTUEL au caractère près — mode unique,
    ''' modèle par défaut du code — et jamais une pièce sans libellés.
    ''' </summary>
    Public Shared Function ParDefaut() As NarrativesWU
        Return New NarrativesWU(ModeNarrativeWU.ModeleUnique, ModeleNarrativeWU.ModeleParDefaut, Nothing)
    End Function

    ''' <summary>Le mode que désigne la valeur lue en base. Tout ce qui n'est pas franchement
    ''' « par nature » vaut MODE UNIQUE : une valeur mal orthographiée ne doit pas faire
    ''' basculer treize libellés sans que personne ne l'ait demandé.</summary>
    Public Shared Function ModeDepuisCode(valeur As String) As ModeNarrativeWU

        If String.Equals(If(valeur, String.Empty).Trim(), CODE_MODE_PAR_NATURE,
                         StringComparison.OrdinalIgnoreCase) Then
            Return ModeNarrativeWU.ParNature
        End If

        Return ModeNarrativeWU.ModeleUnique
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
    ''' pour absorber la différence globale, et ne se rattache ni à un point de vente ni à une
    ''' période — les deux repères du modèle global. Lui appliquer ce modèle le réduirait à
    ''' « LD WU ACTIVITE », c'est-à-dire à une ligne qui ne dit plus ce qu'elle est. Il garde
    ''' donc son texte propre, que la banque peut éditer comme les autres.
    '''
    ''' UNE NATURE LAISSÉE VIDE RETOMBE SUR LE MODÈLE GLOBAL, et non sur rien : basculer en
    ''' mode « par nature » sans avoir rien saisi doit rendre exactement la pièce d'avant.
    ''' </summary>
    Public Function Modele(nature As NatureMouvementWU) As String

        If nature = NatureMouvementWU.EcartArrondi Then
            If _modeles.ContainsKey(nature) Then Return _modeles(nature)
            Return ConstantesWU.NARRATIVE_ECART_DEFAUT
        End If

        If _mode = ModeNarrativeWU.ModeleUnique Then Return _modeleGlobal

        If _modeles.ContainsKey(nature) Then Return _modeles(nature)
        Return _modeleGlobal
    End Function

    ''' <summary>Vrai si cette nature porte un modèle qui lui est propre, saisi par la banque.</summary>
    Public Function EstPersonnalisee(nature As NatureMouvementWU) As Boolean
        Return _modeles.ContainsKey(nature)
    End Function

    ''' <summary>Le libellé d'une ligne de cette nature, prêt à être posé dans la pièce.</summary>
    Public Function Libelle(nature As NatureMouvementWU, designation As String, account As String,
                            codeAgence As String, periode As String) As String

        Return ModeleNarrativeWU.Appliquer(Modele(nature), designation, account, codeAgence, periode)
    End Function

    ''' <summary>
    ''' Les treize libellés d'un point de vente, en un seul appel.
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
