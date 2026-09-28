Option Strict On
Option Explicit On

''' <summary>
''' Le barème d'un produit de transfert : ses taxes, leurs taux, leurs assiettes et leurs
''' comptes, dans l'ordre où ils se lisent.
'''
''' POURQUOI CELUI DE WESTERN UNION N'EST PAS EN BASE, ET N'IRA PAS
'''
''' Il serait facile de descendre ces cinq taux dans une table et de les rendre modifiables.
''' La banque a décidé le contraire, et elle a raison : ce barème est RÉCONCILIÉ. Il a été
''' vérifié ligne à ligne contre les pièces manuelles de la banque — quatre-vingt-seize
''' contrôles jour par jour sur AHB020211, puis sur l'agence propre AHB020013 — et la
''' compensation qu'il produit part chaque jour au core banking.
'''
''' Une table se modifie depuis un écran, mais aussi en SQL direct, par quelqu'un qui ne sait
''' pas ce qu'il touche. Le jour où un taux changerait ainsi, la compensation changerait avec
''' lui, sans commit, sans livraison, sans trace ailleurs que dans une ligne de table. On ne
''' met pas une production réconciliée à la portée d'un UPDATE.
'''
''' Ce barème est donc CONSTRUIT À PARTIR DU CODE — ConstantesWU pour les taux,
''' ComptesSystemeWU pour les comptes — et l'écran qui l'affiche est en lecture seule. Le
''' changer suppose de changer le code, de recompiler et de livrer : trois gestes, trois
''' traces.
'''
''' CE QU'IL APPORTE MALGRÉ TOUT, ET CE N'EST PAS RIEN. Aujourd'hui personne à la banque ne
''' peut lire ce barème sans ouvrir Visual Studio : cinq taux, cinq comptes, QUATRE assiettes
''' différentes et un résidu partagé en deux. C'est précisément ce que la banque nous a
''' contesté deux fois ce mois-ci. Le rendre lisible sans le rendre modifiable, c'est tout
''' l'objet de cette classe.
'''
''' LES PRODUITS À VENIR. Ria, puis les suivants, auront leur propre barème — lu en base et
''' modifiable, puisque rien n'y est encore réconcilié. Ils se construiront ici même, par une
''' autre fabrique que <see cref="DeWesternUnion"/>, et <see cref="Modifiable"/> dira à
''' l'écran s'il ouvre la saisie. La structure les attend ; leur table, elle, n'existera
''' qu'avec eux.
''' </summary>
Public NotInheritable Class BaremeTaxesWU

#Region "Contenu"

    ''' <summary>Nom du produit auquel ce barème appartient, tel qu'il s'affiche.</summary>
    Public ReadOnly Property Produit As String

    ''' <summary>
    ''' Vrai si l'écran peut ouvrir la saisie. Faux pour Western Union, et c'est une décision
    ''' de la banque, pas une limite technique.
    ''' </summary>
    Public ReadOnly Property Modifiable As Boolean

    ''' <summary>Les lignes du barème, déjà dans l'ordre d'affichage.</summary>
    Public ReadOnly Property Lignes As List(Of TaxeWU)

    Private Sub New(produit As String, modifiable As Boolean, lignes As List(Of TaxeWU))

        _Produit = produit
        _Modifiable = modifiable
        _Lignes = lignes
    End Sub

#End Region

#Region "Codes des lignes"

    ''' <summary>TVA collectée sur les frais d'envoi.</summary>
    Public Const CODE_TVA As String = "TVA"

    ''' <summary>TTA sur transfert de fonds — l'envoi.</summary>
    Public Const CODE_TTA_ENVOI As String = "TTAE"

    ''' <summary>TTA sur réception de fonds — le paiement.</summary>
    Public Const CODE_TTA_RECEPTION As String = "TTAR"

    ''' <summary>Impôts et taxe sur envoi : 75 % du solde de taxes.</summary>
    Public Const CODE_IMPOTS_ENVOI As String = "ITE"

    ''' <summary>Commission sur transfert : 25 % du solde de taxes.</summary>
    Public Const CODE_COMMISSION_TRANSFERT As String = "CT"

#End Region

#Region "Fabrique"

    ''' <summary>
    ''' Le barème de Western Union, tel que le code l'applique EN CE MOMENT.
    '''
    ''' Les taux sont lus dans ConstantesWU et les comptes dans le paramétrage en service, à
    ''' chaque appel. Rien n'est recopié : si un taux changeait dans les constantes, ou un
    ''' compte dans SystemeWU, l'écran suivrait sans qu'on ait à y penser. Une valeur
    ''' recopiée finit toujours par mentir — l'écran de connexion vient de le rappeler, qui
    ''' annonçait « authentification Windows intégrée » sur des postes connectés par un
    ''' compte SQL.
    ''' </summary>
    Public Shared Function DeWesternUnion() As BaremeTaxesWU

        Dim comptes As ComptesSystemeWU = ComptesSystemeWU.Actuels

        Dim lignes As New List(Of TaxeWU) From {
            New TaxeWU() With {
                .Code = CODE_TVA,
                .Libelle = ConstantesWU.LIB_TVA,
                .Ordre = 1,
                .Taux = ConstantesWU.TAUX_TVA,
                .Assiette = AssietteWU.ChargeEnvoi,
                .Compte = comptes.TVACollectee
            },
            New TaxeWU() With {
                .Code = CODE_TTA_ENVOI,
                .Libelle = ConstantesWU.LIB_TTA_ENVOI,
                .Ordre = 2,
                .Taux = ConstantesWU.TAUX_TTA,
                .Assiette = AssietteWU.PrincipalEnvoi,
                .Compte = comptes.TTAEnvoi
            },
            New TaxeWU() With {
                .Code = CODE_TTA_RECEPTION,
                .Libelle = ConstantesWU.LIB_TTA_RECEPTION,
                .Ordre = 3,
                .Taux = ConstantesWU.TAUX_TTA,
                .Assiette = AssietteWU.PrincipalPaye,
                .Compte = comptes.TTAReception,
                .SousAgentSeulement = True
            },
            New TaxeWU() With {
                .Code = CODE_IMPOTS_ENVOI,
                .Libelle = ConstantesWU.LIB_IMPOTS_TAXE_ENVOI,
                .Ordre = 4,
                .Taux = ConstantesWU.TAUX_TAXE_ENVOI,
                .Assiette = AssietteWU.SoldeTaxes,
                .Compte = comptes.ImpotsTaxeEnvoi
            },
            New TaxeWU() With {
                .Code = CODE_COMMISSION_TRANSFERT,
                .Libelle = ConstantesWU.LIB_COMMISSION_TRANSFERT_BANQUE,
                .Ordre = 5,
                .Taux = ConstantesWU.TAUX_COM_TRANSFERT,
                .Assiette = AssietteWU.SoldeTaxes,
                .Compte = comptes.CommissionTransfertBanque,
                .EstUnProduitBancaire = True
            }
        }

        Return New BaremeTaxesWU("Western Union", False, lignes)
    End Function

#End Region

#Region "Barème en service"

    ''' <summary>
    ''' Le barème du produit en service. Un seul produit existe aujourd'hui, Western Union ;
    ''' quand il y en aura plusieurs, c'est ici que le choix se fera, et nulle part ailleurs.
    ''' </summary>
    Public Shared ReadOnly Property Actuel As BaremeTaxesWU
        Get
            Return DeWesternUnion()
        End Get
    End Property

#End Region

#Region "Lecture"

    ''' <summary>
    ''' La ligne portant ce code, ou Nothing. La comparaison ignore la casse : un code saisi
    ''' à la main dans une table à venir ne doit pas rendre sa taxe introuvable.
    ''' </summary>
    Public Function ParCode(code As String) As TaxeWU

        If String.IsNullOrWhiteSpace(code) Then Return Nothing

        Dim cherche As String = code.Trim()

        For Each ligne As TaxeWU In Lignes
            If String.Equals(ligne.Code, cherche, StringComparison.OrdinalIgnoreCase) Then Return ligne
        Next

        Return Nothing
    End Function

    ''' <summary>
    ''' La formule du solde de taxes, écrite avec les libellés en service.
    '''
    ''' Elle est affichée parce que deux des cinq lignes s'y réfèrent et qu'elles sont
    ''' incompréhensibles sans elle : le solde est un RÉSIDU, et c'est ce qui explique que
    ''' tout franc retiré à la TVA se retrouve réparti 25 / 75 sur les deux autres.
    ''' </summary>
    Public Function FormuleDuSolde() As String

        Return "solde de taxes  =  TOTAL TAXES lu dans le rapport  −  TVA  −  TTA sur envoi"
    End Function

#End Region

End Class
