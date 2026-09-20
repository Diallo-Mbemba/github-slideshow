Option Strict On
Option Explicit On

''' <summary>
''' Population sur laquelle porte un rapport d'activité.
'''
''' POURQUOI DEUX RAPPORTS, ET NON UN AVEC UN FILTRE
'''
''' Les deux populations n'ont pas les mêmes axes d'analyse, et ce n'est pas une affaire de
''' présentation :
'''
'''   — un SOUS-AGENT appartient à un groupe statistique, porte un taux de rétrocession, et sa
'''     commission se partage entre lui et la banque ;
'''   — une AGENCE PROPRE n'a ni groupe (T_Pdv_EC n'a pas la colonne), ni taux, ni partage : la
'''     banque garde 100 %. Elle appartient en revanche à une AGENCE, qui peut en compter
'''     plusieurs — le code agence n'est pas unique dans T_Pdv_EC.
'''
''' Les mêler obligeait l'onglet « par groupe » à ranger toutes les agences dans une ligne
''' vide, et l'onglet des commissions à additionner une part qui se partage avec une part qui
''' ne se partage pas. Séparer n'ajoute pas une vue : cela en retire une fausse.
''' </summary>
Public Enum PorteeRapportWU

    ''' <summary>Les sous-agents, et les Accounts non encore paramétrés — c'est là qu'on les corrige.</summary>
    SousAgents = 0

    ''' <summary>Le réseau propre de la banque, regroupé par agence.</summary>
    AgencesPropres = 1

End Enum
