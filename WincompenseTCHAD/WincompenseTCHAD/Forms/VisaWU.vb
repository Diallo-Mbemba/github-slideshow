Option Strict On
Option Explicit On

Imports System.Windows.Forms

''' <summary>
''' La règle du visa devant le fichier core banking, écrite UNE FOIS.
'''
''' Deux écrans produisent ce fichier : celui de la compense, pour la journée qu'on vient de
''' traiter, et celui des pièces conservées, pour une journée ancienne. Les deux doivent
''' appliquer la même règle — et deux copies d'une règle finissent toujours par diverger.
'''
''' CE QUE LA RÈGLE DIT
'''
'''   — journée visée : rien ne s'oppose à la production ;
'''   — journée non visée, option ACTIVE : la production est refusée, et l'écran dit qui doit
'''     viser et où ;
'''   — journée non visée, option INACTIVE : la production est proposée, avec un avertissement
'''     qui propose « Non » par défaut.
'''
''' L'option est choisie par la banque, pas par l'agent : elle vit en base (T_ParametreWU) et
''' vaut pour tout le monde.
'''
''' EN CAS DE DOUTE, ON LAISSE PASSER
'''
''' En-tête de traitement absent, table des options absente, base injoignable : le fichier se
''' produit, avec un avertissement. Le fichier core banking est l'opération qui met la journée
''' en comptabilité ; la bloquer parce qu'une lecture accessoire a échoué coûterait plus cher
''' que le contrôle ne rapporte.
''' </summary>
Public NotInheritable Class VisaWU

    Private Sub New()
    End Sub

    ''' <summary>
    ''' Autorise, ou non, la production du fichier core banking d'une journée.
    ''' </summary>
    ''' <param name="proprietaire">Fenêtre devant laquelle afficher les messages.</param>
    ''' <returns>Vrai si la production peut continuer.</returns>
    Public Shared Function AutoriserLeCoreBanking(proprietaire As IWin32Window, jour As Date) As Boolean

        Dim messageErreur As String = String.Empty
        Dim traitement As TraitementJourneeWU = TraitementRepository.Charger(jour, messageErreur)

        ' Rien à contrôler : pas d'en-tête, donc pas de visa possible — la journée est
        ' antérieure au dispositif, ou la table n'existe pas encore.
        If traitement Is Nothing OrElse Not traitement.Enregistre Then Return True

        If traitement.EstVisee Then Return True

        If OptionsWU.VisaAvantCoreBanking Then
            Refuser(proprietaire, jour, traitement)
            Return False
        End If

        Return Avertir(proprietaire, jour, traitement)
    End Function

    ''' <summary>
    ''' Refuse, et dit quoi faire. Un refus qui n'indique pas la sortie transforme un contrôle
    ''' en obstacle.
    ''' </summary>
    Private Shared Sub Refuser(proprietaire As IWin32Window, jour As Date,
                               traitement As TraitementJourneeWU)

        MessageBox.Show(proprietaire,
            $"La journée du {jour:dd/MM/yyyy} n'a pas encore été visée." & Environment.NewLine &
            Environment.NewLine &
            "La banque a choisi que le fichier destiné au core banking ne puisse pas être " &
            "produit avant le visa : ce fichier passe les écritures, et le visa atteste qu'un " &
            "supérieur les a relues." & Environment.NewLine & Environment.NewLine &
            "Un utilisateur ayant la fonction « authorizer » doit ouvrir le bordereau de cette " &
            "journée et la viser — écran « Pièces comptables conservées », bouton « Bordereau " &
            "de la journée… »." & Environment.NewLine & Environment.NewLine &
            Auteur(traitement),
            "Journée non visée", MessageBoxButtons.OK, MessageBoxIcon.Warning)
    End Sub

    ''' <summary>
    ''' Avertit, et laisse décider. « Non » est proposé par défaut : produire le fichier
    ''' engage la comptabilité, et le geste par défaut ne doit pas être celui qui engage.
    ''' </summary>
    Private Shared Function Avertir(proprietaire As IWin32Window, jour As Date,
                                    traitement As TraitementJourneeWU) As Boolean

        Dim alerte As String = String.Empty

        If traitement.NombreEcartes > 0 Then
            alerte = Environment.NewLine & Environment.NewLine &
                     $"De plus, {traitement.NombreEcartes} Account(s) de cette journée " &
                     "N'ONT PAS ÉTÉ COMPTABILISÉS : ils ne figurent ni sur la pièce ni dans " &
                     "ce fichier."
        End If

        Return MessageBox.Show(proprietaire,
            $"La journée du {jour:dd/MM/yyyy} n'a pas encore été visée par un supérieur." &
            alerte & Environment.NewLine & Environment.NewLine &
            "Ce fichier passe les écritures dans les livres de la banque : une fois injecté, " &
            "il ne se retire pas depuis Wincompense." & Environment.NewLine & Environment.NewLine &
            Auteur(traitement) & Environment.NewLine & Environment.NewLine &
            "Produire le fichier tout de même ?",
            "Journée non visée", MessageBoxButtons.YesNo, MessageBoxIcon.Warning,
            MessageBoxDefaultButton.Button2) = DialogResult.Yes
    End Function

    ''' <summary>Qui a comptabilisé la journée, pour savoir à qui s'adresser.</summary>
    Private Shared Function Auteur(traitement As TraitementJourneeWU) As String

        If traitement.ComptabilisePar.Length = 0 Then Return "Journée comptabilisée par un auteur inconnu."

        Return $"Journée comptabilisée par {traitement.ComptabilisePar}" &
               If(traitement.DateComptabilisation.HasValue,
                  $" le {traitement.DateComptabilisation.Value:dd/MM/yyyy à HH:mm}", String.Empty) & "."
    End Function

End Class
