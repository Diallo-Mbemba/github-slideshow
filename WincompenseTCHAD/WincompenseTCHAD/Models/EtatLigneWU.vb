Option Strict On
Option Explicit On

Imports System.Drawing

''' <summary>Niveaux de gravité d'une ligne de la grille de contrôle.</summary>
Public Enum GraviteWU
    ''' <summary>Rien à signaler.</summary>
    Aucune = 0
    ''' <summary>Paramétrage incomplet : la ligne se calcule mal, sans que ce soit une erreur.</summary>
    Incomplet = 1
    ''' <summary>À regarder avant de générer la pièce.</summary>
    AVerifier = 2
    ''' <summary>Bloquant : la pièce comptable ne peut pas être générée sur cette ligne.</summary>
    Bloquant = 3
End Enum

''' <summary>
''' État d'une ligne de la grille de contrôle : une gravité, qui donne la couleur, et un
''' libellé, qui donne la raison.
'''
''' La grille distinguait auparavant cinq cas par cinq couleurs pastel — dont trois jaunâtres
''' et deux rosées — sans aucune légende. Sur un écran de bureau ordinaire, un agent ne les
''' distinguait pas, et rien ne lui disait ce que chacune signifiait.
'''
''' Les cinq cas sont conservés : c'est leur restitution qui change. La couleur ne porte plus
''' que la gravité, sur trois niveaux nettement séparés, et le libellé porte la raison exacte
''' dans une colonne « État ». Rien n'est perdu, et tout est nommé.
''' </summary>
Public Class EtatLigneWU

#Region "Libellés des cas"

    Public Const LIBELLE_ERREUR_SQL As String = "Erreur SQL"
    Public Const LIBELLE_ACCOUNT_INCONNU As String = "Account inconnu"
    Public Const LIBELLE_PARAMETRAGE_ABSENT As String = "Paramétrage absent"
    Public Const LIBELLE_ECART_ANORMAL As String = "Écart anormal"
    Public Const LIBELLE_SOLDE_NON_NUL As String = "Solde non nul"

#End Region

    Public Property Gravite As GraviteWU = GraviteWU.Aucune
    Public Property Libelle As String = String.Empty

    Private Sub New(niveau As GraviteWU, texte As String)
        Gravite = niveau
        Libelle = texte
    End Sub

    ''' <summary>Vrai lorsque la ligne ne présente aucune anomalie.</summary>
    Public ReadOnly Property EstNormale As Boolean
        Get
            Return Gravite = GraviteWU.Aucune
        End Get
    End Property

#Region "Couleurs"

    ''' <summary>Fond de la ligne.</summary>
    Public ReadOnly Property Fond As Color
        Get
            Select Case Gravite
                Case GraviteWU.Bloquant : Return Color.FromArgb(&HFD, &HE7, &HE5)
                Case GraviteWU.AVerifier : Return Color.FromArgb(&HFE, &HF3, &HE2)
                Case GraviteWU.Incomplet : Return Color.FromArgb(&HF2, &HF4, &HF6)
                Case Else : Return Color.White
            End Select
        End Get
    End Property

    ''' <summary>
    ''' Couleur du texte. Seul le cas bloquant la change : ailleurs, un texte coloré sur un fond
    ''' coloré se lit moins bien qu'un texte noir, et la teinte du fond suffit.
    ''' </summary>
    Public ReadOnly Property Encre As Color
        Get
            If Gravite = GraviteWU.Bloquant Then Return Color.FromArgb(&H8A, &H1B, &H12)
            Return Color.FromArgb(&H1B, &H1B, &H1B)
        End Get
    End Property

    ''' <summary>Couleur du filet de gauche, qui sépare les trois niveaux à l'œil.</summary>
    Public ReadOnly Property Filet As Color
        Get
            Select Case Gravite
                Case GraviteWU.Bloquant : Return Color.FromArgb(&HB4, &H23, &H18)
                Case GraviteWU.AVerifier : Return Color.FromArgb(&HB5, &H47, &H8)
                Case GraviteWU.Incomplet : Return Color.FromArgb(&H98, &HA2, &HB3)
                Case Else : Return Color.Transparent
            End Select
        End Get
    End Property

#End Region

#Region "Détermination"

    ''' <summary>
    ''' Détermine l'état d'une ligne. L'ordre d'examen reprend exactement la priorité de gravité
    ''' retenue à l'origine : erreur SQL d'abord, puis Account inconnu, puis données manquantes,
    ''' puis écart d'arrondi anormal, puis solde non nul.
    ''' </summary>
    Public Shared Function Determiner(erreurSql As Boolean, donneesManquantes As Boolean,
                                      type As String, solde As Decimal,
                                      ecartArrondi As Long) As EtatLigneWU

        If erreurSql Then
            Return New EtatLigneWU(GraviteWU.Bloquant, LIBELLE_ERREUR_SQL)
        End If

        If String.Equals(If(type, String.Empty), CalculWU.TYPE_INCONNU, StringComparison.OrdinalIgnoreCase) Then
            Return New EtatLigneWU(GraviteWU.AVerifier, LIBELLE_ACCOUNT_INCONNU)
        End If

        If donneesManquantes Then
            Return New EtatLigneWU(GraviteWU.Incomplet, LIBELLE_PARAMETRAGE_ABSENT)
        End If

        If Math.Abs(ecartArrondi) > ConstantesWU.SEUIL_ECART_LIGNE_ANORMAL Then
            Return New EtatLigneWU(GraviteWU.AVerifier, LIBELLE_ECART_ANORMAL)
        End If

        If solde <> 0D Then
            Return New EtatLigneWU(GraviteWU.AVerifier, LIBELLE_SOLDE_NON_NUL)
        End If

        Return New EtatLigneWU(GraviteWU.Aucune, String.Empty)
    End Function

#End Region

End Class
