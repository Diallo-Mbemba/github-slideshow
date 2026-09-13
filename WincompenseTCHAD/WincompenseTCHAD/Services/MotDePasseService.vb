Option Strict On
Option Explicit On

Imports System.Security.Cryptography
Imports System.Text

''' <summary>
''' Hachage et vérification des mots de passe.
'''
''' Aucun mot de passe n'est jamais enregistré : seule une empreinte PBKDF2 l'est, accompagnée
''' de son sel et de son nombre d'itérations. Un mot de passe perdu ne se retrouve donc pas ;
''' il se réinitialise, ce qui est exactement le comportement attendu.
'''
''' Le nombre d'itérations est enregistré AVEC chaque empreinte plutôt que fixé une fois pour
''' toutes : il pourra être relevé quand les machines seront plus rapides, sans invalider les
''' comptes existants — chacun se vérifie avec le nombre qui a servi à le créer.
'''
''' PBKDF2 est retenu parce qu'il est fourni par le framework (Rfc2898DeriveBytes) et ne
''' demande aucune bibliothèque supplémentaire, contrainte posée dès l'origine du projet.
''' </summary>
Public NotInheritable Class MotDePasseService

    Private Sub New()
    End Sub

    ''' <summary>Nombre d'itérations appliqué aux nouveaux mots de passe.</summary>
    Public Const ITERATIONS_PAR_DEFAUT As Integer = 100000

    ''' <summary>Longueur du sel, en octets.</summary>
    Private Const LONGUEUR_SEL As Integer = 16

    ''' <summary>Longueur de l'empreinte, en octets.</summary>
    Private Const LONGUEUR_EMPREINTE As Integer = 32

    ''' <summary>Longueur minimale exigée d'un mot de passe.</summary>
    Public Const LONGUEUR_MINIMALE As Integer = 8

    ''' <summary>
    ''' Calcule l'empreinte d'un mot de passe et produit le sel qui lui correspond.
    ''' </summary>
    ''' <param name="sel">Sel généré, en Base64. À conserver avec l'empreinte.</param>
    ''' <param name="iterations">Nombre d'itérations appliqué. À conserver également.</param>
    ''' <returns>Empreinte en Base64.</returns>
    Public Shared Function Hacher(motDePasse As String, ByRef sel As String, ByRef iterations As Integer) As String

        iterations = ITERATIONS_PAR_DEFAUT

        Dim octetsSel(LONGUEUR_SEL - 1) As Byte

        ' Générateur cryptographique : un Random ordinaire produirait des sels prévisibles.
        Using generateur As RandomNumberGenerator = RandomNumberGenerator.Create()
            generateur.GetBytes(octetsSel)
        End Using

        sel = Convert.ToBase64String(octetsSel)

        Return CalculerEmpreinte(motDePasse, octetsSel, iterations)
    End Function

    ''' <summary>
    ''' Vérifie qu'un mot de passe correspond à une empreinte.
    ''' Retourne False sur toute donnée inexploitable plutôt que de lever : une empreinte
    ''' corrompue doit refuser la connexion, pas faire tomber l'application.
    ''' </summary>
    Public Shared Function Verifier(motDePasse As String, empreinte As String,
                                    sel As String, iterations As Integer) As Boolean

        If String.IsNullOrEmpty(motDePasse) OrElse String.IsNullOrEmpty(empreinte) OrElse
           String.IsNullOrEmpty(sel) OrElse iterations <= 0 Then
            Return False
        End If

        Try
            Dim octetsSel As Byte() = Convert.FromBase64String(sel)
            Dim calculee As String = CalculerEmpreinte(motDePasse, octetsSel, iterations)

            Return ComparerEnTempsConstant(calculee, empreinte)

        Catch ex As FormatException
            ' Sel ou empreinte non décodables : compte inutilisable, connexion refusée.
            Return False
        End Try
    End Function

    Private Shared Function CalculerEmpreinte(motDePasse As String, sel As Byte(), iterations As Integer) As String

        Using derivation As New Rfc2898DeriveBytes(motDePasse, sel, iterations)
            Return Convert.ToBase64String(derivation.GetBytes(LONGUEUR_EMPREINTE))
        End Using
    End Function

    ''' <summary>
    ''' Compare deux empreintes en parcourant toujours toute leur longueur.
    '''
    ''' Une comparaison ordinaire s'arrête au premier caractère différent : le temps de réponse
    ''' renseignerait alors sur le nombre de caractères déjà corrects. La précaution est ténue
    ''' pour une application de poste de travail, mais elle ne coûte rien.
    ''' </summary>
    Private Shared Function ComparerEnTempsConstant(gauche As String, droite As String) As Boolean

        If gauche Is Nothing OrElse droite Is Nothing Then Return False
        If gauche.Length <> droite.Length Then Return False

        Dim difference As Integer = 0

        For index As Integer = 0 To gauche.Length - 1
            difference = difference Or (AscW(gauche(index)) Xor AscW(droite(index)))
        Next

        Return difference = 0
    End Function

    ''' <summary>
    ''' Tire un mot de passe provisoire, pour une réinitialisation par l'administrateur.
    '''
    ''' Il est tiré au hasard plutôt que saisi : un administrateur qui choisit lui-même les mots
    ''' de passe des autres finit par leur donner toujours le même. Les caractères ambigus
    ''' (O et 0, I, l et 1) sont écartés, parce que ce mot de passe sera lu puis recopié à la main.
    '''
    ''' Le tirage vient du générateur cryptographique et non de Random : un mot de passe
    ''' prévisible, même provisoire, laisse la porte ouverte pendant tout l'intervalle qui
    ''' précède son changement. Le biais introduit par le modulo est négligeable ici, les
    ''' alphabets étant petits devant 256.
    ''' </summary>
    Public Shared Function GenererProvisoire() As String

        Const MAJUSCULES As String = "ABCDEFGHJKLMNPQRSTUVWXYZ"
        Const MINUSCULES As String = "abcdefghijkmnpqrstuvwxyz"
        Const CHIFFRES As String = "23456789"
        Const LONGUEUR As Integer = 12

        Dim alphabet As String = MAJUSCULES & MINUSCULES & CHIFFRES
        Dim caracteres As New List(Of Char)

        Using generateur As RandomNumberGenerator = RandomNumberGenerator.Create()

            ' Les trois natures exigées sont posées d'abord : un tirage entièrement libre
            ' produirait de temps en temps un mot de passe que nos propres règles refusent.
            caracteres.Add(Tirer(generateur, MAJUSCULES))
            caracteres.Add(Tirer(generateur, MINUSCULES))
            caracteres.Add(Tirer(generateur, CHIFFRES))

            While caracteres.Count < LONGUEUR
                caracteres.Add(Tirer(generateur, alphabet))
            End While

            ' Sans mélange, les trois premières positions auraient toujours la même nature.
            For position As Integer = caracteres.Count - 1 To 1 Step -1
                Dim cible As Integer = TirerEntier(generateur, position + 1)
                Dim memoire As Char = caracteres(position)
                caracteres(position) = caracteres(cible)
                caracteres(cible) = memoire
            Next
        End Using

        Return New String(caracteres.ToArray())
    End Function

    ''' <summary>Tire un caractère de l'alphabet donné.</summary>
    Private Shared Function Tirer(generateur As RandomNumberGenerator, alphabet As String) As Char
        Return alphabet(TirerEntier(generateur, alphabet.Length))
    End Function

    ''' <summary>Tire un entier compris entre 0 inclus et la borne exclue.</summary>
    Private Shared Function TirerEntier(generateur As RandomNumberGenerator, borne As Integer) As Integer

        Dim octet As Byte() = New Byte(0) {}
        generateur.GetBytes(octet)
        Return CInt(octet(0)) Mod borne
    End Function

    ''' <summary>
    ''' Contrôle de robustesse d'un mot de passe. Retourne les anomalies constatées — liste vide
    ''' si le mot de passe est acceptable.
    '''
    ''' La règle est volontairement simple : une longueur minimale et trois natures de
    ''' caractères. Empiler les contraintes pousse les utilisateurs à noter leur mot de passe,
    ''' ce qui dégrade la sécurité au lieu de l'améliorer.
    ''' </summary>
    Public Shared Function Anomalies(motDePasse As String, identifiant As String) As List(Of String)

        Dim messages As New List(Of String)

        If String.IsNullOrEmpty(motDePasse) Then
            messages.Add("Le mot de passe est obligatoire.")
            Return messages
        End If

        If motDePasse.Length < LONGUEUR_MINIMALE Then
            messages.Add($"Le mot de passe doit comporter au moins {LONGUEUR_MINIMALE} caractères.")
        End If

        Dim majuscule As Boolean = False
        Dim minuscule As Boolean = False
        Dim chiffre As Boolean = False

        For Each caractere As Char In motDePasse
            If Char.IsUpper(caractere) Then majuscule = True
            If Char.IsLower(caractere) Then minuscule = True
            If Char.IsDigit(caractere) Then chiffre = True
        Next

        If Not (majuscule AndAlso minuscule AndAlso chiffre) Then
            messages.Add("Le mot de passe doit contenir au moins une majuscule, une minuscule et un chiffre.")
        End If

        If Not String.IsNullOrWhiteSpace(identifiant) AndAlso
           motDePasse.IndexOf(identifiant, StringComparison.OrdinalIgnoreCase) >= 0 Then
            messages.Add("Le mot de passe ne doit pas contenir l'identifiant.")
        End If

        Return messages
    End Function

End Class
