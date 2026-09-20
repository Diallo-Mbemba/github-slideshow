Option Strict On
Option Explicit On

Imports System.Globalization

Imports System.Text

''' <summary>
''' Écriture et lecture de fichiers CSV, sans Excel.
'''
''' POURQUOI DU CSV, ET NON UN CLASSEUR
'''
''' Le paramétrage exporté doit pouvoir être relu sur un poste NEUF, celui d'une installation
''' qui commence — donc éventuellement sans Excel installé, et en tout cas sans qu'on puisse
''' le parier. Excel Interop est exclu ici : il l'est déjà pour toute la logique métier, il
''' l'est a fortiori pour le fichier qui sert à redémarrer.
'''
''' Un CSV en point-virgule et en UTF-8 avec BOM s'ouvre néanmoins d'un double-clic dans Excel,
''' en français, avec les colonnes séparées. On garde donc la lisibilité sans la dépendance.
'''
''' LES DEUX PIÈGES DU CSV, ET COMMENT ILS SONT TRAITÉS
'''
'''   — une valeur qui contient le séparateur, un guillemet ou un saut de ligne : elle est
'''     entourée de guillemets, et ses guillemets internes sont doublés ;
'''   — une valeur qui ressemble à autre chose qu'à du texte : ce n'est pas le problème du
'''     CSV mais de celui qui le relit. Les nombres sont donc écrits en culture INVARIANTE,
'''     point décimal compris, et relus de même. Un taux écrit 0,70 sur un poste français et
'''     relu 0.70 sur un autre vaudrait soixante-dix fois trop.
''' </summary>
Public NotInheritable Class FichierCsvWU

    Private Sub New()
    End Sub

    ''' <summary>
    ''' Le point-virgule, et non la virgule : c'est le séparateur qu'Excel attend d'un CSV en
    ''' configuration française, et il ne se rencontre pas dans un numéro de compte.
    ''' </summary>
    Public Const SEPARATEUR As Char = ";"c

    Private Const GUILLEMET As Char = """"c

    ''' <summary>
    ''' UTF-8 AVEC marque d'ordre des octets. Sans elle, Excel lit le fichier en ANSI et
    ''' « Société » devient « SociÃ©tÃ© » — ce qui n'empêche pas l'application de le relire,
    ''' mais fait douter de tout le reste.
    ''' </summary>
    Public Shared ReadOnly Property Encodage As Encoding
        Get
            Return New UTF8Encoding(True)
        End Get
    End Property

#Region "Écriture"

    ''' <summary>Assemble une ligne CSV à partir de ses valeurs.</summary>
    Public Shared Function Ligne(ParamArray valeurs() As String) As String

        If valeurs Is Nothing OrElse valeurs.Length = 0 Then Return String.Empty

        Dim morceaux(valeurs.Length - 1) As String

        For rang As Integer = 0 To valeurs.Length - 1
            morceaux(rang) = Echapper(valeurs(rang))
        Next

        Return String.Join(SEPARATEUR.ToString(), morceaux)
    End Function

    ''' <summary>
    ''' Protège une valeur. Les guillemets internes sont doublés, comme le veut la convention :
    ''' c'est la seule façon d'écrire un guillemet dans un champ entre guillemets.
    ''' </summary>
    Public Shared Function Echapper(valeur As String) As String

        Dim texte As String = If(valeur, String.Empty)

        Dim aProteger As Boolean =
            texte.IndexOf(SEPARATEUR) >= 0 OrElse
            texte.IndexOf(GUILLEMET) >= 0 OrElse
            texte.IndexOf(ControlChars.Cr) >= 0 OrElse
            texte.IndexOf(ControlChars.Lf) >= 0

        If Not aProteger Then Return texte

        Return GUILLEMET & texte.Replace(GUILLEMET.ToString(), GUILLEMET.ToString() & GUILLEMET.ToString()) & GUILLEMET
    End Function

    ''' <summary>
    ''' Un nombre décimal, en culture invariante. Le taux d'un sous-agent ne doit pas dépendre
    ''' des paramètres régionaux du poste qui l'exporte.
    ''' </summary>
    Public Shared Function Nombre(valeur As Decimal) As String
        Return valeur.ToString("0.00", CultureInfo.InvariantCulture)
    End Function

    ''' <summary>Un booléen, écrit en toutes lettres pour que le fichier reste lisible.</summary>
    Public Shared Function Booleen(valeur As Boolean) As String
        Return If(valeur, "oui", "non")
    End Function

#End Region

#Region "Lecture"

    ''' <summary>
    ''' Découpe une ligne CSV en respectant les guillemets.
    '''
    ''' Écrite caractère par caractère, et non par Split : un Split sur le point-virgule
    ''' couperait en deux une désignation qui en contient un, et personne ne s'en apercevrait
    ''' avant que le sous-agent correspondant soit mal créé.
    ''' </summary>
    Public Shared Function Decouper(texteLigne As String) As List(Of String)

        Dim valeurs As New List(Of String)()
        If texteLigne Is Nothing Then Return valeurs

        Dim courant As New StringBuilder()
        Dim entreGuillemets As Boolean = False
        Dim position As Integer = 0

        While position < texteLigne.Length

            Dim caractere As Char = texteLigne(position)

            If entreGuillemets Then

                If caractere = GUILLEMET Then

                    ' Deux guillemets de suite : c'est un guillemet littéral, pas la fin du champ.
                    If position + 1 < texteLigne.Length AndAlso texteLigne(position + 1) = GUILLEMET Then
                        courant.Append(GUILLEMET)
                        position += 2
                        Continue While
                    End If

                    entreGuillemets = False
                    position += 1
                    Continue While
                End If

                courant.Append(caractere)
                position += 1
                Continue While
            End If

            If caractere = GUILLEMET Then
                entreGuillemets = True
                position += 1
                Continue While
            End If

            If caractere = SEPARATEUR Then
                valeurs.Add(courant.ToString())
                courant.Clear()
                position += 1
                Continue While
            End If

            courant.Append(caractere)
            position += 1
        End While

        valeurs.Add(courant.ToString())
        Return valeurs
    End Function

    ''' <summary>
    ''' Découpe un texte entier en lignes de valeurs, en-tête compris.
    '''
    ''' Les lignes vides sont ignorées : un fichier édité à la main en porte presque toujours
    ''' une à la fin, et elle produirait un point de vente sans nom.
    ''' </summary>
    Public Shared Function Lire(texte As String) As List(Of List(Of String))

        Dim lignes As New List(Of List(Of String))()
        If String.IsNullOrEmpty(texte) Then Return lignes

        For Each texteLigne As String In texte.Split(New String() {ControlChars.CrLf, ControlChars.Lf},
                                                     StringSplitOptions.None)

            If String.IsNullOrWhiteSpace(texteLigne) Then Continue For
            lignes.Add(Decouper(texteLigne))
        Next

        Return lignes
    End Function

    ''' <summary>
    ''' La valeur d'une colonne, ou une chaîne vide si la ligne s'arrête avant.
    ''' Une ligne plus courte que l'en-tête n'est pas une erreur fatale : mieux vaut une
    ''' valeur manquante signalée au contrôle qu'une exception au milieu d'un import.
    ''' </summary>
    Public Shared Function Valeur(valeursLues As List(Of String), rang As Integer) As String

        If valeursLues Is Nothing OrElse rang < 0 OrElse rang >= valeursLues.Count Then Return String.Empty
        Return If(valeursLues(rang), String.Empty).Trim()
    End Function

    ''' <summary>
    ''' Relit un nombre décimal écrit par Nombre(). La culture invariante d'abord, la culture
    ''' du poste ensuite : un fichier corrigé à la main dans Excel repart avec une virgule.
    ''' </summary>
    Public Shared Function EssayerNombre(texte As String, ByRef resultat As Decimal) As Boolean

        resultat = 0D
        Dim brut As String = If(texte, String.Empty).Trim()
        If brut.Length = 0 Then Return False

        If Decimal.TryParse(brut, NumberStyles.Any, CultureInfo.InvariantCulture, resultat) Then Return True
        Return Decimal.TryParse(brut, NumberStyles.Any, CultureInfo.CurrentCulture, resultat)
    End Function

    ''' <summary>Relit un booléen écrit par Booleen(), et tolère les graphies voisines.</summary>
    Public Shared Function EstVrai(texte As String) As Boolean

        Select Case If(texte, String.Empty).Trim().ToUpperInvariant()
            Case "OUI", "VRAI", "TRUE", "1", "O", "X" : Return True
            Case Else : Return False
        End Select
    End Function

#End Region

#Region "Empreinte"

    ''' <summary>
    ''' Empreinte SHA-256 d'un texte, en hexadécimal.
    '''
    ''' Elle ne protège de rien — qui modifie un fichier peut recalculer son empreinte. Elle
    ''' sert à RÉPONDRE À UNE QUESTION : « ce fichier est-il celui qui est sorti de
    ''' l'application, ou quelqu'un l'a-t-il retouché entre-temps ? ». Retoucher est parfois
    ''' légitime ; l'ignorer ne l'est pas.
    ''' </summary>
    ''' <summary>
    ''' Empreinte SHA-256 d'un FICHIER, en hexadécimal. Chaîne vide si le fichier est
    ''' illisible : une empreinte manquante se dit, elle ne fait pas échouer un traitement
    ''' qui a par ailleurs abouti.
    '''
    ''' Le fichier est lu par blocs et non chargé en mémoire : un rapport Western Union
    ''' décompressé pèse plusieurs dizaines de mégaoctets.
    ''' </summary>
    Public Shared Function EmpreinteFichier(chemin As String) As String

        If String.IsNullOrWhiteSpace(chemin) Then Return String.Empty

        Try
            Using flux As IO.FileStream = IO.File.OpenRead(chemin)
                Using algorithme As System.Security.Cryptography.SHA256 = System.Security.Cryptography.SHA256.Create()

                    Dim octets() As Byte = algorithme.ComputeHash(flux)
                    Dim assemblage As New StringBuilder(octets.Length * 2)

                    For Each octet As Byte In octets
                        assemblage.Append(octet.ToString("x2", CultureInfo.InvariantCulture))
                    Next

                    Return assemblage.ToString()
                End Using
            End Using

        Catch ex As IO.IOException
            Return String.Empty
        Catch ex As UnauthorizedAccessException
            Return String.Empty
        End Try
    End Function

    Public Shared Function Empreinte(texte As String) As String

        Using algorithme As System.Security.Cryptography.SHA256 = System.Security.Cryptography.SHA256.Create()

            Dim octets() As Byte = algorithme.ComputeHash(New UTF8Encoding(False).GetBytes(If(texte, String.Empty)))
            Dim assemblage As New StringBuilder(octets.Length * 2)

            For Each octet As Byte In octets
                assemblage.Append(octet.ToString("x2", CultureInfo.InvariantCulture))
            Next

            Return assemblage.ToString()
        End Using
    End Function

#End Region

End Class
