Option Strict On
Option Explicit On

Imports System.Security.Cryptography
Imports System.Text

''' <summary>
''' Met un mot de passe SQL Server hors de portée d'une simple lecture du fichier de
''' configuration.
'''
''' CE QUE CELA PROTÈGE, ET CE QUE CELA NE PROTÈGE PAS
'''
''' Le chiffrement employé est celui de Windows (DPAPI), en portée MACHINE : la clé appartient
''' au poste, et n'est jamais écrite nulle part. Il en découle deux choses, qu'il faut annoncer
''' telles quelles plutôt que de laisser croire à une sécurité qu'elles n'apportent pas.
'''
'''   CE QUI EST PROTÉGÉ — le mot de passe n'apparaît plus en clair dans wincompense.config.
'''   Le fichier copié sur une autre machine, ou emporté sur une clé USB, ne donne rien :
'''   lui seul ne suffit pas à déchiffrer.
'''
'''   CE QUI NE L'EST PAS — sur CE poste, un programme lancé par n'importe quel utilisateur
'''   local peut redemander le déchiffrement à Windows. La portée machine est indispensable
'''   ici, le fichier étant commun à tous les comptes du poste ; elle n'oppose donc rien à
'''   quelqu'un qui exécute du code sur la machine elle-même.
'''
''' Autrement dit : cela ferme la lecture accidentelle et le vol de fichier, pas l'accès
''' administrateur au poste. L'authentification Windows reste préférable quand la banque
''' l'accepte, puisqu'aucun secret n'est alors conservé nulle part.
''' </summary>
Public NotInheritable Class SecretWU

    Private Sub New()
    End Sub

    ''' <summary>
    ''' Entropie fixe, mêlée au chiffrement. Elle n'est pas un secret — elle est dans le code —
    ''' et ne prétend pas en être un : son seul rôle est d'empêcher qu'un blob chiffré par une
    ''' autre application de ce poste soit déchiffrable par celle-ci, et réciproquement.
    ''' </summary>
    Private Shared ReadOnly ENTROPIE As Byte() =
        Encoding.UTF8.GetBytes("Wincompense.Connexion.v1")

    ''' <summary>
    ''' Chiffre un texte pour ce poste. Rend une chaîne vide si le texte l'est, ou si Windows
    ''' refuse — l'appelant écrira alors simplement moins, jamais en clair par accident.
    ''' </summary>
    Public Shared Function Proteger(texte As String) As String

        If String.IsNullOrEmpty(texte) Then Return String.Empty

        Try
            Dim chiffre As Byte() = ProtectedData.Protect(
                Encoding.UTF8.GetBytes(texte), ENTROPIE, DataProtectionScope.LocalMachine)

            Return Convert.ToBase64String(chiffre)

        Catch ex As CryptographicException
            ' DPAPI indisponible : mieux vaut ne rien conserver que conserver en clair.
            Return String.Empty
        End Try
    End Function

    ''' <summary>
    ''' Déchiffre ce que <see cref="Proteger"/> a écrit sur CE poste. Rend une chaîne vide si la
    ''' valeur vient d'une autre machine, a été tronquée, ou n'est pas du tout un secret protégé.
    ''' </summary>
    Public Shared Function Lire(protege As String) As String

        If String.IsNullOrWhiteSpace(protege) Then Return String.Empty

        Try
            Dim clair As Byte() = ProtectedData.Unprotect(
                Convert.FromBase64String(protege.Trim()), ENTROPIE, DataProtectionScope.LocalMachine)

            Return Encoding.UTF8.GetString(clair)

        Catch ex As CryptographicException
            ' Fichier recopié depuis un autre poste : le cas est normal, et se corrige en
            ' ressaisissant la chaîne sur cette machine.
            Return String.Empty

        Catch ex As FormatException
            ' Valeur qui n'est pas du Base64 : un fichier modifié à la main, par exemple.
            Return String.Empty
        End Try
    End Function

End Class
