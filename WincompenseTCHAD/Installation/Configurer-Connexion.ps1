<#
    Wincompense TCHAD - changement de serveur en ligne de commande

    A QUOI SERT CE SCRIPT
        Changer le serveur de toute la banque sans ouvrir l'application. Utile
        a l'informatique, notamment une migration faite hors des heures de
        bureau, quand personne n'est la pour ouvrir Wincompense.

        L'ecran « Securite > Connexion a la base de donnees » fait la meme
        chose, et sait tester la connexion : c'est la voie normale.

    EXEMPLES

        # Changer le serveur pour toute la banque
        .\Configurer-Connexion.ps1 -Serveur SRV-SQL02 `
                                   -Partage \\SRV-FICHIERS\Wincompense\connexion.config

        # Ne regler que le poste courant
        .\Configurer-Connexion.ps1 -Serveur SRV-SQL02 -PosteSeulement

        # Voir la configuration en place, sans rien modifier
        .\Configurer-Connexion.ps1 -Afficher

    A SAVOIR
        Le script TESTE la connexion avant d'ecrire quoi que ce soit. Un
        serveur mal orthographie propage a toute la banque arreterait tout le
        monde, et se corrigerait depuis un poste qui ne se connecte plus.
        L'option -Force passe outre, pour le cas ou le nouveau serveur n'est
        pas encore en ligne au moment de la bascule.
#>

[CmdletBinding()]
param(
    [string] $Serveur,
    [string] $Base = 'GWC_WINCOMPENSE_ETD',
    [int]    $Delai = 10,
    [string] $Partage,
    [switch] $PosteSeulement,
    [switch] $Afficher,
    [switch] $Force
)

$ErrorActionPreference = 'Stop'

$CheminLocal = Join-Path $env:ProgramData 'Wincompense\wincompense.config'

function Lire-Configuration {
    param([string] $Chemin)

    $valeurs = @{}
    if ([string]::IsNullOrWhiteSpace($Chemin) -or -not (Test-Path -LiteralPath $Chemin)) {
        return $valeurs
    }

    foreach ($ligne in Get-Content -LiteralPath $Chemin -Encoding UTF8) {
        $nette = $ligne.Trim()
        if ($nette.Length -eq 0 -or $nette.StartsWith('#')) { continue }

        $separateur = $nette.IndexOf('=')
        if ($separateur -le 0) { continue }

        $cle = $nette.Substring(0, $separateur).Trim()
        $valeurs[$cle] = $nette.Substring($separateur + 1).Trim()
    }

    return $valeurs
}

function Ecrire-Configuration {
    param([string] $Chemin, [hashtable] $Valeurs, [string] $Titre)

    $dossier = Split-Path -Parent $Chemin
    if ($dossier -and -not (Test-Path -LiteralPath $dossier)) {
        New-Item -ItemType Directory -Path $dossier -Force | Out-Null
    }

    $lignes = @(
        "# Wincompense TCHAD - $Titre",
        "# Une ligne CLE=VALEUR. Le caractere # ouvre un commentaire.",
        "# Ecrit le $(Get-Date -Format 'dd/MM/yyyy HH:mm') par $env:USERNAME",
        ''
    )

    foreach ($cle in $Valeurs.Keys | Sort-Object) {
        if (-not [string]::IsNullOrWhiteSpace($Valeurs[$cle])) {
            $lignes += "$cle=$($Valeurs[$cle])"
        }
    }

    Set-Content -LiteralPath $Chemin -Value $lignes -Encoding UTF8
}

function Tester-Connexion {
    param([string] $Serveur, [string] $Base, [int] $Delai)

    $chaine = "Server=$Serveur;Database=$Base;Integrated Security=True;Connect Timeout=$Delai;"
    $connexion = New-Object System.Data.SqlClient.SqlConnection $chaine

    try {
        $connexion.Open()
        $commande = $connexion.CreateCommand()
        $commande.CommandText = 'SELECT DB_NAME()'
        $nom = $commande.ExecuteScalar()
        Write-Host "  Connexion reussie - base $nom" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "  Connexion impossible : $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
    finally {
        $connexion.Dispose()
    }
}

# --- Affichage seul --------------------------------------------------------

$local = Lire-Configuration $CheminLocal

if ($Afficher) {
    Write-Host ''
    Write-Host "Configuration de ce poste : $CheminLocal"
    if ($local.Count -eq 0) {
        Write-Host '  (aucune - ce poste n''a pas ete installe, ou le fichier a ete supprime)'
    } else {
        $local.GetEnumerator() | Sort-Object Name | ForEach-Object {
            Write-Host ("  {0,-10} {1}" -f $_.Name, $_.Value)
        }
    }

    $cheminPartage = $local['PARTAGE']
    Write-Host ''
    if ([string]::IsNullOrWhiteSpace($cheminPartage)) {
        Write-Host 'Aucun fichier partage : ce poste est autonome.'
    } else {
        Write-Host "Fichier partage : $cheminPartage"
        $distant = Lire-Configuration $cheminPartage
        if ($distant.Count -eq 0) {
            Write-Host '  (injoignable ou vide - le poste travaille sur sa copie locale)' -ForegroundColor Yellow
        } else {
            $distant.GetEnumerator() | Sort-Object Name | ForEach-Object {
                Write-Host ("  {0,-10} {1}" -f $_.Name, $_.Value)
            }
        }
    }
    Write-Host ''
    return
}

# --- Changement ------------------------------------------------------------

if ([string]::IsNullOrWhiteSpace($Serveur)) {
    throw 'Indiquez -Serveur, ou -Afficher pour consulter la configuration en place.'
}

if ([string]::IsNullOrWhiteSpace($Partage)) { $Partage = $local['PARTAGE'] }

Write-Host ''
Write-Host "Test de $Serveur ..."

if (-not (Tester-Connexion $Serveur $Base $Delai)) {
    if (-not $Force) {
        throw 'Rien n''a ete modifie. Relancez avec -Force si le serveur n''est pas encore en ligne.'
    }
    Write-Host '  -Force : on ecrit malgre l''echec du test.' -ForegroundColor Yellow
}

$valeurs = @{ SERVEUR = $Serveur; BASE = $Base; DELAI = "$Delai" }

# Le partage d'abord : c'est lui qui peut echouer faute de droits, et mieux vaut
# n'avoir rien change que d'avoir un poste regle sur un serveur ignore des autres.
if (-not $PosteSeulement) {
    if ([string]::IsNullOrWhiteSpace($Partage)) {
        throw 'Aucun fichier partage connu. Indiquez -Partage, ou -PosteSeulement.'
    }

    Ecrire-Configuration $Partage $valeurs 'connexion commune a tous les postes'
    Write-Host "  Fichier partage mis a jour : $Partage" -ForegroundColor Green
}

$valeursLocales = $valeurs.Clone()
$valeursLocales['PARTAGE'] = $Partage
Ecrire-Configuration $CheminLocal $valeursLocales 'configuration de ce poste'
Write-Host "  Poste courant mis a jour : $CheminLocal" -ForegroundColor Green

Write-Host ''
if ($PosteSeulement) {
    Write-Host 'Seul ce poste est change. Les autres lisent toujours l''ancien serveur.' -ForegroundColor Yellow
} else {
    Write-Host 'Les postes prendront le nouveau serveur a leur prochain demarrage.'
}
Write-Host ''
