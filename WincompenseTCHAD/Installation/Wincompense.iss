; ============================================================================
;  Wincompense TCHAD - script d'installation Inno Setup
; ============================================================================
;
;  Produit WincompenseTCHAD_Setup.exe, a deployer sur les postes des agents.
;
;  CE QUE FAIT CE PROGRAMME D'INSTALLATION
;    1. verifie que Microsoft .NET Framework 4.8 est present ;
;    2. copie l'application dans Program Files ;
;    3. DEMANDE le serveur SQL et le chemin du fichier partage ;
;    4. ecrit la configuration dans %PROGRAMDATA%\Wincompense, dossier que les
;       utilisateurs peuvent modifier - contrairement a Program Files ;
;    5. cree le fichier partage s'il n'existe pas encore ;
;    6. pose les raccourcis et l'entree de desinstallation.
;
;  POURQUOI LA CONFIGURATION N'EST PAS DANS PROGRAM FILES
;    La banque change souvent de serveur. Une configuration posee a cote de
;    l'executable serait protegee en ecriture, propre a chaque poste, et perdue
;    a la reinstallation : changer de serveur redeviendrait une tournee dans
;    les bureaux. Voir Services\ConfigurationWU.vb.
;
;  POUR CONSTRUIRE LE SETUP
;    1. Visual Studio : compiler la solution en Release (Generer > Configuration
;       Release, puis Generer la solution).
;    2. Installer Inno Setup 6 (gratuit) : https://jrsoftware.org/isdl.php
;    3. Ouvrir ce fichier dans Inno Setup, puis Build > Compile (Ctrl+F9).
;    4. Le setup apparait dans Installation\Sortie\.
; ============================================================================

#define NomApplication      "Wincompense TCHAD"
#define VersionApplication  "1.0.0"
#define Editeur             "Ecobank Tchad"
#define ExecutablePrincipal "WincompenseTCHAD.exe"

; Dossier de compilation Release, relatif a ce script.
#define DossierRelease      "..\WincompenseTCHAD\bin\Release"

[Setup]
; Cet identifiant ne doit JAMAIS changer : c'est par lui que Windows reconnait
; une mise a jour d'une installation existante plutot qu'un second produit.
AppId={{6FCC1274-03A5-5832-9B2D-CB6797179959}
AppName={#NomApplication}
AppVersion={#VersionApplication}
AppVerName={#NomApplication} {#VersionApplication}
AppPublisher={#Editeur}
DefaultDirName={autopf}\{#NomApplication}
DefaultGroupName={#NomApplication}
OutputDir=Sortie
OutputBaseFilename=WincompenseTCHAD_Setup_{#VersionApplication}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern

; L'installation ecrit dans Program Files et cree un dossier sous ProgramData :
; les droits administrateur sont indispensables.
PrivilegesRequired=admin

; L'application est compilee en AnyCPU : elle tourne des deux cotes, et
; s'installe dans le Program Files natif du poste.
; Sur une version d'Inno Setup anterieure a la 6.3, remplacer x64compatible par x64.
ArchitecturesInstallIn64BitMode=x64compatible

; Windows 7 SP1 est le plancher du .NET Framework 4.8.
MinVersion=6.1sp1

UninstallDisplayName={#NomApplication}
UninstallDisplayIcon={app}\{#ExecutablePrincipal}

[Languages]
Name: "francais"; MessagesFile: "compiler:Languages\French.isl"

[Tasks]
Name: "raccourcibureau"; Description: "Creer un raccourci sur le Bureau"; GroupDescription: "Raccourcis :"

[Files]
; L'application. Les fichiers de debogage ne partent pas en production.
Source: "{#DossierRelease}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs; \
    Excludes: "*.pdb,*.xml,*.vshost.*"

; Les scripts SQL accompagnent l'installation : l'informatique les a sous la main
; le jour ou elle monte la base sur un nouveau serveur.
Source: "..\Scripts\*.sql"; DestDir: "{app}\Scripts"; Flags: ignoreversion

; Le modele du fichier partage et le script de changement de serveur.
Source: "connexion.config.modele"; DestDir: "{app}\Installation"; Flags: ignoreversion
Source: "Configurer-Connexion.ps1"; DestDir: "{app}\Installation"; Flags: ignoreversion
Source: "LISEZMOI-Installation.md"; DestDir: "{app}\Installation"; Flags: ignoreversion

[Dirs]
; LE POINT IMPORTANT. Sans "users-modify", l'application ne pourrait pas
; rafraichir sa copie locale de la connexion, et l'administrateur ne pourrait
; pas changer de serveur depuis l'ecran prevu pour cela : ProgramData n'est pas
; modifiable par un utilisateur standard sans cette permission explicite.
Name: "{commonappdata}\Wincompense"; Permissions: users-modify

[Icons]
Name: "{group}\{#NomApplication}"; Filename: "{app}\{#ExecutablePrincipal}"
Name: "{group}\Desinstaller {#NomApplication}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#NomApplication}"; Filename: "{app}\{#ExecutablePrincipal}"; Tasks: raccourcibureau

[Run]
Filename: "{app}\{#ExecutablePrincipal}"; Description: "Lancer {#NomApplication}"; \
    Flags: nowait postinstall skipifsilent

[Code]

var
  PageConnexion: TInputQueryWizardPage;

// ---------------------------------------------------------------------------
//  Prerequis : Microsoft .NET Framework 4.8
// ---------------------------------------------------------------------------
//  Le numero 528040 est la valeur "Release" du 4.8 : toute version egale ou
//  superieure convient. Verifier la presence de la cle ne suffirait pas, elle
//  existe depuis le 4.5.
function DotNet48Present(): Boolean;
var
  Publication: Cardinal;
begin
  Result := RegQueryDWordValue(HKEY_LOCAL_MACHINE,
    'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full', 'Release', Publication)
    and (Publication >= 528040);
end;

function InitializeSetup(): Boolean;
begin
  Result := True;

  if not DotNet48Present() then
  begin
    if MsgBox('Microsoft .NET Framework 4.8 est absent de ce poste.' + #13#10#13#10 +
              'Wincompense TCHAD ne pourra pas demarrer sans lui. Il se telecharge ' +
              'gratuitement chez Microsoft, ou s''installe par Windows Update.' + #13#10#13#10 +
              'Poursuivre malgre tout ?',
              mbConfirmation, MB_YESNO or MB_DEFBUTTON2) <> IDYES then
      Result := False;
  end;
end;

// ---------------------------------------------------------------------------
//  Page supplementaire : serveur SQL et fichier partage
// ---------------------------------------------------------------------------
//  Les valeurs sont demandees A L'INSTALLATION, et non ecrites en dur : c'est
//  tout l'objet de l'exercice.
procedure InitializeWizard();
begin
  PageConnexion := CreateInputQueryPage(wpSelectTasks,
    'Connexion a la base de donnees',
    'Ou Wincompense doit-il chercher ses donnees ?',
    'Ces valeurs pourront etre changees ensuite sans reinstaller, par le menu ' +
    'Securite > Connexion a la base de donnees.');

  PageConnexion.Add('Serveur SQL Server (exemple : SRV-SQL01 ou SRV-SQL01\SQLEXPRESS) :', False);
  PageConnexion.Add('Fichier partage de connexion (exemple : \\SRV-FICHIERS\Wincompense\connexion.config) :', False);

  PageConnexion.Values[0] := '.\SQLEXPRESS';
  PageConnexion.Values[1] := '';
end;

function NextButtonClick(IdPage: Integer): Boolean;
begin
  Result := True;

  if IdPage = PageConnexion.ID then
  begin
    if Trim(PageConnexion.Values[0]) = '' then
    begin
      MsgBox('Indiquez le serveur SQL Server : sans lui, l''application n''a rien a joindre.',
             mbError, MB_OK);
      Result := False;
      Exit;
    end;

    // Un partage vide n'est pas une erreur - le poste travaillera sur sa seule
    // configuration locale - mais c'est presque toujours un oubli.
    if Trim(PageConnexion.Values[1]) = '' then
    begin
      if MsgBox('Aucun fichier partage n''est indique.' + #13#10#13#10 +
                'Ce poste gardera sa propre configuration : le jour ou la banque changera ' +
                'de serveur, il faudra revenir sur cette machine.' + #13#10#13#10 +
                'Continuer sans fichier partage ?',
                mbConfirmation, MB_YESNO or MB_DEFBUTTON2) <> IDYES then
        Result := False;
    end;
  end;
end;

// ---------------------------------------------------------------------------
//  Ecriture de la configuration
// ---------------------------------------------------------------------------
procedure EcrireConfigurationLocale();
var
  Lignes: TArrayOfString;
  Chemin: String;
begin
  Chemin := ExpandConstant('{commonappdata}\Wincompense\wincompense.config');

  SetArrayLength(Lignes, 6);
  Lignes[0] := '# Wincompense TCHAD - configuration de ce poste';
  Lignes[1] := '# Ecrit par le programme d''installation. Une ligne CLE=VALEUR.';
  Lignes[2] := '';
  Lignes[3] := 'PARTAGE=' + Trim(PageConnexion.Values[1]);
  Lignes[4] := 'SERVEUR=' + Trim(PageConnexion.Values[0]);
  Lignes[5] := 'BASE=GWC_WINCOMPENSE_ETD';

  if not SaveStringsToUTF8File(Chemin, Lignes, False) then
    MsgBox('La configuration n''a pas pu etre ecrite dans :' + #13#10 + Chemin + #13#10#13#10 +
           'L''application demarrera sur sa valeur par defaut. Reglez la connexion par le menu ' +
           'Securite > Connexion a la base de donnees.', mbError, MB_OK);
end;

//  Le fichier partage n'est cree que s'il n'existe pas : sur le deuxieme poste
//  installe, il porte deja le reglage de la banque, et l'ecraser avec la saisie
//  d'un technicien ferait basculer tout le monde par accident.
procedure CreerLeFichierPartageSiAbsent();
var
  Lignes: TArrayOfString;
  Chemin: String;
begin
  Chemin := Trim(PageConnexion.Values[1]);
  if Chemin = '' then Exit;
  if FileExists(Chemin) then Exit;

  SetArrayLength(Lignes, 7);
  Lignes[0] := '# Wincompense TCHAD - connexion commune a tous les postes';
  Lignes[1] := '# Changer de serveur : modifier la ligne SERVEUR ci-dessous, puis enregistrer.';
  Lignes[2] := '# Les postes le prendront a leur prochain demarrage.';
  Lignes[3] := '';
  Lignes[4] := 'SERVEUR=' + Trim(PageConnexion.Values[0]);
  Lignes[5] := 'BASE=GWC_WINCOMPENSE_ETD';
  Lignes[6] := 'DELAI=10';

  if not ForceDirectories(ExtractFileDir(Chemin)) then
  begin
    MsgBox('Le dossier du fichier partage n''a pas pu etre cree :' + #13#10 +
           ExtractFileDir(Chemin) + #13#10#13#10 +
           'Ce poste fonctionnera sur sa configuration locale.', mbInformation, MB_OK);
    Exit;
  end;

  if not SaveStringsToUTF8File(Chemin, Lignes, False) then
    MsgBox('Le fichier partage n''a pas pu etre cree :' + #13#10 + Chemin + #13#10#13#10 +
           'Verifiez vos droits d''ecriture sur ce partage. Ce poste fonctionnera sur sa ' +
           'configuration locale.', mbInformation, MB_OK);
end;

procedure CurStepChanged(EtapeCourante: TSetupStep);
begin
  if EtapeCourante = ssPostInstall then
  begin
    EcrireConfigurationLocale();
    CreerLeFichierPartageSiAbsent();
  end;
end;
