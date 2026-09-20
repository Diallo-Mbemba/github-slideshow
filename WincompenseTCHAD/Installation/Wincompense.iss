; ============================================================================
;  Wincompense TCHAD - script d'installation Inno Setup
; ============================================================================
;
;  Produit Wincompense_Setup.exe, a deployer sur les postes des agents.
;
;  CE QUE FAIT CE PROGRAMME D'INSTALLATION
;    1. verifie que Microsoft .NET Framework 4.8 est present ;
;    2. copie l'application DANS LE CHEMIN AUTORISE PAR LA BANQUE (voir ci-dessous) ;
;    3. DEMANDE le serveur SQL et le chemin du fichier partage ;
;    4. ecrit la configuration dans %PROGRAMDATA%\Wincompense, dossier que les
;       utilisateurs peuvent modifier - contrairement a Program Files ;
;    5. cree le fichier partage s'il n'existe pas encore ;
;    6. pose les raccourcis et l'entree de desinstallation.
;
;  POURQUOI CE DOSSIER D'INSTALLATION, ET PAS UN AUTRE
;    La securite de la banque n'autorise pas un executable : elle autorise un
;    CHEMIN. Celui qui a ete autorise est
;
;        C:\Program Files\Default Company Name\SetupWincompense\Wincompense.exe
;
;    (l'Explorateur Windows en francais affiche "C:\Programmes\..." : c'est le
;    meme dossier, Windows traduit seulement son nom a l'ecran).
;
;    Ce chemin vient de l'ancien deploiement, fait avec un projet d'installation
;    Visual Studio : "Default Company Name" et "SetupWincompense" sont les valeurs
;    par defaut que ce projet donne a l'editeur et au produit. Elles n'ont aucun
;    sens pour nous - et c'est exactement pour cela qu'il ne faut pas y toucher :
;    les changer demanderait a la banque de refaire son autorisation.
;
;    DefaultDirName reproduit donc ce chemin a l'identique. Ne pas le "corriger"
;    en un nom plus presentable sans une nouvelle autorisation ecrite de la
;    securite : l'application cesserait de demarrer sur les postes.
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
#define Editeur             "Ecobank Tchad"
#define ExecutablePrincipal "Wincompense.exe"

; Le chemin autorise par la securite de la banque, decoupe en ses deux dossiers.
; A ne modifier que sur une nouvelle autorisation ecrite de la securite.
#define EditeurHistorique   "Default Company Name"
#define ProduitHistorique   "SetupWincompense"

; Racine de Program Files.
;   {autopf}   -> C:\Program Files        (affiche "Programmes")
;   {autopf32} -> C:\Program Files (x86)  (affiche "Programmes (x86)")
;
; CONFIRME PAR LA BANQUE : le chemin autorise est
;     C:\Program Files\Default Company Name\SetupWincompense
; soit le Program Files NATIF, sans "(x86)". C'est donc {autopf}, et l'ancien
; projet d'installation Visual Studio n'etait pas en 32 bits. La question est
; tranchee : ne pas basculer sur {autopf32}.
;
; Attention : {autopf} seul ne suffit PAS a designer C:\Program Files. Sur un
; Windows 64 bits, Inno Setup le resout en "Program Files (x86)" tant que
; l'installation n'est pas en mode 64 bits. C'est ArchitecturesInstallIn64BitMode,
; plus bas dans [Setup], qui fait tomber le chemin au bon endroit. Les deux lignes
; tiennent ensemble : retirer l'une sans l'autre change le dossier d'installation,
; et l'application sortirait du chemin autorise par la securite.
#define RacineProgrammes    "{autopf}"

; Dossier de compilation Release.
;
; SourcePath est le dossier de CE script : le chemin obtenu est donc absolu, et ne
; depend pas du dossier courant du compilateur.
;
; Pour compiler depuis un autre emplacement (serveur de construction, dossier
; deplace), passer le chemin en ligne de commande :
;     ISCC.exe /DDossierRelease="C:\chemin\vers\bin\Release" Wincompense.iss
#ifndef DossierRelease
  #define DossierRelease SourcePath + "..\WincompenseTCHAD\bin\Release"
#endif

; Sans ce controle, un projet jamais compile en Release donne un message illisible
; ("No files found matching ..."), qui laisse croire a une erreur du script.
#if !FileExists(DossierRelease + "\" + ExecutablePrincipal)
  #error "Wincompense.exe est introuvable dans bin\Release : la solution n'a pas ete compilee en Release. Dans Visual Studio, choisir Release au lieu de Debug dans la liste de la barre d'outils, puis Generer > Generer la solution. Recompiler ensuite ce script."
#endif

; Version du setup : LUE DANS L'EXECUTABLE, jamais recopiee ici.
;
; Elle etait ecrite a la main, et c'etait une source de derive silencieuse : le nom du
; setup annoncait une version, l'application en portait une autre, et version.txt sur le
; partage comparait ses nombres a celle de l'application. Trois endroits, deux verites.
;
; L'application compare Assembly.GetName().Version (AssemblyVersion de My Project\
; AssemblyInfo.vb) ; GetFileVersion lit AssemblyFileVersion. Visual Studio les tient
; egales par defaut, et elles doivent le rester.
#define VersionApplication GetFileVersion(DossierRelease + "\" + ExecutablePrincipal)

#if VersionApplication == ""
  #define VersionApplication "1.0.0.0"
  #pragma message "Version illisible dans l'executable : 1.0.0.0 est employe par defaut."
#endif

[Setup]
; Cet identifiant ne doit JAMAIS changer : c'est par lui que Windows reconnait
; une mise a jour d'une installation existante plutot qu'un second produit.
AppId={{6FCC1274-03A5-5832-9B2D-CB6797179959}
AppName={#NomApplication}
AppVersion={#VersionApplication}
AppVerName={#NomApplication} {#VersionApplication}
AppPublisher={#Editeur}
; LE POINT DE CE FICHIER. Le chemin autorise par la banque, reproduit tel quel.
DefaultDirName={#RacineProgrammes}\{#EditeurHistorique}\{#ProduitHistorique}
DefaultGroupName={#NomApplication}
OutputDir=Sortie
OutputBaseFilename=Wincompense_Setup_{#VersionApplication}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern

; L'installation ecrit dans Program Files et cree un dossier sous ProgramData :
; les droits administrateur sont indispensables.
PrivilegesRequired=admin

; L'application est compilee en AnyCPU : elle tourne des deux cotes, et
; s'installe dans le Program Files natif du poste.
;
; CETTE LIGNE PORTE LE CHEMIN AUTORISE. Sans elle, {autopf} ci-dessus donnerait
; "C:\Program Files (x86)" sur les postes 64 bits, et l'executable atterrirait
; hors du dossier autorise par la securite de la banque.
;
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
Source: "version.txt.modele"; DestDir: "{app}\Installation"; Flags: ignoreversion
Source: "Configurer-Connexion.ps1"; DestDir: "{app}\Installation"; Flags: ignoreversion
Source: "LISEZMOI-Installation.md"; DestDir: "{app}\Installation"; Flags: ignoreversion

[InstallDelete]
; L'ancien executable occupe le meme dossier que le nouveau : la mise a jour
; precedente l'y a depose sous le nom WincompenseTCHAD.exe. Le laisser la
; reviendrait a garder dans un dossier surveille par la securite un binaire qui,
; lui, n'est pas autorise - et a laisser un utilisateur le lancer par habitude.
Type: files; Name: "{app}\WincompenseTCHAD.exe"
Type: files; Name: "{app}\WincompenseTCHAD.exe.config"
Type: files; Name: "{app}\WincompenseTCHAD.pdb"

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
//  Le fichier de configuration du poste : le relire avant de l'ecrire
// ---------------------------------------------------------------------------
//  Ce fichier ne contient pas que ce que l'installateur y met. L'application y
//  range aussi le MOT DE PASSE du compte SQL, chiffre par Windows pour cette
//  machine, et d'autres valeurs qui lui appartiennent.
//
//  L'ecraser a chaque reinstallation effacerait ce mot de passe, et le poste
//  afficherait « Login failed for user » apres une simple mise a jour - sans que
//  rien n'explique pourquoi, puisque la mise a jour semblait n'avoir touche qu'a
//  l'executable. Les cles que nous n'ecrivons pas sont donc CONSERVEES.
//  NOTE DE VERSION : LoadStringsFromUTF8File demande Inno Setup 6.1 ou superieur, comme
//  SaveStringsToUTF8File employe plus bas. Sur une version anterieure - que ce projet ne
//  vise pas - il faudrait lire par LoadStringsFromFile, au prix des accents.
function CheminConfigLocale(): String;
begin
  Result := ExpandConstant('{commonappdata}\Wincompense\wincompense.config');
end;

// Cle d'une ligne CLE=VALEUR, en majuscules. Chaine vide pour un commentaire,
// une ligne blanche, ou une ligne sans signe egal.
function CleDeLaLigne(Ligne: String): String;
var
  Nette: String;
  Separateur: Integer;
begin
  Result := '';
  Nette := Trim(Ligne);

  if Nette = '' then Exit;
  if Copy(Nette, 1, 1) = '#' then Exit;

  Separateur := Pos('=', Nette);
  if Separateur <= 1 then Exit;

  Result := Uppercase(Trim(Copy(Nette, 1, Separateur - 1)));
end;

// Valeur deja posee sur ce poste pour une cle, ou chaine vide.
function ValeurLocale(Cle: String): String;
var
  Lignes: TArrayOfString;
  Index, Separateur: Integer;
  Nette: String;
begin
  Result := '';

  if not FileExists(CheminConfigLocale()) then Exit;
  if not LoadStringsFromUTF8File(CheminConfigLocale(), Lignes) then Exit;

  for Index := 0 to GetArrayLength(Lignes) - 1 do
  begin
    if CompareText(CleDeLaLigne(Lignes[Index]), Cle) = 0 then
    begin
      Nette := Trim(Lignes[Index]);
      Separateur := Pos('=', Nette);
      Result := Trim(Copy(Nette, Separateur + 1, Length(Nette) - Separateur));
      Exit;
    end;
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

  // Un seul champ pour les deux formes : la banque fournit tantot un nom de serveur,
  // tantot une chaine de connexion complete. Faire choisir entre les deux imposerait a
  // l'installateur de comprendre une distinction qui ne le concerne pas.
  PageConnexion.Add('Serveur SQL Server, OU chaine de connexion complete fournie par la banque :', False);
  PageConnexion.Add('Fichier partage de connexion (exemple : \\SRV-FICHIERS\Wincompense\connexion.config) :', False);

  // Les deux valeurs peuvent venir de la ligne de commande. C'est indispensable au
  // deploiement de masse : en installation silencieuse, la page ne s'affiche pas et
  // personne ne saisit rien, et sans cela, tous les postes repartiraient sur la valeur
  // par defaut, c'est-a-dire sur aucun serveur.
  //
  //   Wincompense_Setup.exe /VERYSILENT ^
  //       /SERVEUR="SRV-SQL01\SQLEXPRESS" ^
  //       /PARTAGE="\\SRV-FICHIERS\Wincompense\connexion.config"
  //
  // /SERVEUR accepte aussi bien un nom de serveur qu'une chaine de connexion complete.
  //
  // SUR UNE REINSTALLATION, les champs reprennent ce que le poste porte deja. Sans cela,
  // une mise a jour silencieuse lancee sans /SERVEUR remettrait le poste sur .\SQLEXPRESS
  // et lui retirerait son partage : l'installateur croirait n'avoir change que
  // l'executable, et le poste ne trouverait plus la base de la banque.
  //
  // L'ordre est : ce qui est passe en ligne de commande, puis ce que le poste porte, puis
  // seulement la valeur par defaut.
  PageConnexion.Values[0] := ExpandConstant('{param:SERVEUR|}');

  if Trim(PageConnexion.Values[0]) = '' then
  begin
    // CHAINE prime sur SERVEUR dans le fichier, comme dans l'application.
    PageConnexion.Values[0] := ValeurLocale('CHAINE');

    if Trim(PageConnexion.Values[0]) = '' then
      PageConnexion.Values[0] := ValeurLocale('SERVEUR');
  end;

  if Trim(PageConnexion.Values[0]) = '' then
    PageConnexion.Values[0] := '.\SQLEXPRESS';

  PageConnexion.Values[1] := ExpandConstant('{param:PARTAGE|}');

  if Trim(PageConnexion.Values[1]) = '' then
    PageConnexion.Values[1] := ValeurLocale('PARTAGE');
end;

// Une chaine de connexion porte toujours au moins un "mot-cle=valeur" ; un nom de
// serveur, jamais. Le signe egal suffit donc a les distinguer, sans rien demander.
function EstUneChaineDeConnexion(Valeur: String): Boolean;
begin
  Result := Pos('=', Valeur) > 0;
end;

// La meme chaine, son mot de passe retire. Sert au fichier partage : il est lisible par
// tous les utilisateurs de l'application, un secret n'y a pas sa place.
//
// Le decoupage est litteral, segment par segment. L'application, elle, relit la chaine avec
// SqlConnectionStringBuilder ; ici on n'a que du texte, et c'est suffisant : il s'agit de
// retirer un mot-cle, pas de comprendre la chaine.
//
// Le decoupage est ecrit a la main plutot qu'avec StringSplitEx : cette fonction demande
// Inno Setup 6.3, et un script qui refuse de compiler sur la version installee a la banque
// couterait plus cher que les six lignes ci-dessous.
function SansMotDePasse(Valeur: String): String;
var
  Reste, Segment, Minuscule: String;
  Separateur: Integer;
begin
  Result := '';
  Reste := Valeur;

  while Reste <> '' do
  begin
    Separateur := Pos(';', Reste);

    if Separateur = 0 then
    begin
      Segment := Trim(Reste);
      Reste := '';
    end
    else
    begin
      Segment := Trim(Copy(Reste, 1, Separateur - 1));
      Reste := Copy(Reste, Separateur + 1, Length(Reste) - Separateur);
    end;

    Minuscule := Lowercase(Segment);

    if (Segment <> '') and (Pos('password', Minuscule) <> 1) and (Pos('pwd', Minuscule) <> 1) then
    begin
      if Result <> '' then Result := Result + ';';
      Result := Result + Segment;
    end;
  end;
end;

function PorteUnMotDePasse(Valeur: String): Boolean;
var
  Minuscules: String;
begin
  Minuscules := Lowercase(Valeur);
  Result := (Pos('password', Minuscules) > 0) or (Pos('pwd', Minuscules) > 0);
end;

// ---------------------------------------------------------------------------
//  Le dossier autorise par la securite
// ---------------------------------------------------------------------------
//  L'assistant laisse changer le dossier d'installation - c'est le comportement
//  attendu d'un installateur, et il faut pouvoir le faire sur un poste de test.
//  Mais en sortir sur un poste de production, c'est installer une application qui
//  ne demarrera pas : la securite de la banque autorise un chemin, pas un produit.
//  D'ou cet avertissement, qui explique au lieu d'interdire.
function CheminAutorise(): String;
begin
  Result := ExpandConstant('{#RacineProgrammes}\{#EditeurHistorique}\{#ProduitHistorique}');
end;

function EstLeCheminAutorise(Chemin: String): Boolean;
begin
  Result := CompareText(RemoveBackslashUnlessRoot(Trim(Chemin)),
                        RemoveBackslashUnlessRoot(CheminAutorise())) = 0;
end;

function NextButtonClick(IdPage: Integer): Boolean;
begin
  Result := True;

  if IdPage = wpSelectDir then
  begin
    if not EstLeCheminAutorise(WizardDirValue) then
      if MsgBox('Ce dossier n''est pas celui que la securite de la banque a autorise.' + #13#10#13#10 +
                'Autorise :' + #13#10 + '    ' + CheminAutorise() + #13#10#13#10 +
                'Choisi :' + #13#10 + '    ' + WizardDirValue + #13#10#13#10 +
                'Installee ailleurs, l''application sera tres probablement bloquee au ' +
                'demarrage sur les postes de la banque.' + #13#10#13#10 +
                'Installer quand meme dans ce dossier ?',
                mbConfirmation, MB_YESNO or MB_DEFBUTTON2) <> IDYES then
        Result := False;
    Exit;
  end;

  if IdPage = PageConnexion.ID then
  begin
    if Trim(PageConnexion.Values[0]) = '' then
    begin
      MsgBox('Indiquez le serveur SQL Server, ou la chaine de connexion fournie par la banque.',
             mbError, MB_OK);
      Result := False;
      Exit;
    end;

    if EstUneChaineDeConnexion(PageConnexion.Values[0]) then
    begin
      if Pos('server=', Lowercase(PageConnexion.Values[0])) = 0 then
        if Pos('data source=', Lowercase(PageConnexion.Values[0])) = 0 then
        begin
          MsgBox('Cette chaine de connexion n''indique aucun serveur.' + #13#10#13#10 +
                 'Elle devrait contenir Server=... ou Data Source=...',
                 mbError, MB_OK);
          Result := False;
          Exit;
        end;

      // Le fichier partage est lisible par TOUS les utilisateurs de l'application : c'est
      // ce qui permet a un changement de serveur de valoir pour tout le monde. Un mot de
      // passe n'y est donc jamais ecrit - il est retire avant, et ne reste que sur ce poste.
      //
      // Ce n'est pas une erreur, c'est une consequence a annoncer : le partage suffira a
      // changer de serveur pour toute la banque, mais pas a donner le mot de passe aux
      // autres postes. Le taire ferait croire a un deploiement termine qui ne l'est pas.
      if PorteUnMotDePasse(PageConnexion.Values[0]) then
        if Trim(PageConnexion.Values[1]) <> '' then
          if MsgBox('Cette chaine contient un mot de passe.' + #13#10#13#10 +
                    'Il restera sur CE poste, chiffre par Windows au premier demarrage. Le ' +
                    'fichier partage ne recevra que le serveur, la base et le nom du compte : ' +
                    'il est lisible par tous les utilisateurs de l''application.' + #13#10#13#10 +
                    'Chaque autre poste devra donc recevoir ce mot de passe une fois, a son ' +
                    'installation. Sans quoi il affichera « Login failed for user ».' + #13#10#13#10 +
                    'Continuer ?',
                    mbConfirmation, MB_YESNO or MB_DEFBUTTON1) <> IDYES then
          begin
            Result := False;
            Exit;
          end;
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
  Anciennes, Lignes: TArrayOfString;
  Chemin, Cle: String;
  Index, Rang: Integer;
begin
  Chemin := CheminConfigLocale();

  // Les cles que nous NE touchons pas sont reprises telles quelles - MOTDEPASSE en
  // premier lieu, chiffre par Windows pour cette machine. Les commentaires, eux, ne sont
  // pas repris : ils seraient recopies a chaque reinstallation et le fichier grossirait
  // d'un en-tete de plus a chaque fois.
  SetArrayLength(Lignes, 0);
  Rang := 0;

  if FileExists(Chemin) then
    if LoadStringsFromUTF8File(Chemin, Anciennes) then
      for Index := 0 to GetArrayLength(Anciennes) - 1 do
      begin
        Cle := CleDeLaLigne(Anciennes[Index]);

        if (Cle <> '') and (Cle <> 'PARTAGE') and (Cle <> 'CHAINE') and
           (Cle <> 'SERVEUR') and (Cle <> 'BASE') then
        begin
          SetArrayLength(Lignes, Rang + 1);
          Lignes[Rang] := Trim(Anciennes[Index]);
          Rang := Rang + 1;
        end;
      end;

  SetArrayLength(Lignes, Rang + 3);
  Lignes[Rang] := '# Wincompense TCHAD - configuration de ce poste';
  Lignes[Rang + 1] := '# Ecrit par le programme d''installation. Une ligne CLE=VALEUR.';
  Lignes[Rang + 2] := 'PARTAGE=' + Trim(PageConnexion.Values[1]);
  Rang := Rang + 3;

  // CHAINE prime sur SERVEUR et BASE : ecrire les deux ferait coexister deux
  // descriptions du meme serveur, dont une seule compte.
  //
  // Le mot de passe, lui, est ecrit ICI tel quel - contrairement au fichier partage. C'est
  // volontaire : l'installateur ne sait pas chiffrer pour Windows, l'application si. Elle
  // le reprend au premier demarrage, le chiffre, et l'efface de sa forme lisible (voir
  // ConfigurationWU.NettoyerLeFichierLocal). La fenetre d'exposition se limite donc a
  // l'intervalle entre l'installation et le premier lancement, sur un fichier de
  // %PROGRAMDATA% et non sur un partage reseau.
  if EstUneChaineDeConnexion(PageConnexion.Values[0]) then
  begin
    SetArrayLength(Lignes, Rang + 1);
    Lignes[Rang] := 'CHAINE=' + Trim(PageConnexion.Values[0]);
  end
  else
  begin
    SetArrayLength(Lignes, Rang + 2);
    Lignes[Rang] := 'SERVEUR=' + Trim(PageConnexion.Values[0]);
    Lignes[Rang + 1] := 'BASE=GWC_WINCOMPENSE_ETD';
  end;

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
  Lignes[1] := '# Changer de serveur : modifier la ligne ci-dessous, puis enregistrer.';
  Lignes[2] := '# Les postes le prendront a leur prochain demarrage.';
  Lignes[3] := '';

  if EstUneChaineDeConnexion(PageConnexion.Values[0]) then
  begin
    SetArrayLength(Lignes, 5);
    Lignes[4] := 'CHAINE=' + SansMotDePasse(Trim(PageConnexion.Values[0]));
  end
  else
  begin
    Lignes[4] := 'SERVEUR=' + Trim(PageConnexion.Values[0]);
    Lignes[5] := 'BASE=GWC_WINCOMPENSE_ETD';
    Lignes[6] := 'DELAI=10';
  end;

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
