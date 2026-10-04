; mOffice - Script do instalador (Inno Setup 6)
;
; Marco 1 / Plano 5 / Task 3. Derivado do tools/installer/mpdf.iss do mPDF (o exemplar vence em sintaxe
; e mecanica). Instala a suite inteira (mTexto, mPlanilha, mSlides), o motor ONLYOFFICE enxuto e o
; runtime WebView2 fixo, no layout que tools/installer/montar-pacote.ps1 monta (Task 1):
;
;   {app}\                   os tres .exe + runtime .NET 10 self-contained (UMA copia)
;   {app}\motor\             converter (x2t, allfontsgen), editors (sem sdkjs-plugins), fonts, dictionaries
;   {app}\WebView2Runtime\   runtime fixo enxuto
;   {app}\licencas\          docs\licencas (menos o rascunho de e-mail)
;
; Parametros (via /D, normalmente passados pelo tools/installer/build-installer.ps1):
;   MyAppVersion - OBRIGATORIO. <Version> de src/Directory.Build.props (fonte UNICA - nao repetir aqui).
;   PacoteDir    - OBRIGATORIO. Pasta gerada por montar-pacote.ps1 -Saida <pasta>.
;
; Exemplo de chamada manual:
;   iscc.exe /DMyAppVersion=1.0.0 /DPacoteDir=C:\...\moffice-pacote moffice.iss
;
; NAO ha senha nem criptografia (projeto de codigo aberto, AGPL-3.0; o mPDF tambem deixou de usar).

#ifndef MyAppVersion
  #error "MyAppVersion nao definido. Use /DMyAppVersion=<versao> (build-installer.ps1 faz isso automaticamente)."
#endif

#ifndef PacoteDir
  #error "PacoteDir nao definido. Use /DPacoteDir=<pasta do montar-pacote.ps1> (build-installer.ps1 faz isso automaticamente)."
#endif

; Termos exibidos antes da instalacao (texto da Task 4: cobre a clausula 2(b)(ii) da licenca do runtime
; WebView2 - backlog 19). Sem o arquivo NAO se gera instalador: nada de texto provisorio no lugar de um
; texto juridico.
#define ArquivoAviso "..\..\docs\licencas\AVISO-mOffice.txt"
#if !FileExists(AddBackslash(SourcePath) + ArquivoAviso)
  #error "docs\licencas\AVISO-mOffice.txt nao existe. Ele e o texto de termos do instalador (Task 4 do Plano 5) e precisa existir antes de gerar o instalador."
#endif

#define MyAppName "mOffice"
#define MyAppPublisher "Projeto mOffice"

; Mutex de PRESENCA dos apps (NomesDeInstancia.MutexDeAtualizacao em src/mOffice.Comum - FONTE UNICA; um
; teste estrutural (InstaladorTests) garante que este texto == a constante). NAO editar um lado sem o outro.
#define MutexDeAtualizacao "Global\mOffice-Update-Mutex-73C377DD"

[Setup]
; GUID fixo do produto mOffice, gerado UMA vez (2026-09-27). Mantem a identidade entre versoes (atualizar
; por cima, desinstalar). NUNCA regenerar.
AppId={{BEA3C718-DFC3-42D6-BF84-41C61CCAC772}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppVerName={#MyAppName} {#MyAppVersion}
DefaultDirName={autopf}\mOffice
DisableProgramGroupPage=yes
; Assistente enxuto (como o mPDF): sem boas-vindas e sem "pronto para instalar". A pagina de pasta e a de
; tarefas continuam. No /VERYSILENT (RMM) nenhuma pagina aparece; no /SILENT (botao "Verificar atualizacoes")
; so a janela de progresso, sem perguntas.
DisableWelcomePage=yes
DisableReadyPage=yes
UninstallDisplayIcon={app}\mTexto.exe
UninstallDisplayName={#MyAppName}
OutputDir=Output
OutputBaseFilename=mOffice-Setup-{#MyAppVersion}
; Compactacao maxima num instalador UNICO (decisao do usuario, 2026-09-27: "instalador unico, compactacao
; maxima, medir"). ~1,47 GB sem compactar. DiskSpanning so e exigido pelo Inno acima de 4.200.000.000 bytes
; COMPACTADOS; o limite de um asset de release do GitHub e 2 GiB. Tamanho medido: ver o relatorio da Task 3.
Compression=lzma2/ultra64
SolidCompression=yes
DiskSpanning=no
; O compactador roda num processo 64-bit proprio (islzma64.exe): o dicionario de 64 MB do ultra64 sobre
; 1,4 GB nao precisa disputar o espaco de enderecamento do ISCC (32-bit).
LZMAUseSeparateProcess=yes
WizardStyle=modern
SetupIconFile=..\..\src\mTexto\Assets\mtexto.ico
; Instalacao para o computador inteiro (Plano 1): Program Files, HKLM, regra de firewall e o catalogo de
; fontes por maquina exigem administrador.
PrivilegesRequired=admin
; Publicacao self-contained win-x64 e runtime WebView2 x64: so Windows 10+ x64.
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
MinVersion=10.0
InfoBeforeFile={#ArquivoAviso}
; Log SEMPRE (%TEMP%\Setup Log *.txt; o do desinstalador tambem): falha do firewall, do catalogo ou da
; checagem de link fica registrada mesmo na atualizacao silenciosa.
SetupLogging=yes
UninstallLogging=yes
; Mitigacao do Windows contra juncao/symlink criado por usuario comum (Windows 10 22H2+/11), no processo do
; instalador e do desinstalador. O padrao do Inno ja e yes; explicito porque o instalador escreve em
; {commonappdata}. Nao vale para os processos filhos (icacls, mTexto /gerarcatalogo) nem em Windows mais
; antigo - por isso a checagem de link do [Code].
RedirectionGuard=yes
; Avisa o Explorer das associacoes (task "associar") - atualiza icones e o "Abrir com" sem reiniciar.
ChangesAssociations=yes

; --- Atualizacao silenciosa (Plano 5, Task 2) ------------------------------------------------------------
; AppMutex: o mutex Global\ que cada app da suite SEGURA pela vida inteira. Vem de uma funcao [Code]
; (MutexDoApp, abaixo) e nao do texto fixo, por causa de um comportamento MEDIDO do Inno (Task 3,
; 2026-09-27, instalador de teste com o mutex segurado): em /VERYSILENT /SUPPRESSMSGBOXES a pergunta
; "o app esta aberto, feche e clique OK" e respondida com Cancelar ("Defaulting to Cancel for suppressed
; message box") e o instalador ABORTA com codigo 1, em silencio - e e exatamente assim que o botao
; "Verificar atualizacoes" (/SILENT, com progresso) e o /update do RMM (/VERYSILENT) chamam o instalador.
; WizardSilent e True nos DOIS modos, entao a regra abaixo cobre os dois. Por isso:
;   - interativo: MutexDoApp devolve o mutex; quem instala a mao ve a pergunta e fecha os apps;
;   - silencioso: MutexDoApp devolve vazio (sem checagem) e quem fecha os apps e o CloseApplications,
;     pelo Restart Manager: "If set to yes ... and Setup is running silently, Setup will always close ...
;     such applications" (ajuda do Inno). O app fecha pelo encerramento de sessao do Windows, que guarda os
;     documentos nao salvos para recuperacao (AplicativoOffice.OnSessionEnding / JanelaEditor.Closed);
;   - desinstalador: sempre checa (desinstalar com o app aberto deixaria arquivos para tras). Como ele usa o
;     AppMutex gravado na instalacao (vazio, se ela foi silenciosa), a checagem e a do InitializeUninstall.
AppMutex={code:MutexDoApp}
CloseApplications=yes
; O relaunch e da secao [Run] (runasoriginaluser, DES-elevado), NAO do Restart Manager (que reabriria o app
; elevado, como o instalador).
RestartApplications=no

[Languages]
Name: "brazilianportuguese"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"

[Tasks]
Name: "associar"; Description: "Registrar mTexto, mPlanilha e mSlides como opção para abrir documentos (.docx, .odt, .xlsx, .ods, .pptx, .odp)"; GroupDescription: "Associação de arquivos:"
Name: "firewall"; Description: "Bloquear no Firewall do Windows o acesso à internet do navegador interno (recomendado)"; GroupDescription: "Segurança:"
Name: "atalhodesktop"; Description: "Criar atalhos na Área de Trabalho"; GroupDescription: "Atalhos adicionais:"; Flags: unchecked

[InstallDelete]
; Atualizacao por cima: o motor e o runtime sao trocados INTEIROS. Sem isto, um arquivo que a versao nova
; nao traz mais ficaria para tras (e o editor serve o que estiver em motor\editors).
Type: filesandordirs; Name: "{app}\motor"
Type: filesandordirs; Name: "{app}\WebView2Runtime"
; Catalogo de fontes por maquina: sempre recriado do zero, pelo administrador (ver [Run]). Um catalogo
; anterior, gravado quando a pasta ainda aceitava escrita de qualquer usuario, nao pode sobreviver: o
; AllFonts.js alimenta a lista de arquivos que o ServidorAscDesktop aceita servir (backlog 20).
Type: filesandordirs; Name: "{commonappdata}\mOffice\fontes"

[Dirs]
Name: "{commonappdata}\mOffice"; Flags: uninsalwaysuninstall
Name: "{commonappdata}\mOffice\fontes"; Flags: uninsalwaysuninstall

[Files]
; O pacote inteiro da Task 1. Os Excludes sao defesa em profundidade: o montar-pacote.ps1 ja confere a
; ausencia de cada um (e aborta), mas o instalador nao depende disso.
;   *.pdb          simbolos carregam caminho de build;
;   sdkjs-plugins  plugins do ONLYOFFICE (fora do produto; alguns falam com a rede);
;   AllFonts.js    catalogo da maquina de BUILD (o do usuario e gerado na instalacao, ver [Run]);
;   ASC.ttf        fonte proprietaria (substituida pela mOffice Simbolos - backlog 13).
;   *.orig         copias do original deixadas pelo aplicar-patches.ps1 (o sdk-all.js.orig tem a fonte ASCW3 da Ascensio).
Source: "{#PacoteDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs; Excludes: "*.pdb,*.orig,sdkjs-plugins,AllFonts.js,ASC.ttf"

[Icons]
; Menu Iniciar: sempre os tres apps (sem pasta de grupo, como o Windows 11 mostra).
Name: "{autoprograms}\mTexto"; Filename: "{app}\mTexto.exe"
Name: "{autoprograms}\mPlanilha"; Filename: "{app}\mPlanilha.exe"
Name: "{autoprograms}\mSlides"; Filename: "{app}\mSlides.exe"
Name: "{autodesktop}\mTexto"; Filename: "{app}\mTexto.exe"; Tasks: atalhodesktop
Name: "{autodesktop}\mPlanilha"; Filename: "{app}\mPlanilha.exe"; Tasks: atalhodesktop
Name: "{autodesktop}\mSlides"; Filename: "{app}\mSlides.exe"; Tasks: atalhodesktop

[Registry]
; Task "associar" (marcada por padrao). Espelha tools/explorer/registrar-dev.ps1, mas em HKLM (a maquina
; inteira) e com os .exe instalados. REGISTRA os apps como opcao em "Abrir com" e em Configuracoes >
; Aplicativos padrao - NAO forca o padrao (essa escolha e do usuario, ou da politica do RMM em tools/rmm).
;
; Desinstalar remove SO o que e nosso:
;   - os ProgIDs (um por extensao que abre) e as chaves Applications\<exe> e Software\mOffice\<App>: exclusivas do
;     mOffice (uninsdeletekey);
;   - em .<ext>\OpenWithProgids e em RegisteredApplications, so o NOSSO valor (uninsdeletevalue) - as
;     chaves sao compartilhadas com outros programas e nunca sao apagadas.

; As linhas abaixo (ProgIDs, OpenWithProgids, Applications, Capabilities, RegisteredApplications) sao GERADAS da TabelaFormatos
; (Plano 6, Task 2): uma por extensao que o mOffice abre. Icone = o do proprio .exe; abrir = "<exe>" "%1"; modelo (.dotx...) =
; "<exe>" /modelo "%1" (abre como documento NOVO) e apresentacao de slides (.ppsx...) = "<exe>" /apresentar "%1", ambos com o
; verbo edit = "<exe>" "%1". Nao editar a mao: tests/mOffice.Comum.Tests/GeradorAssociacoes.cs (AssociacoesGeradas_EmDia).
#include "associacoes.iss"

; Os dados de cada usuario (%LOCALAPPDATA%\mOffice: configuracoes, recentes, logs, recuperacao) NAO sao
; tocados pelo instalador nem pelo desinstalador.

[Run]
; (1) Catalogo de fontes POR MAQUINA (backlog 20). A pasta {commonappdata}\mOffice ja chega aqui PROTEGIDA
; (so leitura para Usuarios; SYSTEM e Administradores com controle total; dono = Administradores): quem faz isso
; e o [Code], no CurStepChanged(ssInstall), ANTES do [InstallDelete]/[Dirs] - ver PrepararPastaDeDados. Ali a
; ACL e aplicada numa pasta VAZIA (nunca com /T): uma pasta que ja existia e nao e a desta instalacao (pode ter
; sido pre-criada por um usuario comum, que como DONO manteria WRITE_DAC, e ter dentro links ou hard links que
; uma varredura nao enxerga) e renomeada para o lado, sem ser percorrida. fontes\ nasce dentro dela ([Dirs])
; e herda a ACL. O Check: SemLinks abaixo e so DEFESA EM PROFUNDIDADE (varredura de links logo antes).
; ...e depois o catalogo e gerado pelo MESMO caminho do app (CatalogoFontes.Padrao + GarantirAtualizadoAsync,
; verbo /gerarcatalogo roteado na subida do AplicativoOffice), como administrador: o arquivo nasce com a ACL
; acima. Um usuario comum nao consegue escrever la; o app dele cai para %LOCALAPPDATA%\mOffice\fontes
; (ConfiguracaoMotor.SondarPastaFontes). Falha aqui nao aborta a instalacao: o app gera o catalogo na
; primeira abertura (o codigo de saida fica no log do instalador).
Filename: "{app}\mTexto.exe"; Parameters: "/gerarcatalogo"; Flags: runhidden waituntilterminated; StatusMsg: "Gerando o catálogo de fontes..."; Check: SemLinks
; (2) Firewall (backlog 22): no [Code] (CurStepChanged(ssPostInstall)), pelo netsh direto - ver la.
; (3) Reabre o app ao FIM da atualizacao pelo botao "Verificar atualizacoes": o app passa /REABRIR=<app>
; (InstaladorSilencioso.MontarStartInfo) e so ESSE e reaberto. SEM skipifsilent (a atualizacao e sempre
; silenciosa: /SILENT pelo botao, /VERYSILENT pelo RMM); runasoriginaluser = des-elevado, como o usuario;
; nowait = nao prende o instalador.
; DeveReabrir devolve False com /NORELAUNCH (update do RMM: a atualizacao em massa e invisivel) e sem
; /REABRIR (instalacao a mao: o usuario abre pelo Menu Iniciar).
Filename: "{app}\mTexto.exe"; Description: "Abrir o mTexto"; Flags: nowait postinstall runasoriginaluser; Check: DeveReabrir('mTexto')
Filename: "{app}\mPlanilha.exe"; Description: "Abrir o mPlanilha"; Flags: nowait postinstall runasoriginaluser; Check: DeveReabrir('mPlanilha')
Filename: "{app}\mSlides.exe"; Description: "Abrir o mSlides"; Flags: nowait postinstall runasoriginaluser; Check: DeveReabrir('mSlides')

; Sem [UninstallRun]/[UninstallDelete]: regras de firewall, politica do RMM e o catalogo por maquina saem pelo
; [Code] (CurUninstallStepChanged(usUninstall)) - o catalogo so depois da checagem de link, e nada disso
; depende do PowerShell.

[Code]
// AppMutex= (ver o comentario em [Setup]): o mutex so vale quando alguem pode responder a pergunta. No
// instalador silencioso a resposta suprimida e Cancelar e o instalador abortaria; quem fecha os apps ali e o
// CloseApplications (Restart Manager).
function MutexDoApp(Param: String): String;
begin
  Result := '{#MutexDeAtualizacao}';
  if WizardSilent then
    Result := '';
end;

// O desinstalador NAO reavalia o AppMutex=: usa o valor gravado na instalacao (MEDIDO, Task 3: instalado em
// silencio, o desinstalador nao checou o mutex segurado). Por isso a checagem do desinstalador e feita aqui,
// sempre, com a mesma mensagem e o mesmo comportamento do Inno (em silencio: Cancelar = nao desinstala).
function InitializeUninstall(): Boolean;
begin
  Result := True;
  while CheckForMutexes('{#MutexDeAtualizacao}') do
    if SuppressibleMsgBox(FmtMessage(SetupMessage(msgUninstallAppRunningError), ['{#MyAppName}']),
         mbError, MB_OKCANCEL, IDCANCEL) <> IDOK then
    begin
      Result := False;
      Exit;
    end;
end;

// Parametro de linha de comando exato (sem diferenciar maiusculas).
function TemParametro(const Nome: String): Boolean;
var
  i: Integer;
begin
  Result := False;
  for i := 1 to ParamCount do
    if CompareText(ParamStr(i), Nome) = 0 then
    begin
      Result := True;
      Exit;
    end;
end;

// Valor de um parametro /NOME=valor, ou '' se ausente.
function ValorDoParametro(const Prefixo: String): String;
var
  i: Integer;
begin
  Result := '';
  for i := 1 to ParamCount do
    if CompareText(Copy(ParamStr(i), 1, Length(Prefixo)), Prefixo) = 0 then
    begin
      Result := Copy(ParamStr(i), Length(Prefixo) + 1, MaxInt);
      Exit;
    end;
end;

// Reabre SO o app que pediu a atualizacao (/REABRIR=<app>), e nunca com /NORELAUNCH (update do RMM).
function DeveReabrir(App: String): Boolean;
begin
  Result := False;
  if TemParametro('/NORELAUNCH') then
    Exit;
  Result := CompareText(ValorDoParametro('/REABRIR='), App) = 0;
end;

// ---------------------------------------------------------------------------------------------------------
// LINKS em {commonappdata}\mOffice - DEFESA EM PROFUNDIDADE. A protecao principal e o PrepararPastaDeDados
// (pasta criada vazia pela instalacao; a de outra origem e renomeada sem ser percorrida). Esta varredura roda
// de novo antes do /gerarcatalogo e na desinstalacao. FindFirst sem curinga devolve a PROPRIA entrada, sem
// seguir o link, com FILE_ATTRIBUTE_REPARSE_POINT quando ela e juncao/symlink (hard link ela NAO ve).
// ---------------------------------------------------------------------------------------------------------
function PastaDoMOffice(): String;
begin
  Result := ExpandConstant('{commonappdata}\mOffice');
end;

function EhLink(const Caminho: String): Boolean;
var
  R: TFindRec;
begin
  Result := False;
  if FindFirst(Caminho, R) then
  try
    Result := (R.Attributes and FILE_ATTRIBUTE_REPARSE_POINT) <> 0;
  finally
    FindClose(R);
  end;
end;

// Apaga o LINK (RemoveDir/DeleteFile num reparse point removem o link, nunca o destino). True se saiu.
function ApagarLink(const Caminho: String): Boolean;
begin
  Log('mOffice: link (reparse point) encontrado em ' + Caminho + ' - apagando o link, nao o destino.');
  if not RemoveDir(Caminho) then
    DeleteFile(Caminho);
  Result := not EhLink(Caminho);
  if not Result then
    Log('mOffice: NAO foi possivel apagar o link ' + Caminho + '.');
end;

// Percorre a arvore SEM entrar em link; apaga cada link achado. Devolve quantos links NAO sairam.
function LinksRestantes(const Pasta: String): Integer;
var
  R: TFindRec;
  Filho: String;
begin
  Result := 0;
  if FindFirst(Pasta + '\*', R) then
  try
    repeat
      if (R.Name <> '.') and (R.Name <> '..') then
      begin
        Filho := Pasta + '\' + R.Name;
        if (R.Attributes and FILE_ATTRIBUTE_REPARSE_POINT) <> 0 then
        begin
          if not ApagarLink(Filho) then
            Result := Result + 1;
        end
        else if (R.Attributes and FILE_ATTRIBUTE_DIRECTORY) <> 0 then
          Result := Result + LinksRestantes(Filho);
      end;
    until not FindNext(R);
  finally
    FindClose(R);
  end;
end;

// True quando {commonappdata}\mOffice e tudo dentro dele estao sem link (apagando os que houver).
function RemoverLinks(): Boolean;
var
  Raiz: String;
begin
  Raiz := PastaDoMOffice();
  if EhLink(Raiz) then
    if not ApagarLink(Raiz) then
    begin
      Result := False;
      Exit;
    end;
  Result := LinksRestantes(Raiz) = 0;
end;

// Check: do /gerarcatalogo no [Run] - defesa em profundidade: repete a varredura imediatamente antes.
function SemLinks(): Boolean;
begin
  Result := RemoverLinks();
  if not Result then
    Log('mOffice: link em ' + PastaDoMOffice() + ' que nao pode ser apagado - geracao do catalogo PULADA.');
end;

// ---------------------------------------------------------------------------------------------------------
// FIREWALL (backlog 22) pelo netsh direto: sem PowerShell (politica de execucao/AllSigned nao impede) e com
// o codigo de saida de CADA regra conferido. Mesmo nome das regras do tools/egresso/firewall.ps1 -
// "<prefixo> (<exe>)" -, entao o -Remover dele e o FirewallTests continuam enxergando as regras.
// ---------------------------------------------------------------------------------------------------------
const
  PrefixoRegra = 'mOffice WebView2 - bloqueio de saida';

function Netsh(const Parametros: String): Integer;
begin
  if not Exec(ExpandConstant('{sys}\netsh.exe'), Parametros, '', SW_HIDE, ewWaitUntilTerminated, Result) then
    Result := -1;
end;

procedure Acrescentar(var Lista: TArrayOfString; const Item: String);
var
  n: Integer;
begin
  n := GetArrayLength(Lista);
  SetArrayLength(Lista, n + 1);
  Lista[n] := Item;
end;

// Todos os .exe de Pasta, recursivo (sem entrar em link) - a mesma lista do Get-ChildItem -Recurse do script.
procedure ColetarExes(const Pasta: String; var Lista: TArrayOfString);
var
  R: TFindRec;
begin
  if FindFirst(Pasta + '\*', R) then
  try
    repeat
      if (R.Name <> '.') and (R.Name <> '..') then
      begin
        if (R.Attributes and FILE_ATTRIBUTE_DIRECTORY) <> 0 then
        begin
          if (R.Attributes and FILE_ATTRIBUTE_REPARSE_POINT) = 0 then
            ColetarExes(Pasta + '\' + R.Name, Lista);
        end
        else if CompareText(ExtractFileExt(R.Name), '.exe') = 0 then
          Acrescentar(Lista, Pasta + '\' + R.Name);
      end;
    until not FindNext(R);
  finally
    FindClose(R);
  end;
end;

function NomeDaRegra(const Exe: String): String;
begin
  Result := PrefixoRegra + ' (' + ExtractFileName(Exe) + ')';
end;

// Nomes das regras EXISTENTES com o prefixo (le "netsh ... show rule name=all", rotulo em ingles ou em
// portugues, como o firewall.ps1). Em outro idioma a lista vem vazia e a remocao usa os nomes dos .exe.
function RegrasExistentes(): TArrayOfString;
var
  Arquivo, L, Rotulo: String;
  Linhas: TArrayOfString;
  i, p, Codigo: Integer;
begin
  SetArrayLength(Result, 0);
  Arquivo := GenerateUniqueName(GetTempDir(), '.txt');
  if Exec(ExpandConstant('{cmd}'), '/C ""' + ExpandConstant('{sys}\netsh.exe') + '" advfirewall firewall show rule name=all > "' + Arquivo + '""',
      '', SW_HIDE, ewWaitUntilTerminated, Codigo) then
    if LoadStringsFromFile(Arquivo, Linhas) then
      for i := 0 to GetArrayLength(Linhas) - 1 do
      begin
        L := Trim(Linhas[i]);
        p := Pos(':', L);
        if p > 0 then
        begin
          Rotulo := Trim(Copy(L, 1, p - 1));
          if (CompareText(Rotulo, 'Rule Name') = 0) or (CompareText(Rotulo, 'Nome da Regra') = 0) then
          begin
            L := Trim(Copy(L, p + 1, MaxInt));
            if Pos(PrefixoRegra, L) = 1 then
              Acrescentar(Result, L);
          end;
        end;
      end;
  DeleteFile(Arquivo);
end;

// Apaga as regras do mOffice: as achadas pelo prefixo + as dos .exe atuais (idioma sem rotulo conhecido).
// Devolve quantas continuam depois (so contaveis pelo prefixo).
function RemoverRegrasFirewall(): Integer;
var
  Exes, Nomes: TArrayOfString;
  i: Integer;
begin
  Nomes := RegrasExistentes();
  SetArrayLength(Exes, 0);
  ColetarExes(ExpandConstant('{app}\WebView2Runtime'), Exes);
  for i := 0 to GetArrayLength(Exes) - 1 do
    Acrescentar(Nomes, NomeDaRegra(Exes[i]));
  for i := 0 to GetArrayLength(Nomes) - 1 do
    Netsh('advfirewall firewall delete rule name="' + Nomes[i] + '"');
  Result := GetArrayLength(RegrasExistentes());
  if Result > 0 then
    Log('mOffice: ' + IntToStr(Result) + ' regra(s) de firewall do mOffice continuam depois da remocao.')
  else
    Log('mOffice: regras de firewall do mOffice removidas.');
end;

const
  AvisoFirewall = 'Não foi possível criar no Firewall do Windows a regra que bloqueia o acesso à internet do ' +
    'navegador interno do mOffice. O mOffice funciona normalmente, mas sem essa proteção extra. ' +
    'Os detalhes estão no log da instalação (%TEMP%, "Setup Log").';

// Uma regra de bloqueio de SAIDA por .exe de {app}\WebView2Runtime. True se TODAS foram criadas.
function CriarRegrasFirewall(): Boolean;
var
  Exes: TArrayOfString;
  i, Codigo, Falhas: Integer;
begin
  SetArrayLength(Exes, 0);
  ColetarExes(ExpandConstant('{app}\WebView2Runtime'), Exes);
  Falhas := 0;
  for i := 0 to GetArrayLength(Exes) - 1 do
  begin
    Netsh('advfirewall firewall delete rule name="' + NomeDaRegra(Exes[i]) + '"');
    Codigo := Netsh('advfirewall firewall add rule name="' + NomeDaRegra(Exes[i]) + '" dir=out action=block program="' +
      Exes[i] + '" enable=yes profile=any');
    if Codigo = 0 then
      Log('mOffice: regra de firewall criada: ' + NomeDaRegra(Exes[i]))
    else
    begin
      Log('mOffice: FALHA ao criar a regra de firewall ' + NomeDaRegra(Exes[i]) + ' (netsh codigo ' + IntToStr(Codigo) + ').');
      Falhas := Falhas + 1;
    end;
  end;
  if GetArrayLength(Exes) = 0 then
    Log('mOffice: FALHA - nenhum .exe em ' + ExpandConstant('{app}\WebView2Runtime') + ' para bloquear.');
  Result := (Falhas = 0) and (GetArrayLength(Exes) > 0);
end;

// ---------------------------------------------------------------------------------------------------------
// PASTA DE DADOS POR MAQUINA ({commonappdata}\mOffice; revisao da Task 3, rodada 2). Ameaca: %ProgramData%
// deixa qualquer usuario criar pastas. Um usuario comum pode pre-criar mOffice (ou trocar por juncao, ou por
// dentro criar links e hard links para arquivos protegidos) e, como DONO, manter WRITE_DAC para reabrir a ACL
// no meio de um icacls /T. Por isso a instalacao NUNCA percorre com direitos de admin uma pasta que nao criou:
//   - identidade da pasta (volume + indice NTFS do arquivo, que inclui o numero de sequencia) gravada em HKLM
//     na criacao; na instalacao seguinte, a MESMA identidade = pasta desta instalacao (so admin cria algo
//     dentro dela) -> mantida (o XML da politica do RMM continua), com a ACL reaplicada so nela (sem /T);
//   - qualquer outra coisa (pasta pre-existente, identidade diferente, marcador ausente) e RENOMEADA para
//     mOffice.antigo-<data>-<aleatorio> (rename nao percorre nada; nome aleatorio com uma nova tentativa) e deixada para limpeza manual (log). Nada de dentro dela
//     e restaurado: o aplicar-padrao-office.ps1 do RMM so roda com os ProgIDs registrados, ou seja, depois de
//     uma instalacao - numa pasta de outra origem nao ha XML legitimo a preservar; o RMM o regenera;
//   - a pasta nova e criada em {tmp} (so do administrador), recebe a ACL final VAZIA e so entao e movida para
//     {commonappdata} (mesmo volume: move atomico, a ACL protegida vai junto). Se o move falhar (volumes
//     diferentes), FALHA FECHADO: a instalacao e cancelada; a pasta nunca e criada direto em %ProgramData%.
// ---------------------------------------------------------------------------------------------------------
type
  TInfoArquivo = record
    Atributos: DWORD;
    CriacaoLo, CriacaoHi, AcessoLo, AcessoHi, EscritaLo, EscritaHi: DWORD;
    Volume: DWORD;
    TamanhoHi, TamanhoLo: DWORD;
    Links: DWORD;
    IndiceHi, IndiceLo: DWORD;
  end;

function CreateFileW(lpFileName: String; dwDesiredAccess, dwShareMode: DWORD; lpSecurityAttributes: THandle;
  dwCreationDisposition, dwFlagsAndAttributes: DWORD; hTemplateFile: THandle): THandle;
  external 'CreateFileW@kernel32.dll stdcall';
function GetFileInformationByHandle(hFile: THandle; var Info: TInfoArquivo): BOOL;
  external 'GetFileInformationByHandle@kernel32.dll stdcall';
function CloseHandle(hObject: THandle): BOOL;
  external 'CloseHandle@kernel32.dll stdcall';

const
  ChaveInstalacao = 'Software\mOffice\Instalacao';
  ValorIdentidade = 'PastaDeDados';

function ChaveRaizDaMarca(): Integer;
begin
  Result := HKLM;
end;

// "<volume>-<indice alto>-<indice baixo>" da propria entrada (sem seguir link), ou '' se nao abrir ou se for link.
function IdentidadeDaPasta(const Caminho: String): String;
var
  H: THandle;
  Info: TInfoArquivo;
begin
  Result := '';
  // FILE_READ_ATTRIBUTES; compartilhamento total; OPEN_EXISTING; BACKUP_SEMANTICS (pasta) + OPEN_REPARSE_POINT.
  H := CreateFileW(Caminho, $80, 7, 0, 3, $02000000 or $00200000, 0);
  if (H = 0) or (H = THandle(-1)) then
    Exit;
  try
    if GetFileInformationByHandle(H, Info) then
      if ((Info.Atributos and FILE_ATTRIBUTE_REPARSE_POINT) = 0) and ((Info.Atributos and FILE_ATTRIBUTE_DIRECTORY) <> 0) then
        Result := IntToStr(Info.Volume) + '-' + IntToStr(Info.IndiceHi) + '-' + IntToStr(Info.IndiceLo);
  finally
    CloseHandle(H);
  end;
end;

function PastaEhDaInstalacao(): Boolean;
var
  Marca, Id: String;
begin
  Id := IdentidadeDaPasta(PastaDoMOffice());
  Result := (Id <> '') and RegQueryStringValue(ChaveRaizDaMarca(), ChaveInstalacao, ValorIdentidade, Marca) and (Marca = Id);
end;

// ACL final numa pasta SEM /T: dono Administradores; heranca cortada; SYSTEM e Administradores F, Usuarios RX,
// herdaveis. /L: se o caminho for link, age no link. SIDs conhecidos: os nomes mudam com o idioma do Windows.
function AplicarAcl(const Pasta: String): Boolean;
var
  Dono, Acl: Integer;
begin
  Dono := -1;
  Acl := -1;
  if not Exec(ExpandConstant('{sys}\icacls.exe'), '"' + Pasta + '" /setowner *S-1-5-32-544 /L', '', SW_HIDE, ewWaitUntilTerminated, Dono) then
    Dono := -1;
  if not Exec(ExpandConstant('{sys}\icacls.exe'), '"' + Pasta + '" /inheritance:r /grant:r *S-1-5-18:(OI)(CI)F *S-1-5-32-544:(OI)(CI)F *S-1-5-32-545:(OI)(CI)RX /L', '', SW_HIDE, ewWaitUntilTerminated, Acl) then
    Acl := -1;
  Result := (Dono = 0) and (Acl = 0);
  Log('mOffice: ACL em ' + Pasta + ': dono=' + IntToStr(Dono) + ' acl=' + IntToStr(Acl));
end;

// "<raiz>.antigo-<data>-<aleatorio>".
function NomeDaAntiga(const Raiz: String): String;
begin
  Result := Raiz + '.antigo-' + GetDateTimeString('yyyymmddhhnnss', #0, #0) + '-' + IntToStr(Random(1000000000));
end;

function PrepararPastaDeDados(): Boolean;
var
  Raiz, Antiga, Nova, Id: String;
begin
  Result := False;
  Raiz := PastaDoMOffice();
  if EhLink(Raiz) then
    if not ApagarLink(Raiz) then
      Exit;
  if DirExists(Raiz) and PastaEhDaInstalacao() then
  begin
    Log('mOffice: ' + Raiz + ' e a pasta criada por uma instalacao anterior (mesma identidade) - mantida.');
    Result := AplicarAcl(Raiz);
    Exit;
  end;
  if DirExists(Raiz) or FileExists(Raiz) then
  begin
    // Sufixo ALEATORIO (e uma segunda tentativa com outro): um usuario comum que pre-criasse o nome de destino
    // nao consegue travar a instalacao.
    Antiga := NomeDaAntiga(Raiz);
    if not RenameFile(Raiz, Antiga) then
    begin
      Log('mOffice: renomear ' + Raiz + ' para ' + Antiga + ' falhou - tentando outro nome.');
      Antiga := NomeDaAntiga(Raiz);
      if not RenameFile(Raiz, Antiga) then
      begin
        Log('mOffice: ' + Raiz + ' existe e nao e desta instalacao, e NAO foi possivel renomea-la.');
        Exit;
      end;
    end;
    Log('mOffice: ' + Raiz + ' existia e nao e desta instalacao - renomeada para ' + Antiga +
      ' SEM ser percorrida. Nada foi aproveitado dela; apague-a a mao se nao for necessaria.');
  end;
  Nova := ExpandConstant('{tmp}\mOffice-dados');
  if not CreateDir(Nova) then
    Exit;
  if not AplicarAcl(Nova) then
    Exit;
  // Falha FECHADO: se o move nao der (ex.: {tmp} em outro volume), NAO cria no lugar - uma pasta criada em
  // %ProgramData% fica, ate a ACL, gravavel por qualquer usuario. A instalacao para com a mensagem do ssInstall.
  if not RenameFile(Nova, Raiz) then
  begin
    Log('mOffice: FALHA ao mover a pasta protegida de ' + Nova + ' para ' + Raiz +
      ' (volumes diferentes?) - instalacao cancelada; a pasta NAO e criada no lugar.');
    Exit;
  end;
  Id := IdentidadeDaPasta(Raiz);
  if Id = '' then
    Exit;
  if not RegWriteStringValue(ChaveRaizDaMarca(), ChaveInstalacao, ValorIdentidade, Id) then
    Exit;
  Log('mOffice: pasta de dados ' + Raiz + ' criada e protegida (identidade ' + Id + ').');
  Result := True;
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssInstall then
  begin
    // Antes do [InstallDelete]/[Dirs]/[Files]: {commonappdata}\mOffice da instalacao, protegida e sem link.
    if not PrepararPastaDeDados() then
    begin
      SuppressibleMsgBox('Não foi possível preparar com segurança a pasta ' + PastaDoMOffice() + '. ' +
        'Por segurança, a instalação foi cancelada. Os detalhes estão no log da instalação (%TEMP%, "Setup Log").',
        mbCriticalError, MB_OK, IDOK);
      Abort;
    end;
  end
  else if CurStep = ssPostInstall then
  begin
    // Sempre limpa as regras anteriores (runtime novo pode ter outros .exe; tarefa desmarcada desfaz a regra).
    RemoverRegrasFirewall();
    if WizardIsTaskSelected('firewall') then
      if not CriarRegrasFirewall() then
        SuppressibleMsgBox(AvisoFirewall, mbError, MB_OK, IDOK);
  end;
end;

// ---------------------------------------------------------------------------------------------------------
// DESINSTALACAO: regras de firewall, politica de associacao do RMM (tools/rmm) e o catalogo por maquina.
// ---------------------------------------------------------------------------------------------------------

// Uma linha <Association .../> de uma das SEIS extensoes do mOffice (o aplicar-padrao-office.ps1 grava um
// elemento por linha).
function AssociacaoDoMOffice(const L: String): Boolean;
begin
  // Plano 6, Task 2: "do mOffice" = aponta para um ProgID mOffice.* (as 17 do padrao do RMM e as seis de antes), em vez de uma
  // lista de extensoes repetida aqui — a lista mora na TabelaFormatos e uma politica antiga continua sendo desfeita por inteiro.
  Result := (Pos('<Association ', L) > 0) and (Pos('ProgId="mOffice.', L) > 0);
end;

// Mesma regra do remover-padrao-office.ps1: so mexe se a politica aponta para o XML DO mOffice; tira as
// associacoes do mOffice (as de ProgID mOffice.*); sobrando associacao de outro app (ex.: .pdf do mPDF), o XML fica so com ela e a
// politica continua; senao apaga o VALOR (nunca a chave) e o XML. Os ProgIDs saem com o desinstalador, e
// uma politica apontando para ProgID inexistente faria o Windows desfazer o padrao do usuario.
procedure RemoverPoliticaDoRmm();
var
  Chave, Valor, Xml: String;
  Linhas, Resto: TArrayOfString;
  i, Outras: Integer;
begin
  Chave := 'SOFTWARE\Policies\Microsoft\Windows\System';
  Xml := PastaDoMOffice() + '\moffice-padrao-office.xml';
  if not RegQueryStringValue(HKLM, Chave, 'DefaultAssociationsConfiguration', Valor) then
    Exit;
  if CompareText(Valor, Xml) <> 0 then
  begin
    Log('mOffice: a politica de associacao aponta para ' + Valor + ' (nao e do mOffice) - nao alterada.');
    Exit;
  end;
  Outras := 0;
  SetArrayLength(Resto, 0);
  // XML presente mas que e link ou nao pode ser lido: nao se sabe o que ele tem (pode ter a associacao de
  // outro app) - registra e NAO mexe na politica. Ausente: a politica aponta para o vazio e sai.
  if FileExists(Xml) or EhLink(Xml) then
  begin
    if EhLink(Xml) or (not LoadStringsFromFile(Xml, Linhas)) then
    begin
      Log('mOffice: ' + Xml + ' e link ou nao pode ser lido - politica de associacao NAO alterada.');
      Exit;
    end;
    for i := 0 to GetArrayLength(Linhas) - 1 do
      if not AssociacaoDoMOffice(Linhas[i]) then
      begin
        Acrescentar(Resto, Linhas[i]);
        if Pos('<Association ', Linhas[i]) > 0 then
          Outras := Outras + 1;
      end;
  end;
  if Outras > 0 then
  begin
    SaveStringsToUTF8FileWithoutBOM(Xml, Resto, False);
    Log('mOffice: associacoes do mOffice tiradas da politica; ' + IntToStr(Outras) + ' de outros apps mantidas.');
  end
  else
  begin
    RegDeleteValue(HKLM, Chave, 'DefaultAssociationsConfiguration');
    DeleteFile(Xml);
    Log('mOffice: politica de associacao do mOffice removida.');
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usUninstall then
  begin
    // Antes de os arquivos sairem ({app}\WebView2Runtime ainda existe para os nomes das regras).
    if RemoverRegrasFirewall() > 0 then
      SuppressibleMsgBox('Não foi possível remover do Firewall do Windows todas as regras do mOffice ' +
        '("' + PrefixoRegra + ' ..."). Elas podem ser apagadas em "Firewall do Windows com Segurança Avançada".',
        mbError, MB_OK, IDOK);
    // Catalogo por maquina (cache regeneravel) e politica do RMM: so na pasta DESTA instalacao (identidade
    // igual a gravada) e sem link dentro. Outra pasta no lugar: nao e percorrida nem alterada.
    if PastaEhDaInstalacao() and RemoverLinks() then
    begin
      RemoverPoliticaDoRmm();
      DelTree(PastaDoMOffice() + '\fontes', True, True, True);
    end
    else
      Log('mOffice: ' + PastaDoMOffice() + ' nao e a pasta desta instalacao (ou tem link) - catalogo e politica NAO removidos.');
  end
  else if CurUninstallStep = usPostUninstall then
  begin
    // A marca de identidade so sai junto com a pasta: se ela ficou (XML do RMM com a associacao de outro app),
    // a proxima instalacao ainda a reconhece como sua e nao a renomeia.
    if not DirExists(PastaDoMOffice()) then
    begin
      RegDeleteKeyIncludingSubkeys(ChaveRaizDaMarca(), ChaveInstalacao);
      RegDeleteKeyIfEmpty(ChaveRaizDaMarca(), 'Software\mOffice');
    end;
  end;
end;
