# mOffice

Suíte de escritório para Windows com três aplicativos:

- **mTexto** — documentos (DOCX, ODT);
- **mPlanilha** — planilhas (XLSX, ODS);
- **mSlides** — apresentações (PPTX, ODP).

Os documentos são abertos, editados e salvos no próprio computador. O mOffice não tem conta, nuvem,
telemetria nem anúncios.

O mOffice é baseado no **ONLYOFFICE Desktop Editors 9.4.0.129**: o motor de edição do ONLYOFFICE (o
conversor `x2t` e os editores `sdkjs` e `web-apps`) roda dentro de uma casca própria, escrita do zero em
.NET/WPF, que o hospeda num WebView2. É uma **versão modificada** e não é o ONLYOFFICE Desktop Editors:
não é publicada nem endossada pela Ascensio System SIA. As modificações feitas no motor estão em
[`docs/licencas/MODIFICACOES-ONLYOFFICE.md`](docs/licencas/MODIFICACOES-ONLYOFFICE.md).

## Instalar

O instalador de cada versão está em [Releases](../../releases), com o sha256 do arquivo. Ele instala os
três aplicativos, o motor e o runtime WebView2 (Fixed Version) que eles usam; não precisa de Office nem
de ONLYOFFICE instalados.

## Sem rede

No uso normal o mOffice **não se conecta a nada** (egresso zero):

- o editor roda a partir de arquivos locais, e o runtime WebView2 é uma versão fixa embarcada, sem o
  serviço de atualização do Edge;
- o instalador cria regras de firewall que bloqueiam os executáveis do runtime;
- a única conexão que o produto sabe fazer é a verificação de atualização, e ela só acontece quando você
  pede (botão "Verificar atualizações" nas Configurações, ou os parâmetros `/checkupdate` e `/update`
  para quem administra várias máquinas). Ela consulta as releases deste repositório no GitHub. Não há
  verificação automática.

Todo o código de rede fica num único arquivo (`src/mOffice.Comum/Atualizacao/ServicoAtualizacao.cs`), e
testes estruturais falham se aparecer rede em qualquer outro lugar.

## Compilar

Pré-requisitos: Windows x64 e [.NET 10 SDK](https://dotnet.microsoft.com/).

```powershell
dotnet build mOffice.slnx
dotnet test mOffice.slnx
```

O motor do ONLYOFFICE e o runtime WebView2 **não** ficam no git. Sem eles o código compila e os testes
que precisam do motor são pulados. Para rodar os aplicativos, obtenha os dois com os scripts de `tools/`
(cada um confere versão e sha256):

```powershell
# 1. motor: instale o ONLYOFFICE Desktop Editors 9.4.0.129 (MSI oficial) e copie as pastas necessárias
.\tools\onlyoffice\obter-motor.ps1
# 2. gerador do catálogo de fontes, do pacote do ONLYOFFICE Docs 9.4.0.129 para Windows
.\tools\onlyoffice\obter-allfontsgen.ps1 -Pacote <caminho>\onlyoffice-documentserver.exe
# 3. runtime WebView2 Fixed Version
.\tools\webview2\obter-runtime.ps1

dotnet build src/mTexto
.\src\mTexto\bin\Debug\net10.0-windows\mTexto.exe
```

O instalador é gerado por `tools/installer/build-installer.ps1` (Inno Setup 6), que roda a suíte de
testes, monta o pacote com `tools/installer/montar-pacote.ps1` e confere que nenhum caminho da máquina de
build ficou nos binários. Os links, versões e sha256 exatos de tudo o que vem do ONLYOFFICE estão em
[`docs/licencas/FONTE-CORRESPONDENTE.md`](docs/licencas/FONTE-CORRESPONDENTE.md).

## Licença

O mOffice é software livre, sob a **GNU Affero General Public License versão 3** ([`LICENSE`](LICENSE)),
a mesma licença do ONLYOFFICE. Os componentes de terceiros que vão no instalador, com versão, licença e
texto integral de cada um, estão em [`docs/licencas/INVENTARIO.md`](docs/licencas/INVENTARIO.md). A
fonte correspondente do motor (tags e commits exatos, pacotes oficiais e modificações) está em
[`docs/licencas/FONTE-CORRESPONDENTE.md`](docs/licencas/FONTE-CORRESPONDENTE.md).

## Marcas

ONLYOFFICE is a trademark of Ascensio System SIA. O nome aparece aqui só para identificar a origem do
motor, como a licença exige. A AGPL-3.0 não concede direito algum sobre as marcas, logotipos ou
identidade visual da Ascensio System SIA, e o mOffice não usa nenhum deles: nome, ícones e identidade
visual são próprios.

Microsoft, Windows, Microsoft Edge e WebView2 são marcas do grupo de empresas Microsoft. O mOffice não é
produzido nem endossado pela Microsoft.
