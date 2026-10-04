; GERADO por tests/mOffice.Comum.Tests/GeradorAssociacoes.cs a partir de TabelaFormatos - nao editar a mao.
; Regenerar: $env:MOFFICE_REGENERAR=1; dotnet test tests/mOffice.Comum.Tests --filter AssociacoesGeradas_EmDia
; Incluido pelo [Registry] do moffice.iss (#include "associacoes.iss"): comentarios e regras ficam la.
; --- ProgIDs: um por extensao que o mOffice abre (icone do tipo; abrir = "<exe>" "%1"; modelo = /modelo, apresentacao = /apresentar; imprimir = /imprimir) ---
Root: HKLM; Subkey: "Software\Classes\mOffice.Texto.docx"; ValueType: string; ValueName: ""; ValueData: "Documento do Word (mTexto)"; Flags: uninsdeletekey; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Texto.docx\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: """{app}\Icones\Documentos\mtexto-documento.ico"",0"; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Texto.docx\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mTexto.exe"" ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Texto.docx\shell\print\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mTexto.exe"" /imprimir ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Texto.docm"; ValueType: string; ValueName: ""; ValueData: "Documento Habilitado para Macro do Word (mTexto)"; Flags: uninsdeletekey; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Texto.docm\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: """{app}\Icones\Documentos\mtexto-macro.ico"",0"; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Texto.docm\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mTexto.exe"" ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Texto.docm\shell\print\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mTexto.exe"" /imprimir ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Texto.dotx"; ValueType: string; ValueName: ""; ValueData: "Modelo do Word (mTexto)"; Flags: uninsdeletekey; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Texto.dotx\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: """{app}\Icones\Documentos\mtexto-modelo.ico"",0"; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Texto.dotx\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mTexto.exe"" /modelo ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Texto.dotx\shell\print\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mTexto.exe"" /imprimir ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Texto.dotx\shell\edit"; ValueType: string; ValueName: ""; ValueData: "Editar"; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Texto.dotx\shell\edit\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mTexto.exe"" ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Texto.dotm"; ValueType: string; ValueName: ""; ValueData: "Modelo Habilitado para Macro do Word (mTexto)"; Flags: uninsdeletekey; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Texto.dotm\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: """{app}\Icones\Documentos\mtexto-modelo.ico"",0"; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Texto.dotm\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mTexto.exe"" /modelo ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Texto.dotm\shell\print\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mTexto.exe"" /imprimir ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Texto.dotm\shell\edit"; ValueType: string; ValueName: ""; ValueData: "Editar"; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Texto.dotm\shell\edit\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mTexto.exe"" ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Texto.doc"; ValueType: string; ValueName: ""; ValueData: "Documento do Word 97-2003 (mTexto)"; Flags: uninsdeletekey; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Texto.doc\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: """{app}\Icones\Documentos\mtexto-legado.ico"",0"; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Texto.doc\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mTexto.exe"" ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Texto.doc\shell\print\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mTexto.exe"" /imprimir ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Texto.dot"; ValueType: string; ValueName: ""; ValueData: "Modelo do Word 97-2003 (mTexto)"; Flags: uninsdeletekey; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Texto.dot\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: """{app}\Icones\Documentos\mtexto-modelo.ico"",0"; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Texto.dot\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mTexto.exe"" /modelo ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Texto.dot\shell\print\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mTexto.exe"" /imprimir ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Texto.dot\shell\edit"; ValueType: string; ValueName: ""; ValueData: "Editar"; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Texto.dot\shell\edit\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mTexto.exe"" ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Texto.odt"; ValueType: string; ValueName: ""; ValueData: "Documento de texto ODF (mTexto)"; Flags: uninsdeletekey; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Texto.odt\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: """{app}\Icones\Documentos\mtexto-documento.ico"",0"; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Texto.odt\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mTexto.exe"" ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Texto.odt\shell\print\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mTexto.exe"" /imprimir ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Texto.ott"; ValueType: string; ValueName: ""; ValueData: "Modelo de texto ODF (mTexto)"; Flags: uninsdeletekey; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Texto.ott\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: """{app}\Icones\Documentos\mtexto-modelo.ico"",0"; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Texto.ott\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mTexto.exe"" /modelo ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Texto.ott\shell\print\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mTexto.exe"" /imprimir ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Texto.ott\shell\edit"; ValueType: string; ValueName: ""; ValueData: "Editar"; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Texto.ott\shell\edit\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mTexto.exe"" ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Texto.rtf"; ValueType: string; ValueName: ""; ValueData: "Formato Rich Text (mTexto)"; Flags: uninsdeletekey; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Texto.rtf\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: """{app}\Icones\Documentos\mtexto-documento.ico"",0"; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Texto.rtf\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mTexto.exe"" ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Texto.rtf\shell\print\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mTexto.exe"" /imprimir ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Planilha.xlsx"; ValueType: string; ValueName: ""; ValueData: "Pasta de trabalho do Excel (mPlanilha)"; Flags: uninsdeletekey; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Planilha.xlsx\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: """{app}\Icones\Documentos\mplanilha-documento.ico"",0"; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Planilha.xlsx\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mPlanilha.exe"" ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Planilha.xlsx\shell\print\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mPlanilha.exe"" /imprimir ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Planilha.xlsm"; ValueType: string; ValueName: ""; ValueData: "Pasta de trabalho habilitada para macro do Excel (mPlanilha)"; Flags: uninsdeletekey; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Planilha.xlsm\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: """{app}\Icones\Documentos\mplanilha-macro.ico"",0"; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Planilha.xlsm\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mPlanilha.exe"" ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Planilha.xlsm\shell\print\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mPlanilha.exe"" /imprimir ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Planilha.xltx"; ValueType: string; ValueName: ""; ValueData: "Modelo do Excel (mPlanilha)"; Flags: uninsdeletekey; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Planilha.xltx\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: """{app}\Icones\Documentos\mplanilha-modelo.ico"",0"; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Planilha.xltx\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mPlanilha.exe"" /modelo ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Planilha.xltx\shell\print\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mPlanilha.exe"" /imprimir ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Planilha.xltx\shell\edit"; ValueType: string; ValueName: ""; ValueData: "Editar"; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Planilha.xltx\shell\edit\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mPlanilha.exe"" ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Planilha.xltm"; ValueType: string; ValueName: ""; ValueData: "Modelo habilitado para macro do Excel (mPlanilha)"; Flags: uninsdeletekey; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Planilha.xltm\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: """{app}\Icones\Documentos\mplanilha-modelo.ico"",0"; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Planilha.xltm\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mPlanilha.exe"" /modelo ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Planilha.xltm\shell\print\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mPlanilha.exe"" /imprimir ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Planilha.xltm\shell\edit"; ValueType: string; ValueName: ""; ValueData: "Editar"; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Planilha.xltm\shell\edit\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mPlanilha.exe"" ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Planilha.xls"; ValueType: string; ValueName: ""; ValueData: "Pasta de trabalho do Excel 97-2003 (mPlanilha)"; Flags: uninsdeletekey; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Planilha.xls\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: """{app}\Icones\Documentos\mplanilha-legado.ico"",0"; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Planilha.xls\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mPlanilha.exe"" ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Planilha.xls\shell\print\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mPlanilha.exe"" /imprimir ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Planilha.xlt"; ValueType: string; ValueName: ""; ValueData: "Modelo do Excel 97-2003 (mPlanilha)"; Flags: uninsdeletekey; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Planilha.xlt\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: """{app}\Icones\Documentos\mplanilha-modelo.ico"",0"; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Planilha.xlt\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mPlanilha.exe"" /modelo ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Planilha.xlt\shell\print\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mPlanilha.exe"" /imprimir ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Planilha.xlt\shell\edit"; ValueType: string; ValueName: ""; ValueData: "Editar"; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Planilha.xlt\shell\edit\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mPlanilha.exe"" ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Planilha.xlsb"; ValueType: string; ValueName: ""; ValueData: "Pasta de trabalho binária do Excel (mPlanilha)"; Flags: uninsdeletekey; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Planilha.xlsb\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: """{app}\Icones\Documentos\mplanilha-legado.ico"",0"; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Planilha.xlsb\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mPlanilha.exe"" ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Planilha.xlsb\shell\print\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mPlanilha.exe"" /imprimir ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Planilha.ods"; ValueType: string; ValueName: ""; ValueData: "Planilha ODF (mPlanilha)"; Flags: uninsdeletekey; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Planilha.ods\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: """{app}\Icones\Documentos\mplanilha-documento.ico"",0"; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Planilha.ods\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mPlanilha.exe"" ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Planilha.ods\shell\print\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mPlanilha.exe"" /imprimir ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Planilha.ots"; ValueType: string; ValueName: ""; ValueData: "Modelo de planilha ODF (mPlanilha)"; Flags: uninsdeletekey; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Planilha.ots\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: """{app}\Icones\Documentos\mplanilha-modelo.ico"",0"; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Planilha.ots\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mPlanilha.exe"" /modelo ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Planilha.ots\shell\print\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mPlanilha.exe"" /imprimir ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Planilha.ots\shell\edit"; ValueType: string; ValueName: ""; ValueData: "Editar"; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Planilha.ots\shell\edit\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mPlanilha.exe"" ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Planilha.csv"; ValueType: string; ValueName: ""; ValueData: "CSV (separado por ponto e vírgula) (mPlanilha)"; Flags: uninsdeletekey; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Planilha.csv\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: """{app}\Icones\Documentos\mplanilha-documento.ico"",0"; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Planilha.csv\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mPlanilha.exe"" ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Planilha.csv\shell\print\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mPlanilha.exe"" /imprimir ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Slides.pptx"; ValueType: string; ValueName: ""; ValueData: "Apresentação do PowerPoint (mSlides)"; Flags: uninsdeletekey; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Slides.pptx\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: """{app}\Icones\Documentos\mslides-documento.ico"",0"; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Slides.pptx\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mSlides.exe"" ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Slides.pptx\shell\print\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mSlides.exe"" /imprimir ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Slides.pptm"; ValueType: string; ValueName: ""; ValueData: "Apresentação habilitada para macro do PowerPoint (mSlides)"; Flags: uninsdeletekey; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Slides.pptm\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: """{app}\Icones\Documentos\mslides-macro.ico"",0"; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Slides.pptm\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mSlides.exe"" ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Slides.pptm\shell\print\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mSlides.exe"" /imprimir ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Slides.potx"; ValueType: string; ValueName: ""; ValueData: "Modelo do PowerPoint (mSlides)"; Flags: uninsdeletekey; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Slides.potx\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: """{app}\Icones\Documentos\mslides-modelo.ico"",0"; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Slides.potx\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mSlides.exe"" /modelo ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Slides.potx\shell\print\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mSlides.exe"" /imprimir ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Slides.potx\shell\edit"; ValueType: string; ValueName: ""; ValueData: "Editar"; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Slides.potx\shell\edit\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mSlides.exe"" ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Slides.potm"; ValueType: string; ValueName: ""; ValueData: "Modelo habilitado para macro do PowerPoint (mSlides)"; Flags: uninsdeletekey; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Slides.potm\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: """{app}\Icones\Documentos\mslides-modelo.ico"",0"; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Slides.potm\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mSlides.exe"" /modelo ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Slides.potm\shell\print\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mSlides.exe"" /imprimir ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Slides.potm\shell\edit"; ValueType: string; ValueName: ""; ValueData: "Editar"; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Slides.potm\shell\edit\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mSlides.exe"" ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Slides.ppsx"; ValueType: string; ValueName: ""; ValueData: "Apresentação de slides do PowerPoint (mSlides)"; Flags: uninsdeletekey; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Slides.ppsx\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: """{app}\Icones\Documentos\mslides-apresentacao.ico"",0"; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Slides.ppsx\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mSlides.exe"" /apresentar ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Slides.ppsx\shell\print\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mSlides.exe"" /imprimir ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Slides.ppsx\shell\edit"; ValueType: string; ValueName: ""; ValueData: "Editar"; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Slides.ppsx\shell\edit\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mSlides.exe"" ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Slides.ppsm"; ValueType: string; ValueName: ""; ValueData: "Apresentação de slides habilitada para macro do PowerPoint (mSlides)"; Flags: uninsdeletekey; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Slides.ppsm\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: """{app}\Icones\Documentos\mslides-apresentacao.ico"",0"; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Slides.ppsm\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mSlides.exe"" /apresentar ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Slides.ppsm\shell\print\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mSlides.exe"" /imprimir ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Slides.ppsm\shell\edit"; ValueType: string; ValueName: ""; ValueData: "Editar"; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Slides.ppsm\shell\edit\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mSlides.exe"" ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Slides.ppt"; ValueType: string; ValueName: ""; ValueData: "Apresentação do PowerPoint 97-2003 (mSlides)"; Flags: uninsdeletekey; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Slides.ppt\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: """{app}\Icones\Documentos\mslides-legado.ico"",0"; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Slides.ppt\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mSlides.exe"" ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Slides.ppt\shell\print\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mSlides.exe"" /imprimir ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Slides.pps"; ValueType: string; ValueName: ""; ValueData: "Apresentação de slides do PowerPoint 97-2003 (mSlides)"; Flags: uninsdeletekey; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Slides.pps\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: """{app}\Icones\Documentos\mslides-apresentacao.ico"",0"; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Slides.pps\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mSlides.exe"" /apresentar ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Slides.pps\shell\print\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mSlides.exe"" /imprimir ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Slides.pps\shell\edit"; ValueType: string; ValueName: ""; ValueData: "Editar"; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Slides.pps\shell\edit\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mSlides.exe"" ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Slides.pot"; ValueType: string; ValueName: ""; ValueData: "Modelo do PowerPoint 97-2003 (mSlides)"; Flags: uninsdeletekey; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Slides.pot\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: """{app}\Icones\Documentos\mslides-modelo.ico"",0"; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Slides.pot\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mSlides.exe"" /modelo ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Slides.pot\shell\print\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mSlides.exe"" /imprimir ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Slides.pot\shell\edit"; ValueType: string; ValueName: ""; ValueData: "Editar"; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Slides.pot\shell\edit\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mSlides.exe"" ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Slides.odp"; ValueType: string; ValueName: ""; ValueData: "Apresentação ODF (mSlides)"; Flags: uninsdeletekey; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Slides.odp\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: """{app}\Icones\Documentos\mslides-documento.ico"",0"; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Slides.odp\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mSlides.exe"" ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Slides.odp\shell\print\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mSlides.exe"" /imprimir ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Slides.otp"; ValueType: string; ValueName: ""; ValueData: "Modelo de apresentação ODF (mSlides)"; Flags: uninsdeletekey; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Slides.otp\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: """{app}\Icones\Documentos\mslides-modelo.ico"",0"; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Slides.otp\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mSlides.exe"" /modelo ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Slides.otp\shell\print\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mSlides.exe"" /imprimir ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Slides.otp\shell\edit"; ValueType: string; ValueName: ""; ValueData: "Editar"; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\mOffice.Slides.otp\shell\edit\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mSlides.exe"" ""%1"""; Tasks: associar
; --- OpenWithProgids: o app aparece no "Abrir com" da extensao (valor VAZIO; o que conta e o nome) ---
Root: HKLM; Subkey: "Software\Classes\.docx\OpenWithProgids"; ValueType: string; ValueName: "mOffice.Texto.docx"; ValueData: ""; Flags: uninsdeletevalue; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\.docm\OpenWithProgids"; ValueType: string; ValueName: "mOffice.Texto.docm"; ValueData: ""; Flags: uninsdeletevalue; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\.dotx\OpenWithProgids"; ValueType: string; ValueName: "mOffice.Texto.dotx"; ValueData: ""; Flags: uninsdeletevalue; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\.dotm\OpenWithProgids"; ValueType: string; ValueName: "mOffice.Texto.dotm"; ValueData: ""; Flags: uninsdeletevalue; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\.doc\OpenWithProgids"; ValueType: string; ValueName: "mOffice.Texto.doc"; ValueData: ""; Flags: uninsdeletevalue; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\.dot\OpenWithProgids"; ValueType: string; ValueName: "mOffice.Texto.dot"; ValueData: ""; Flags: uninsdeletevalue; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\.odt\OpenWithProgids"; ValueType: string; ValueName: "mOffice.Texto.odt"; ValueData: ""; Flags: uninsdeletevalue; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\.ott\OpenWithProgids"; ValueType: string; ValueName: "mOffice.Texto.ott"; ValueData: ""; Flags: uninsdeletevalue; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\.rtf\OpenWithProgids"; ValueType: string; ValueName: "mOffice.Texto.rtf"; ValueData: ""; Flags: uninsdeletevalue; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\.xlsx\OpenWithProgids"; ValueType: string; ValueName: "mOffice.Planilha.xlsx"; ValueData: ""; Flags: uninsdeletevalue; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\.xlsm\OpenWithProgids"; ValueType: string; ValueName: "mOffice.Planilha.xlsm"; ValueData: ""; Flags: uninsdeletevalue; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\.xltx\OpenWithProgids"; ValueType: string; ValueName: "mOffice.Planilha.xltx"; ValueData: ""; Flags: uninsdeletevalue; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\.xltm\OpenWithProgids"; ValueType: string; ValueName: "mOffice.Planilha.xltm"; ValueData: ""; Flags: uninsdeletevalue; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\.xls\OpenWithProgids"; ValueType: string; ValueName: "mOffice.Planilha.xls"; ValueData: ""; Flags: uninsdeletevalue; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\.xlt\OpenWithProgids"; ValueType: string; ValueName: "mOffice.Planilha.xlt"; ValueData: ""; Flags: uninsdeletevalue; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\.xlsb\OpenWithProgids"; ValueType: string; ValueName: "mOffice.Planilha.xlsb"; ValueData: ""; Flags: uninsdeletevalue; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\.ods\OpenWithProgids"; ValueType: string; ValueName: "mOffice.Planilha.ods"; ValueData: ""; Flags: uninsdeletevalue; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\.ots\OpenWithProgids"; ValueType: string; ValueName: "mOffice.Planilha.ots"; ValueData: ""; Flags: uninsdeletevalue; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\.csv\OpenWithProgids"; ValueType: string; ValueName: "mOffice.Planilha.csv"; ValueData: ""; Flags: uninsdeletevalue; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\.pptx\OpenWithProgids"; ValueType: string; ValueName: "mOffice.Slides.pptx"; ValueData: ""; Flags: uninsdeletevalue; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\.pptm\OpenWithProgids"; ValueType: string; ValueName: "mOffice.Slides.pptm"; ValueData: ""; Flags: uninsdeletevalue; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\.potx\OpenWithProgids"; ValueType: string; ValueName: "mOffice.Slides.potx"; ValueData: ""; Flags: uninsdeletevalue; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\.potm\OpenWithProgids"; ValueType: string; ValueName: "mOffice.Slides.potm"; ValueData: ""; Flags: uninsdeletevalue; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\.ppsx\OpenWithProgids"; ValueType: string; ValueName: "mOffice.Slides.ppsx"; ValueData: ""; Flags: uninsdeletevalue; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\.ppsm\OpenWithProgids"; ValueType: string; ValueName: "mOffice.Slides.ppsm"; ValueData: ""; Flags: uninsdeletevalue; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\.ppt\OpenWithProgids"; ValueType: string; ValueName: "mOffice.Slides.ppt"; ValueData: ""; Flags: uninsdeletevalue; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\.pps\OpenWithProgids"; ValueType: string; ValueName: "mOffice.Slides.pps"; ValueData: ""; Flags: uninsdeletevalue; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\.pot\OpenWithProgids"; ValueType: string; ValueName: "mOffice.Slides.pot"; ValueData: ""; Flags: uninsdeletevalue; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\.odp\OpenWithProgids"; ValueType: string; ValueName: "mOffice.Slides.odp"; ValueData: ""; Flags: uninsdeletevalue; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\.otp\OpenWithProgids"; ValueType: string; ValueName: "mOffice.Slides.otp"; ValueData: ""; Flags: uninsdeletevalue; Tasks: associar
; --- ShellNew: clique direito > Novo (.docx/.xlsx/.pptx). Chave da EXTENSAO, compartilhada com o Office: so' o nosso valor sai (uninsdeletevalue) e a chave so' se vazia; NUNCA uninsdeletekey. ---
Root: HKLM; Subkey: "Software\Classes\.docx\ShellNew"; ValueType: string; ValueName: "FileName"; ValueData: "{app}\Modelos\pt-BR\new.docx"; Flags: uninsdeletevalue; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\.docx\ShellNew"; Flags: uninsdeletekeyifempty; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\.xlsx\ShellNew"; ValueType: string; ValueName: "FileName"; ValueData: "{app}\Modelos\pt-BR\new.xlsx"; Flags: uninsdeletevalue; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\.xlsx\ShellNew"; Flags: uninsdeletekeyifempty; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\.pptx\ShellNew"; ValueType: string; ValueName: "FileName"; ValueData: "{app}\Modelos\pt-BR\new.pptx"; Flags: uninsdeletevalue; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\.pptx\ShellNew"; Flags: uninsdeletekeyifempty; Tasks: associar
; --- Aplicativo REGISTRADO: e disso que o Windows 11 monta "Escolher outro aplicativo" e Aplicativos padrao ---
; mTexto
Root: HKLM; Subkey: "Software\Classes\Applications\mTexto.exe"; ValueType: string; ValueName: "FriendlyAppName"; ValueData: "mTexto"; Flags: uninsdeletekey; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\Applications\mTexto.exe\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: """{app}\mTexto.exe"",0"; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\Applications\mTexto.exe\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mTexto.exe"" ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\Applications\mTexto.exe\SupportedTypes"; ValueType: string; ValueName: ".docx"; ValueData: ""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\Applications\mTexto.exe\SupportedTypes"; ValueType: string; ValueName: ".docm"; ValueData: ""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\Applications\mTexto.exe\SupportedTypes"; ValueType: string; ValueName: ".dotx"; ValueData: ""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\Applications\mTexto.exe\SupportedTypes"; ValueType: string; ValueName: ".dotm"; ValueData: ""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\Applications\mTexto.exe\SupportedTypes"; ValueType: string; ValueName: ".doc"; ValueData: ""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\Applications\mTexto.exe\SupportedTypes"; ValueType: string; ValueName: ".dot"; ValueData: ""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\Applications\mTexto.exe\SupportedTypes"; ValueType: string; ValueName: ".odt"; ValueData: ""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\Applications\mTexto.exe\SupportedTypes"; ValueType: string; ValueName: ".ott"; ValueData: ""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\Applications\mTexto.exe\SupportedTypes"; ValueType: string; ValueName: ".rtf"; ValueData: ""; Tasks: associar
Root: HKLM; Subkey: "Software\mOffice"; Flags: uninsdeletekeyifempty; Tasks: associar
Root: HKLM; Subkey: "Software\mOffice\mTexto"; Flags: uninsdeletekey; Tasks: associar
Root: HKLM; Subkey: "Software\mOffice\mTexto\Capabilities"; ValueType: string; ValueName: "ApplicationName"; ValueData: "mTexto"; Tasks: associar
Root: HKLM; Subkey: "Software\mOffice\mTexto\Capabilities"; ValueType: string; ValueName: "ApplicationDescription"; ValueData: "Editor de textos do mOffice"; Tasks: associar
Root: HKLM; Subkey: "Software\mOffice\mTexto\Capabilities"; ValueType: string; ValueName: "ApplicationIcon"; ValueData: """{app}\mTexto.exe"",0"; Tasks: associar
Root: HKLM; Subkey: "Software\mOffice\mTexto\Capabilities\FileAssociations"; ValueType: string; ValueName: ".docx"; ValueData: "mOffice.Texto.docx"; Tasks: associar
Root: HKLM; Subkey: "Software\mOffice\mTexto\Capabilities\FileAssociations"; ValueType: string; ValueName: ".docm"; ValueData: "mOffice.Texto.docm"; Tasks: associar
Root: HKLM; Subkey: "Software\mOffice\mTexto\Capabilities\FileAssociations"; ValueType: string; ValueName: ".dotx"; ValueData: "mOffice.Texto.dotx"; Tasks: associar
Root: HKLM; Subkey: "Software\mOffice\mTexto\Capabilities\FileAssociations"; ValueType: string; ValueName: ".dotm"; ValueData: "mOffice.Texto.dotm"; Tasks: associar
Root: HKLM; Subkey: "Software\mOffice\mTexto\Capabilities\FileAssociations"; ValueType: string; ValueName: ".doc"; ValueData: "mOffice.Texto.doc"; Tasks: associar
Root: HKLM; Subkey: "Software\mOffice\mTexto\Capabilities\FileAssociations"; ValueType: string; ValueName: ".dot"; ValueData: "mOffice.Texto.dot"; Tasks: associar
Root: HKLM; Subkey: "Software\mOffice\mTexto\Capabilities\FileAssociations"; ValueType: string; ValueName: ".odt"; ValueData: "mOffice.Texto.odt"; Tasks: associar
Root: HKLM; Subkey: "Software\mOffice\mTexto\Capabilities\FileAssociations"; ValueType: string; ValueName: ".ott"; ValueData: "mOffice.Texto.ott"; Tasks: associar
Root: HKLM; Subkey: "Software\mOffice\mTexto\Capabilities\FileAssociations"; ValueType: string; ValueName: ".rtf"; ValueData: "mOffice.Texto.rtf"; Tasks: associar
Root: HKLM; Subkey: "Software\RegisteredApplications"; ValueType: string; ValueName: "mOffice.mTexto"; ValueData: "Software\mOffice\mTexto\Capabilities"; Flags: uninsdeletevalue; Tasks: associar
; mPlanilha
Root: HKLM; Subkey: "Software\Classes\Applications\mPlanilha.exe"; ValueType: string; ValueName: "FriendlyAppName"; ValueData: "mPlanilha"; Flags: uninsdeletekey; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\Applications\mPlanilha.exe\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: """{app}\mPlanilha.exe"",0"; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\Applications\mPlanilha.exe\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mPlanilha.exe"" ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\Applications\mPlanilha.exe\SupportedTypes"; ValueType: string; ValueName: ".xlsx"; ValueData: ""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\Applications\mPlanilha.exe\SupportedTypes"; ValueType: string; ValueName: ".xlsm"; ValueData: ""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\Applications\mPlanilha.exe\SupportedTypes"; ValueType: string; ValueName: ".xltx"; ValueData: ""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\Applications\mPlanilha.exe\SupportedTypes"; ValueType: string; ValueName: ".xltm"; ValueData: ""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\Applications\mPlanilha.exe\SupportedTypes"; ValueType: string; ValueName: ".xls"; ValueData: ""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\Applications\mPlanilha.exe\SupportedTypes"; ValueType: string; ValueName: ".xlt"; ValueData: ""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\Applications\mPlanilha.exe\SupportedTypes"; ValueType: string; ValueName: ".xlsb"; ValueData: ""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\Applications\mPlanilha.exe\SupportedTypes"; ValueType: string; ValueName: ".ods"; ValueData: ""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\Applications\mPlanilha.exe\SupportedTypes"; ValueType: string; ValueName: ".ots"; ValueData: ""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\Applications\mPlanilha.exe\SupportedTypes"; ValueType: string; ValueName: ".csv"; ValueData: ""; Tasks: associar
Root: HKLM; Subkey: "Software\mOffice\mPlanilha"; Flags: uninsdeletekey; Tasks: associar
Root: HKLM; Subkey: "Software\mOffice\mPlanilha\Capabilities"; ValueType: string; ValueName: "ApplicationName"; ValueData: "mPlanilha"; Tasks: associar
Root: HKLM; Subkey: "Software\mOffice\mPlanilha\Capabilities"; ValueType: string; ValueName: "ApplicationDescription"; ValueData: "Planilhas do mOffice"; Tasks: associar
Root: HKLM; Subkey: "Software\mOffice\mPlanilha\Capabilities"; ValueType: string; ValueName: "ApplicationIcon"; ValueData: """{app}\mPlanilha.exe"",0"; Tasks: associar
Root: HKLM; Subkey: "Software\mOffice\mPlanilha\Capabilities\FileAssociations"; ValueType: string; ValueName: ".xlsx"; ValueData: "mOffice.Planilha.xlsx"; Tasks: associar
Root: HKLM; Subkey: "Software\mOffice\mPlanilha\Capabilities\FileAssociations"; ValueType: string; ValueName: ".xlsm"; ValueData: "mOffice.Planilha.xlsm"; Tasks: associar
Root: HKLM; Subkey: "Software\mOffice\mPlanilha\Capabilities\FileAssociations"; ValueType: string; ValueName: ".xltx"; ValueData: "mOffice.Planilha.xltx"; Tasks: associar
Root: HKLM; Subkey: "Software\mOffice\mPlanilha\Capabilities\FileAssociations"; ValueType: string; ValueName: ".xltm"; ValueData: "mOffice.Planilha.xltm"; Tasks: associar
Root: HKLM; Subkey: "Software\mOffice\mPlanilha\Capabilities\FileAssociations"; ValueType: string; ValueName: ".xls"; ValueData: "mOffice.Planilha.xls"; Tasks: associar
Root: HKLM; Subkey: "Software\mOffice\mPlanilha\Capabilities\FileAssociations"; ValueType: string; ValueName: ".xlt"; ValueData: "mOffice.Planilha.xlt"; Tasks: associar
Root: HKLM; Subkey: "Software\mOffice\mPlanilha\Capabilities\FileAssociations"; ValueType: string; ValueName: ".xlsb"; ValueData: "mOffice.Planilha.xlsb"; Tasks: associar
Root: HKLM; Subkey: "Software\mOffice\mPlanilha\Capabilities\FileAssociations"; ValueType: string; ValueName: ".ods"; ValueData: "mOffice.Planilha.ods"; Tasks: associar
Root: HKLM; Subkey: "Software\mOffice\mPlanilha\Capabilities\FileAssociations"; ValueType: string; ValueName: ".ots"; ValueData: "mOffice.Planilha.ots"; Tasks: associar
Root: HKLM; Subkey: "Software\mOffice\mPlanilha\Capabilities\FileAssociations"; ValueType: string; ValueName: ".csv"; ValueData: "mOffice.Planilha.csv"; Tasks: associar
Root: HKLM; Subkey: "Software\RegisteredApplications"; ValueType: string; ValueName: "mOffice.mPlanilha"; ValueData: "Software\mOffice\mPlanilha\Capabilities"; Flags: uninsdeletevalue; Tasks: associar
; mSlides
Root: HKLM; Subkey: "Software\Classes\Applications\mSlides.exe"; ValueType: string; ValueName: "FriendlyAppName"; ValueData: "mSlides"; Flags: uninsdeletekey; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\Applications\mSlides.exe\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: """{app}\mSlides.exe"",0"; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\Applications\mSlides.exe\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\mSlides.exe"" ""%1"""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\Applications\mSlides.exe\SupportedTypes"; ValueType: string; ValueName: ".pptx"; ValueData: ""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\Applications\mSlides.exe\SupportedTypes"; ValueType: string; ValueName: ".pptm"; ValueData: ""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\Applications\mSlides.exe\SupportedTypes"; ValueType: string; ValueName: ".potx"; ValueData: ""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\Applications\mSlides.exe\SupportedTypes"; ValueType: string; ValueName: ".potm"; ValueData: ""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\Applications\mSlides.exe\SupportedTypes"; ValueType: string; ValueName: ".ppsx"; ValueData: ""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\Applications\mSlides.exe\SupportedTypes"; ValueType: string; ValueName: ".ppsm"; ValueData: ""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\Applications\mSlides.exe\SupportedTypes"; ValueType: string; ValueName: ".ppt"; ValueData: ""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\Applications\mSlides.exe\SupportedTypes"; ValueType: string; ValueName: ".pps"; ValueData: ""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\Applications\mSlides.exe\SupportedTypes"; ValueType: string; ValueName: ".pot"; ValueData: ""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\Applications\mSlides.exe\SupportedTypes"; ValueType: string; ValueName: ".odp"; ValueData: ""; Tasks: associar
Root: HKLM; Subkey: "Software\Classes\Applications\mSlides.exe\SupportedTypes"; ValueType: string; ValueName: ".otp"; ValueData: ""; Tasks: associar
Root: HKLM; Subkey: "Software\mOffice\mSlides"; Flags: uninsdeletekey; Tasks: associar
Root: HKLM; Subkey: "Software\mOffice\mSlides\Capabilities"; ValueType: string; ValueName: "ApplicationName"; ValueData: "mSlides"; Tasks: associar
Root: HKLM; Subkey: "Software\mOffice\mSlides\Capabilities"; ValueType: string; ValueName: "ApplicationDescription"; ValueData: "Apresentações do mOffice"; Tasks: associar
Root: HKLM; Subkey: "Software\mOffice\mSlides\Capabilities"; ValueType: string; ValueName: "ApplicationIcon"; ValueData: """{app}\mSlides.exe"",0"; Tasks: associar
Root: HKLM; Subkey: "Software\mOffice\mSlides\Capabilities\FileAssociations"; ValueType: string; ValueName: ".pptx"; ValueData: "mOffice.Slides.pptx"; Tasks: associar
Root: HKLM; Subkey: "Software\mOffice\mSlides\Capabilities\FileAssociations"; ValueType: string; ValueName: ".pptm"; ValueData: "mOffice.Slides.pptm"; Tasks: associar
Root: HKLM; Subkey: "Software\mOffice\mSlides\Capabilities\FileAssociations"; ValueType: string; ValueName: ".potx"; ValueData: "mOffice.Slides.potx"; Tasks: associar
Root: HKLM; Subkey: "Software\mOffice\mSlides\Capabilities\FileAssociations"; ValueType: string; ValueName: ".potm"; ValueData: "mOffice.Slides.potm"; Tasks: associar
Root: HKLM; Subkey: "Software\mOffice\mSlides\Capabilities\FileAssociations"; ValueType: string; ValueName: ".ppsx"; ValueData: "mOffice.Slides.ppsx"; Tasks: associar
Root: HKLM; Subkey: "Software\mOffice\mSlides\Capabilities\FileAssociations"; ValueType: string; ValueName: ".ppsm"; ValueData: "mOffice.Slides.ppsm"; Tasks: associar
Root: HKLM; Subkey: "Software\mOffice\mSlides\Capabilities\FileAssociations"; ValueType: string; ValueName: ".ppt"; ValueData: "mOffice.Slides.ppt"; Tasks: associar
Root: HKLM; Subkey: "Software\mOffice\mSlides\Capabilities\FileAssociations"; ValueType: string; ValueName: ".pps"; ValueData: "mOffice.Slides.pps"; Tasks: associar
Root: HKLM; Subkey: "Software\mOffice\mSlides\Capabilities\FileAssociations"; ValueType: string; ValueName: ".pot"; ValueData: "mOffice.Slides.pot"; Tasks: associar
Root: HKLM; Subkey: "Software\mOffice\mSlides\Capabilities\FileAssociations"; ValueType: string; ValueName: ".odp"; ValueData: "mOffice.Slides.odp"; Tasks: associar
Root: HKLM; Subkey: "Software\mOffice\mSlides\Capabilities\FileAssociations"; ValueType: string; ValueName: ".otp"; ValueData: "mOffice.Slides.otp"; Tasks: associar
Root: HKLM; Subkey: "Software\RegisteredApplications"; ValueType: string; ValueName: "mOffice.mSlides"; ValueData: "Software\mOffice\mSlides\Capabilities"; Flags: uninsdeletevalue; Tasks: associar
