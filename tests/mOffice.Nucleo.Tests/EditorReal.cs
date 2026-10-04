using mOffice.Conversor;

namespace mOffice.Nucleo.Tests;

/// <summary>Helpers compartilhados pelos testes de janela real (<see cref="AbrirDocumentoTests"/>,
/// <see cref="SalvarDocumentoTests"/>): esperar o editor ficar pronto, ler/inserir texto pela API publica dele,
/// e o Skip.If comum de configuracao de maquina. Extraido na revisao 2026-09-17 (fix review 1, minor 10) —
/// estavam duplicados nas duas classes e a Task 7 ia duplicar de novo.</summary>
public static class EditorReal
{
    /// <summary>Skip (nao Fail) quando a maquina nao da para abrir a janela real: sessao RDP (o teste precisa
    /// de sessao de console local) ou <c>allfontsgen.exe</c> ausente da arvore.
    ///
    /// Task 2: ate aqui o skip era "cache de fontes do ONLYOFFICE OFICIAL ausente" — e isso significava que,
    /// no PC sem o ONLYOFFICE instalado (exatamente o PC que o risco 2 descreve), a suite ficava verde sem
    /// testar nada. Agora o catalogo e' GERADO pela propria JanelaEditor na inicializacao, entao a unica
    /// condicao de maquina que resta e' o gerador existir: <c>third_party/</c> e' ignorado pelo git e numa
    /// arvore recem-clonada e' preciso rodar <c>tools/onlyoffice/obter-motor.ps1</c> e
    /// <c>tools/onlyoffice/obter-allfontsgen.ps1</c>. A pasta do catalogo NAO e' checada de proposito: se ela
    /// estiver vazia, o teste TEM de rodar e provar que a janela a preenche sozinha.</summary>
    public static void PularSeNaoDaParaAbrirJanela()
    {
        Skip.If(JanelaSta.SessaoRemota, "teste de janela real: só em sessão local (não RDP)");
        var gerador = Path.Combine(mOffice.Conversor.Tests.MotorDeTeste.Localizar().PastaMotor, "converter", "allfontsgen.exe");
        Skip.If(!File.Exists(gerador), "allfontsgen.exe ausente — rode tools/onlyoffice/obter-allfontsgen.ps1");
    }

    /// <summary>
    /// Janela de teste com a pergunta de fechamento já injetada — <b>use esta, não <c>new JanelaEditor(...)</c></b>,
    /// em todo teste de janela real que chame <c>Close()</c>.
    ///
    /// Por quê (achado da revisão 2 do branch do Plano 2, e a causa REAL do único vermelho da suíte em
    /// sessão de console): com documento modificado, o `Close()` cai no `PerguntarAoFechar` PADRÃO, que é um
    /// `MessageBox` de verdade. Num teste ninguém clica nele: a caixa fica aberta, a thread STA nunca
    /// termina, e o teste morre no timeout de `JanelaSta.Executar` — com a mensagem enganosa "hang do
    /// WPF/WebView2", que aponta para o lugar errado. Era o que acontecia em
    /// `MidiaTests.InserirImagemDeArquivo_PassaPeloDialogoEInsere` (insere imagem ⇒ modificado ⇒ `Close()`
    /// sem costura). Enquanto os testes rodavam só em RDP isso ficava escondido atrás do `[SKIP]`.
    ///
    /// O conserto é ESTRUTURAL de propósito: em vez de lembrar caso a caso, a fábrica é o caminho normal de
    /// criar janela nos testes. <see cref="JanelaEditor.PerguntaDeReserva"/> também vai injetada — senão uma
    /// exceção na pergunta cairia no `MessageBox` de reserva e reproduziria o mesmo travamento.
    /// </summary>
    public static JanelaEditor NovaJanela(string app = JanelaEditor.AppPadrao)
    {
        mOffice.Conversor.Tests.MotorDeTeste.Localizar(); // sem motor (clone publico sem third_party): pula, nao falha
        return PrepararParaTeste(new JanelaEditor(app));
    }

    /// <summary>Mesma injeção de <see cref="NovaJanela"/> para quem precisa do inicializador de objeto
    /// (<c>InstrumentarRede</c>, <c>Tema</c>, <c>NomeApp</c> — os dois últimos são `init` e só podem ser
    /// definidos na construção). Devolve a MESMA janela, para encadear.</summary>
    public static JanelaEditor PrepararParaTeste(JanelaEditor j)
    {
        j.PerguntarAoFechar = _ => Task.FromResult(DecisaoFechar.NaoSalvar);
        j.PerguntaDeReserva = _ => Task.FromResult(DecisaoFechar.NaoSalvar);
        // Plano 6, Task 5: as perguntas da trava (arquivo em uso, conflito ao salvar) tambem sao MessageBox por padrao.
        j.PerguntarEmUso = (_, _, _) => Task.FromResult(false);
        j.PerguntarConflito = _ => Task.FromResult(mOffice.Nucleo.Travas.DecisaoConflito.Cancelar);
        return j;
    }

    public static async Task<string> Esperar(JanelaEditor j, string expressao, string esperado, int tentativas)
    {
        var valor = "";
        for (var i = 0; i < tentativas && valor != esperado; i++)
        {
            await Task.Delay(200);
            valor = (await j.ExecutarJsAsync(expressao)).Trim('"');
        }
        return valor;
    }

    public const string ProntoNoEditor = "String(!!(window.__moffice && window.__moffice.pronto))";

    /// <summary>Espelho JS de <see cref="SessaoDocumento.Modificado"/> (Task 2 do Marco 1 / Plano 2):
    /// <c>ponte-desktop.js</c> grava <c>__moffice.sessaoModificada</c> no quadro E no documento de cima, no
    /// mesmo <c>onDocumentModifiedChanged</c> que avisa o host. E' o jeito DETERMINISTICO de esperar a edicao
    /// chegar ao modelo (em vez de um <c>Task.Delay</c> as cegas) sem perguntar nada ao C#.</summary>
    public const string EspelhoModificado = "String(!!(window.__moffice && window.__moffice.sessaoModificada))";

    public const string LerTextoDoEditor = """
        (function () {
          try {
            var w = window.frames[0];
            var api = (w.Asc && w.Asc.editor) ? w.Asc.editor : w.editor;
            if (!api) return "sem api";
            api.asc_EditSelectAll();
            return String(api.asc_GetSelectedText());
          } catch (e) { return "erro: " + e; }
        })()
        """;

    public const string TemTextoNovo = """
        (function () {
          try {
            var w = window.frames[0];
            var api = (w.Asc && w.Asc.editor) ? w.Asc.editor : w.editor;
            if (!api) return "sem api";
            api.asc_EditSelectAll();
            return String(String(api.asc_GetSelectedText()).indexOf("mOffice PoC salvou") >= 0);
          } catch (e) { return "erro: " + e; }
        })()
        """;

    // Insercao de texto pela API publica do editor. MEDIDO nesta janela (sonda descartavel, 2026-09-17):
    // este build NAO expoe asc_PasteData no objeto do editor — expoe asc_AddText(texto)
    // (sdkjs/word/sdk-all-min.js: `h.prototype.asc_AddText=h.prototype.FPi`, e FPi chama
    // documentoLogico.uQc(texto)). asc_PasteData e pluginMethod_PasteText ficam como alternativas, e o retorno
    // diz qual caminho foi usado (o teste registra isso na saida). Nenhum deles usa a area de transferencia do
    // sistema, entao nao ha permissao de clipboard envolvida.
    public const string InserirTexto = """
        (function () {
          try {
            var w = window.frames[0];
            var api = (w.Asc && w.Asc.editor) ? w.Asc.editor : w.editor;
            if (!api) return "sem api";
            // A leitura de linha de base usa asc_EditSelectAll e deixa o documento INTEIRO selecionado; sem
            // desfazer isso, a insercao substitui o texto antigo em vez de somar (medido: o DOCX salvo ficava
            // so com o texto novo).
            if (typeof api.asc_RemoveSelection === "function") api.asc_RemoveSelection();
            if (typeof api.asc_AddText === "function") { api.asc_AddText("mOffice PoC salvou "); return "ok asc_AddText"; }
            if (typeof api.asc_PasteData === "function") { api.asc_PasteData(1, "mOffice PoC salvou "); return "ok asc_PasteData"; }
            if (typeof api.pluginMethod_PasteText === "function") { api.pluginMethod_PasteText("mOffice PoC salvou "); return "ok pluginMethod_PasteText"; }
            return "sem API de insercao";
          } catch (e) { return "erro: " + e; }
        })()
        """;

    // -----------------------------------------------------------------------------------------------------
    // Task 7 — os TRES editores. As APIs abaixo foram levantadas no sdkjs 9.4.0.129 DESTE motor
    // (third_party/onlyoffice/editors/sdkjs/<cell|slide|word>/sdk-all-min.js) e confirmadas na janela real;
    // as hipoteses do brief (`asc_enterText`/`asc_insertInCell` para cell, `AddText` para slide) NAO foram
    // usadas como estao — ver o relatorio da Task 7, secao "APIs de edicao por editor".
    //
    //   word : asc_RemoveSelection() + asc_AddText(texto)                          (ja provado na Task 6)
    //   cell : asc_findCell("<ref>")  -> pluginMethod_InputText(texto) -> asc_closeCellEditor()
    //          asc_findCell e' a "caixa de nome" do editor (sdk-all-min.js: `Y.asc_findCell=Y.zLg`, que
    //          resolve a referencia e seleciona o intervalo); pluginMethod_InputText entra pelo mesmo
    //          caminho do teclado (`AscCommon.Cv.XJc`), e asc_closeCellEditor() (sem argumento, como o
    //          web-apps chama) confirma a edicao da celula.
    //   slide: pluginMethod_PasteText(texto) — cola texto na selecao corrente. NAO se faz asc_EditSelectAll
    //          antes (em slide ele seleciona TODOS os objetos do slide: `uye -> ta.Ga.qh()`, e a colagem
    //          substituiria o slide inteiro).
    //
    // Os tres sao `pluginMethod_*`/`asc_*` publicos do proprio motor — nada de interno do sdkjs foi tocado.
    // -----------------------------------------------------------------------------------------------------

    /// <summary>Preambulo comum: pega o objeto do editor dentro do quadro. `CORPO` recebe `api` (e `w`, o quadro).
    /// Publico desde a Task 2 do Plano 3 (a recuperacao chama `api.asc_Save(true)` por aqui).</summary>
    public static string NoEditor(string corpo) => """
        (function () {
          try {
            var w = window.frames[0];
            var api = (w.Asc && w.Asc.editor) ? w.Asc.editor : w.editor;
            if (!api) return "sem api";
            CORPO
          } catch (e) { return "erro: " + e; }
        })()
        """.Replace("CORPO", corpo);

    private static string Js(string texto) => System.Text.Json.JsonSerializer.Serialize(texto);

    /// <summary>Insercao no editor de TEXTO (word): igual a <see cref="InserirTexto"/>, com o texto por parametro.</summary>
    public static string InserirTextoWord(string texto) => NoEditor($$"""
            if (typeof api.asc_RemoveSelection === "function") api.asc_RemoveSelection();
            if (typeof api.asc_AddText === "function") { api.asc_AddText({{Js(texto)}}); return "ok asc_AddText"; }
            if (typeof api.asc_PasteData === "function") { api.asc_PasteData(1, {{Js(texto)}}); return "ok asc_PasteData"; }
            if (typeof api.pluginMethod_PasteText === "function") { api.pluginMethod_PasteText({{Js(texto)}}); return "ok pluginMethod_PasteText"; }
            return "sem API de insercao";
        """);

    /// <summary>Insercao no editor de PLANILHA (cell), numa celula escolhida (a origem fica intacta em A1).</summary>
    public static string InserirTextoCell(string texto, string celula) => NoEditor($$"""
            var passos = [];
            if (typeof api.asc_findCell !== "function") return "sem asc_findCell";
            api.asc_findCell({{Js(celula)}});
            passos.push("asc_findCell(" + {{Js(celula)}} + ")");
            if (typeof api.pluginMethod_InputText === "function") {
              api.pluginMethod_InputText({{Js(texto)}});
              passos.push("pluginMethod_InputText");
              if (typeof api.asc_closeCellEditor === "function") { api.asc_closeCellEditor(); passos.push("asc_closeCellEditor"); }
              return "ok " + passos.join(" -> ");
            }
            if (typeof api.pluginMethod_PasteText === "function") {
              api.pluginMethod_PasteText({{Js(texto)}});
              passos.push("pluginMethod_PasteText");
              return "ok " + passos.join(" -> ");
            }
            return "sem API de insercao";
        """);

    /// <summary>Insercao no editor de APRESENTACAO (slide).</summary>
    public static string InserirTextoSlide(string texto) => NoEditor($$"""
            if (typeof api.pluginMethod_PasteText === "function") { api.pluginMethod_PasteText({{Js(texto)}}); return "ok pluginMethod_PasteText"; }
            if (typeof api.pluginMethod_InputText === "function") { api.pluginMethod_InputText({{Js(texto)}}); return "ok pluginMethod_InputText"; }
            return "sem API de insercao";
        """);

    public static string InserirPorDocType(string doctype, string texto, string celula = "B1") => doctype switch
    {
        "cell" => InserirTextoCell(texto, celula),
        "slide" => InserirTextoSlide(texto),
        _ => InserirTextoWord(texto),
    };

    /// <summary>`asc_isDocumentCanSave()` — exportado pelos TRES editores (`Y.asc_isDocumentCanSave`,
    /// `d.prototype.Vze`, `h.prototype.TMe`). E' o sinal universal de "ha' alteracao nao salva": serve para
    /// esperar de forma DETERMINISTICA a colagem assincrona de cell/slide chegar ao modelo antes de salvar,
    /// em vez de um Task.Delay as cegas.</summary>
    public const string PodeSalvarNoEditor = """
        (function () {
          try {
            var w = window.frames[0];
            var api = (w.Asc && w.Asc.editor) ? w.Asc.editor : w.editor;
            if (!api || typeof api.asc_isDocumentCanSave !== "function") return "sem api";
            return String(api.asc_isDocumentCanSave());
          } catch (e) { return "erro: " + e; }
        })()
        """;

    /// <summary>Texto que o EDITOR (nao o arquivo) enxerga — `asc_EditSelectAll` + `asc_GetSelectedText`,
    /// exportados pelos tres. Usado so DEPOIS de salvar: em cell/slide o select-all mexe na selecao e
    /// atrapalharia uma insercao seguinte.</summary>
    public const string TextoSelecionado = """
        (function () {
          try {
            var w = window.frames[0];
            var api = (w.Asc && w.Asc.editor) ? w.Asc.editor : w.editor;
            if (!api) return "sem api";
            if (typeof api.asc_EditSelectAll === "function") api.asc_EditSelectAll();
            return String(api.asc_GetSelectedText());
          } catch (e) { return "erro: " + e; }
        })()
        """;
}
