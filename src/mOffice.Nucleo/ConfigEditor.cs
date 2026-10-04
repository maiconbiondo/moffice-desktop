namespace mOffice.Nucleo;

/// <summary>O que o host injeta no editorConfig (CE-7, Plano 6). Lido pelo window.__moffice.mesclarConfig da ponte, que o
/// patch 01 chama sobre o que o index.html do motor monta. Tudo documentado em api.onlyoffice.com (editorConfig.region,
/// customization.help/feedback/suggestFeature/macros/features.featuresTips) e conferido no app.js 9.4.
///
/// DONO DE CADA CAMPO (nao duplicar): a URL de entrada (<see cref="JanelaEditor.MontarUrlDeEntrada(mOffice.Conversor.FormatoDocumento, string, TemaEditor, Autor, string?, bool)"/>,
/// lida pelo index.html do motor) e' dona de user (userid/username), lang, mode, title/doctype/filetype e customization.uiTheme;
/// este JSON e' dono de region e do que DESLIGA recursos; e a ponte (estado.mesclarConfig) e' dona de customization.goback, que
/// depende do <c>&amp;arquivo=</c> da URL. Nenhum campo aparece nos dois canais (ConfigEditorTests confere).</summary>
public static class ConfigEditor
{
    public static string Json() =>
        """{"region":"pt-BR","customization":{"help":false,"feedback":false,"suggestFeature":false,"macros":false,"features":{"featuresTips":false}}}""";
}
