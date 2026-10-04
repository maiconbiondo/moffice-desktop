using mOffice.Nucleo.Travas;

namespace mOffice.Nucleo.Tests;

/// <summary>Plano 6, Task 5 (SO-02/MvP-05): arquivo de dono <c>~$</c> no formato do Office, sondagem de acesso, trava do original e
/// carimbo. Puros (sem janela).</summary>
public class TravasTests
{
    private static string Pasta() { var p = Path.Combine(Path.GetTempPath(), "mOffice-testes", Guid.NewGuid().ToString("N")); Directory.CreateDirectory(p); return p; }

    [Theory] // nome do arquivo de dono como o Office: no Word, >= 8 letras antes da extensao troca as 2 primeiras; 7 troca 1; menos, prefixa.
             // Excel/PowerPoint: sempre "~$" + o nome inteiro (regra do LibreOffice, svl/source/misc/msodocumentlockfile.cxx — o brief dizia o contrario).
    [InlineData("Documento1.docx", "~$cumento1.docx")]
    [InlineData("Peticao.docx", "~$eticao.docx")]
    [InlineData("abc.docx", "~$abc.docx")]
    [InlineData("Contrato social.odt", "~$ntrato social.odt")]
    [InlineData("Planilha de custos.xlsx", "~$Planilha de custos.xlsx")]
    [InlineData("Apresentacao final.pptx", "~$Apresentacao final.pptx")]
    public void CaminhoDoDono_RegraDoOffice(string doc, string dono) =>
        Assert.Equal(dono, Path.GetFileName(ArquivoDeDono.CaminhoPara(Path.Combine(@"C:\x", doc))));

    [Fact] // conteudo: Word 162 bytes, Excel/PowerPoint 165; o nome volta pela leitura nos tres
    public void Montar_TamanhoELeitura()
    {
        var p = Pasta();
        foreach (var (app, tamanho) in new[] { (AppDoDono.Word, 162), (AppDoDono.Excel, 165), (AppDoDono.PowerPoint, 165) })
        {
            var b = ArquivoDeDono.Montar("Dra. Ana Souza", app);
            Assert.Equal(tamanho, b.Length);
            var a = Path.Combine(p, $"~$teste-{app}.docx"); File.WriteAllBytes(a, b);
            Assert.Equal("Dra. Ana Souza", ArquivoDeDono.LerNome(a));
        }
    }

    [Fact] // layout byte a byte (o do LibreOffice): Word = N, ANSI, 0x00 ate' 54, N 0, UTF-16, 0x00; Excel = N, ANSI, 0x20 ate' 55, N 0, UTF-16, "20 00"
    public void Montar_LayoutDoOffice()
    {
        var w = ArquivoDeDono.Montar("Joã", AppDoDono.Word);
        Assert.Equal(new byte[] { 3, (byte)'J', (byte)'o', 0xE3, 0 }, w[..5]);
        Assert.All(w[4..54], b => Assert.Equal(0, b));
        Assert.Equal(new byte[] { 3, 0, (byte)'J', 0, (byte)'o', 0, 0xE3, 0, 0, 0 }, w[54..64]);
        Assert.All(w[62..], b => Assert.Equal(0, b));

        var x = ArquivoDeDono.Montar("Joã", AppDoDono.Excel);
        Assert.Equal(new byte[] { 3, (byte)'J', (byte)'o', 0xE3, 0x20 }, x[..5]);
        Assert.All(x[4..55], b => Assert.Equal(0x20, b));
        Assert.Equal(new byte[] { 3, 0, (byte)'J', 0, (byte)'o', 0, 0xE3, 0, 0x20, 0 }, x[55..65]);

        var pp = ArquivoDeDono.Montar("Joã", AppDoDono.PowerPoint);
        Assert.Equal(new byte[] { 3, (byte)'J', (byte)'o', 0xE3, 0, 0x20 }, pp[..6]);   // PowerPoint: um 0x00 antes dos espacos
        Assert.Equal(new byte[] { 3, 0 }, pp[55..57]);
    }

    [Fact] // nome com mais de 52 caracteres e' cortado (limite do Office); fora do Latin-1 vira '?' no ANSI e fica certo no UTF-16
    public void Montar_LimitesDoNome()
    {
        var p = Pasta();
        var longo = new string('a', 60);
        var a = Path.Combine(p, "~$longo.docx"); File.WriteAllBytes(a, ArquivoDeDono.Montar(longo, AppDoDono.Word));
        Assert.Equal(new string('a', 52), ArquivoDeDono.LerNome(a));
        var b = Path.Combine(p, "~$grego.docx"); File.WriteAllBytes(b, ArquivoDeDono.Montar("Ωmega", AppDoDono.Word));
        Assert.Equal((byte)'?', File.ReadAllBytes(b)[1]);
        Assert.Equal("Ωmega", ArquivoDeDono.LerNome(b));
    }

    [Fact] // ~$ ilegivel, vazio ou inexistente -> null (nunca lanca: so' decide o texto da pergunta)
    public void LerNome_Invalido_Nulo()
    {
        var p = Pasta();
        Assert.Null(ArquivoDeDono.LerNome(Path.Combine(p, "nao-existe")));
        var vazio = Path.Combine(p, "~$vazio.docx"); File.WriteAllBytes(vazio, []);
        Assert.Null(ArquivoDeDono.LerNome(vazio));
        var lixo = Path.Combine(p, "~$lixo.docx"); File.WriteAllBytes(lixo, [0, 1, 2]);
        Assert.Null(ArquivoDeDono.LerNome(lixo));
    }

    [Fact] // atributo somente leitura -> SomenteLeituraAtributo; livre -> Livre
    public void Sondar_AtributoSomenteLeitura()
    {
        var a = Path.Combine(Pasta(), "a.docx"); File.WriteAllText(a, "x");
        Assert.Equal(EstadoAcesso.Livre, AcessoAoArquivo.Sondar(a));
        File.SetAttributes(a, FileAttributes.ReadOnly);
        try { Assert.Equal(EstadoAcesso.SomenteLeituraAtributo, AcessoAoArquivo.Sondar(a)); }
        finally { File.SetAttributes(a, FileAttributes.Normal); }
    }

    [Fact] // outro processo com o arquivo aberto para escrita (simula o Word) -> EmUso
    public void Sondar_AbertoParaEscritaPorOutro_EmUso()
    {
        var a = Path.Combine(Pasta(), "a.docx"); File.WriteAllText(a, "x");
        using var outro = new FileStream(a, FileMode.Open, FileAccess.ReadWrite, FileShare.Read);
        Assert.Equal(EstadoAcesso.EmUso, AcessoAoArquivo.Sondar(a));
    }

    [Fact] // a sondagem nao e' bloqueada pela trava de OUTRA sessao? E' sim: o original travado por outra janela/pessoa esta' em uso
    public void Sondar_TravadoPorOutraTrava_EmUso()
    {
        var a = Path.Combine(Pasta(), "a.docx"); File.WriteAllText(a, "x");
        using var t = TravaDoDocumento.Adquirir(a, "Outro", _ => { })!;
        Assert.Equal(EstadoAcesso.EmUso, AcessoAoArquivo.Sondar(a));
    }

    [Fact] // arquivo que nao existe: Livre (quem reclama e' o x2t, com a mensagem de sempre)
    public void Sondar_Inexistente_Livre() =>
        Assert.Equal(EstadoAcesso.Livre, AcessoAoArquivo.Sondar(Path.Combine(Pasta(), "nao-existe.docx")));

    [Fact] // a trava impede escrita de terceiros, cria o ~$ oculto e, ao soltar/descartar, libera e apaga
    public void Trava_ImpedeEscritaECriaDono()
    {
        var a = Path.Combine(Pasta(), "Contrato.docx"); File.WriteAllText(a, "x");
        using (var t = TravaDoDocumento.Adquirir(a, "Fulano", _ => { })!)
        {
            Assert.Throws<IOException>(() => new FileStream(a, FileMode.Open, FileAccess.Write, FileShare.ReadWrite).Dispose());
            var dono = ArquivoDeDono.CaminhoPara(a);
            Assert.True(File.GetAttributes(dono).HasFlag(FileAttributes.Hidden));
            Assert.Equal("Fulano", ArquivoDeDono.LerNome(dono));
            t.Soltar();
            new FileStream(a, FileMode.Open, FileAccess.Write, FileShare.ReadWrite).Dispose();   // solta = escrita liberada
            Assert.True(t.Retomar());
            Assert.Throws<IOException>(() => new FileStream(a, FileMode.Open, FileAccess.Write, FileShare.ReadWrite).Dispose());
        }
        Assert.False(File.Exists(ArquivoDeDono.CaminhoPara(a)));
        new FileStream(a, FileMode.Open, FileAccess.Write, FileShare.ReadWrite).Dispose();
    }

    [Fact] // a trava segura so' leitura: o colega le' o documento (abre "somente leitura") e le' o nome no ~$
    public void Trava_OutroConsegueLerODocumentoEODono()
    {
        var a = Path.Combine(Pasta(), "Contrato.docx"); File.WriteAllText(a, "conteudo");
        using var t = TravaDoDocumento.Adquirir(a, "Fulano", _ => { })!;
        using (var leitura = new FileStream(a, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) Assert.Equal(8, leitura.Length);
        using var dono = new FileStream(ArquivoDeDono.CaminhoPara(a), FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        Assert.Equal(162, dono.Length);
    }

    [Fact] // MEDIDO: segurando so' FileShare.Read, o File.Replace do proprio processo falha (o salvar nao pode segurar a trava assim);
           // durante a troca a trava passa a um handle Read+Delete (terceiros continuam sem gravar), o Substituir funciona e Retomar
           // volta a travar o arquivo NOVO
    public void Trava_SalvamentoComATravaNaMao_SubstituiERetrava()
    {
        var p = Pasta();
        var a = Path.Combine(p, "Contrato.docx"); File.WriteAllText(a, "velho");
        using var t = TravaDoDocumento.Adquirir(a, "Fulano", _ => { })!;
        // controle negativo: com a trava normal o Replace do proprio processo e' recusado
        var tmp0 = Path.Combine(p, "Contrato.docx.moffice-tmp"); File.WriteAllText(tmp0, "novo");
        Assert.ThrowsAny<IOException>(() => File.Replace(tmp0, a, a + ".bak-teste", true));
        File.Delete(tmp0);

        Assert.True(t.PrepararSubstituicao());
        Assert.Throws<IOException>(() => new FileStream(a, FileMode.Open, FileAccess.Write, FileShare.ReadWrite).Dispose());   // ainda travado
        var tmp = Path.Combine(p, "Contrato.docx.moffice-tmp"); File.WriteAllText(tmp, "novo");
        Assert.Equal(ResultadoGravacao.MetadadosPreservados, GravacaoNoDestino.Substituir(tmp, a));
        Assert.True(t.Retomar());
        Assert.Equal("novo", File.ReadAllText(a));
        Assert.Empty(Directory.GetFiles(p, "*.moffice-bak*"));
        Assert.Throws<IOException>(() => new FileStream(a, FileMode.Open, FileAccess.Write, FileShare.ReadWrite).Dispose());   // o NOVO esta' travado
    }

    [Fact] // a trava muda de arquivo no "Salvar como": o antigo fica livre (e sem ~$), o novo travado e com ~$
    public void Trava_Mover_TravaODestino()
    {
        var p = Pasta();
        var a = Path.Combine(p, "Antigo1.docx"); File.WriteAllText(a, "x");
        var b = Path.Combine(p, "Novo1234.docx"); File.WriteAllText(b, "y");
        using var t = TravaDoDocumento.Adquirir(a, "Fulano", _ => { })!;
        Assert.True(t.Mover(b));
        new FileStream(a, FileMode.Open, FileAccess.Write, FileShare.ReadWrite).Dispose();
        Assert.False(File.Exists(ArquivoDeDono.CaminhoPara(a)));
        Assert.Throws<IOException>(() => new FileStream(b, FileMode.Open, FileAccess.Write, FileShare.ReadWrite).Dispose());
        Assert.Equal("Fulano", ArquivoDeDono.LerNome(ArquivoDeDono.CaminhoPara(b)));
    }

    [Fact] // sobra de queda: um ~$ (de outra pessoa ou nosso, oculto) SEM ninguem segurando nao impede a trava — e' sobrescrito com o nosso nome
    public void Trava_DonoVelhoDeQueda_ESobrescrito()
    {
        var a = Path.Combine(Pasta(), "Contrato.docx"); File.WriteAllText(a, "x");
        var dono = ArquivoDeDono.CaminhoPara(a);
        File.WriteAllBytes(dono, ArquivoDeDono.Montar("Outro", AppDoDono.Word));
        File.SetAttributes(dono, FileAttributes.Hidden);
        Assert.Equal(EstadoAcesso.Livre, AcessoAoArquivo.Sondar(a));   // o ~$ sozinho nao manda, como no Word
        using (var t = TravaDoDocumento.Adquirir(a, "Fulano", _ => { })!)
            Assert.Equal("Fulano", ArquivoDeDono.LerNome(dono));
        Assert.False(File.Exists(dono));
    }

    [Fact] // ~$ preso por outro processo (o colega ainda esta' com ele aberto): a trava do documento vale, o ~$ dele fica intacto
    public void Trava_DonoPresoPorOutro_NaoMexeNele()
    {
        var a = Path.Combine(Pasta(), "Contrato.docx"); File.WriteAllText(a, "x");
        var dono = ArquivoDeDono.CaminhoPara(a);
        File.WriteAllBytes(dono, ArquivoDeDono.Montar("Outro", AppDoDono.Word));
        var log = new List<string>();
        using (var preso = new FileStream(dono, FileMode.Open, FileAccess.ReadWrite, FileShare.Read))
        using (var t = TravaDoDocumento.Adquirir(a, "Fulano", log.Add)!)
        {
            Assert.Throws<IOException>(() => new FileStream(a, FileMode.Open, FileAccess.Write, FileShare.ReadWrite).Dispose());
            Assert.Contains(log, l => l.Contains("~$"));
        }
        Assert.Equal("Outro", ArquivoDeDono.LerNome(dono));
    }

    [Fact] // documento que outro processo segura para escrita: Adquirir devolve null (com log) — a sessao segue sem trava
    public void Trava_DocumentoPresoParaEscrita_Nulo()
    {
        var a = Path.Combine(Pasta(), "Contrato.docx"); File.WriteAllText(a, "x");
        var log = new List<string>();
        using var outro = new FileStream(a, FileMode.Open, FileAccess.ReadWrite, FileShare.Read);
        Assert.Null(TravaDoDocumento.Adquirir(a, "Fulano", log.Add));
        Assert.NotEmpty(log);
        Assert.False(File.Exists(ArquivoDeDono.CaminhoPara(a)));
    }

    [Fact] // carimbo muda quando o arquivo muda
    public void Carimbo_DetectaMudanca()
    {
        var a = Path.Combine(Pasta(), "a.docx"); File.WriteAllText(a, "x");
        var c1 = CarimboArquivo.De(a);
        File.SetLastWriteTimeUtc(a, DateTime.UtcNow.AddMinutes(5)); File.AppendAllText(a, "y");
        Assert.NotEqual(c1, CarimboArquivo.De(a));
        Assert.Null(CarimboArquivo.De(Path.Combine(Pasta(), "nao-existe")));
    }

    [Fact] // assinatura digital: OOXML com _xmlsignatures/ ou ODF com META-INF/documentsignatures.xml; zip invalido = false
    public void TemAssinaturaDigital_OoxmlEOdf()
    {
        var p = Pasta();
        string Zip(string nome, params string[] entradas)
        {
            var c = Path.Combine(p, nome);
            using var z = System.IO.Compression.ZipFile.Open(c, System.IO.Compression.ZipArchiveMode.Create);
            foreach (var e in entradas) using (var s = z.CreateEntry(e).Open()) s.WriteByte(1);
            return c;
        }
        Assert.True(AcessoAoArquivo.TemAssinaturaDigital(Zip("a.docx", "word/document.xml", "_xmlsignatures/sig1.xml")));
        Assert.True(AcessoAoArquivo.TemAssinaturaDigital(Zip("a.odt", "content.xml", "META-INF/documentsignatures.xml")));
        Assert.False(AcessoAoArquivo.TemAssinaturaDigital(Zip("b.docx", "word/document.xml")));
        Assert.False(AcessoAoArquivo.TemAssinaturaDigital(Zip("b.odt", "content.xml", "META-INF/manifest.xml")));
        var lixo = Path.Combine(p, "lixo.docx"); File.WriteAllText(lixo, "nao e' zip");
        Assert.False(AcessoAoArquivo.TemAssinaturaDigital(lixo));
    }

    // --- Fix 1 da revisao da Task 5 ---------------------------------------------------------------------------------------------

    [Fact] // MINOR 1: o aviso "aposTroca" do Substituir roda logo depois da troca, ANTES de limpar o backup (o salvar retrava e tira o
           // carimbo ali — uma gravacao de terceiros no intervalo vira conflito, nunca a nova referencia); tambem no destino novo
    public void Substituir_AposTroca_RodaAntesDeLimparOBackup()
    {
        var p = Pasta();
        var a = Path.Combine(p, "a.docx"); File.WriteAllText(a, "velho");
        var tmp = Path.Combine(p, "a.docx.moffice-tmp"); File.WriteAllText(tmp, "novo");
        int backupsNaHora = -1; string? conteudoNaHora = null;
        GravacaoNoDestino.Substituir(tmp, a, aposTroca: () => { backupsNaHora = Directory.GetFiles(p, "*.moffice-bak*").Length; conteudoNaHora = File.ReadAllText(a); });
        Assert.Equal(1, backupsNaHora);
        Assert.Equal("novo", conteudoNaHora);
        Assert.Empty(Directory.GetFiles(p, "*.moffice-bak*"));
        var b = Path.Combine(p, "b.docx"); var tmpB = Path.Combine(p, "b.docx.moffice-tmp"); File.WriteAllText(tmpB, "x");
        var chamadas = 0;
        Assert.Equal(ResultadoGravacao.Novo, GravacaoNoDestino.Substituir(tmpB, b, aposTroca: () => chamadas++));
        Assert.Equal(1, chamadas);
    }

    [Fact] // MINOR 4: um ~$ OCULTO de outra pessoa, preso por ela, nunca e' desocultado (antes o SetAttributes vinha antes da abertura)
    public void Criar_DonoDeOutroPreso_NaoMexeNosAtributos()
    {
        var a = Path.Combine(Pasta(), "Contrato.docx"); File.WriteAllText(a, "x");
        var dono = ArquivoDeDono.CaminhoPara(a);
        File.WriteAllBytes(dono, ArquivoDeDono.Montar("Outro", AppDoDono.Word));
        File.SetAttributes(dono, FileAttributes.Hidden);
        using (new FileStream(dono, FileMode.Open, FileAccess.ReadWrite, FileShare.Read))
        {
            Assert.Null(ArquivoDeDono.Criar(a, "Fulano"));
            Assert.Equal(FileAttributes.Hidden, File.GetAttributes(dono) & (FileAttributes.Hidden | FileAttributes.Normal));
        }
        Assert.Equal("Outro", ArquivoDeDono.LerNome(dono));
    }

    [Fact] // MINOR 5: Mover para um arquivo que (ainda) nao da' para travar devolve false; o proximo Retomar que der certo recria o ~$
    public void Trava_MoverFalha_RecriaODonoNoRetomar()
    {
        var p = Pasta();
        var a = Path.Combine(p, "Antigo1.docx"); File.WriteAllText(a, "x");
        var b = Path.Combine(p, "Novo1234.docx");
        using var t = TravaDoDocumento.Adquirir(a, "Fulano", _ => { })!;
        Assert.False(t.Mover(b));                 // b nao existe
        Assert.False(t.Segurando);
        File.WriteAllText(b, "y");
        Assert.True(t.Retomar());
        Assert.False(PodeGravarPuro(b));
        Assert.Equal("Fulano", ArquivoDeDono.LerNome(ArquivoDeDono.CaminhoPara(b)));
    }

    private static bool PodeGravarPuro(string arquivo)
    {
        try { new FileStream(arquivo, FileMode.Open, FileAccess.Write, FileShare.ReadWrite).Dispose(); return true; }
        catch (IOException) { return false; }
    }
}
