namespace mTexto;

/// <summary>mTexto: editor de textos (DOCX e ODT). Todo o comportamento esta' em
/// <see cref="mOffice.Comum.AplicativoOffice"/>; aqui so' se diz QUAL app da suite este executavel e'.</summary>
public partial class App : mOffice.Comum.AplicativoOffice
{
    public App() : base(mOffice.Comum.PerfilApp.Texto) { }
}
