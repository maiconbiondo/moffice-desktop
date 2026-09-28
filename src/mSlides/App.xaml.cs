namespace mSlides;

/// <summary>mSlides: editor de apresentacoes (PPTX e ODP). Todo o comportamento esta' em
/// <see cref="mOffice.Comum.AplicativoOffice"/>; aqui so' se diz QUAL app da suite este executavel e'.</summary>
public partial class App : mOffice.Comum.AplicativoOffice
{
    public App() : base(mOffice.Comum.PerfilApp.Slides) { }
}
