namespace mPlanilha;

/// <summary>mPlanilha: editor de planilhas (XLSX e ODS). Todo o comportamento esta' em
/// <see cref="mOffice.Comum.AplicativoOffice"/>; aqui so' se diz QUAL app da suite este executavel e'.</summary>
public partial class App : mOffice.Comum.AplicativoOffice
{
    public App() : base(mOffice.Comum.PerfilApp.Planilha) { }
}
