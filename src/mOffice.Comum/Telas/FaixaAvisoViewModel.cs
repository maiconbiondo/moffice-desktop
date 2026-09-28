using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

namespace mOffice.Comum.Telas;

/// <summary>Faixa de aviso acima do editor. Substituiu a barra de comandos do host (pedido do usuario em
/// 2026-09-25: Novo/Abrir/Salvar/Salvar como ja' existem no menu Arquivo do proprio editor, e a barra so'
/// duplicava). Fica INVISIVEL enquanto <see cref="Aviso"/> e' nulo; aparece para o que o usuario precisa
/// saber e o editor nao mostra: midias indisponiveis, salvar que falhou, fechar que nao pode concluir.</summary>
public sealed partial class FaixaAvisoViewModel : ObservableObject
{
    [ObservableProperty] private string? _aviso;

    /// <summary>O "x" da faixa: o usuario leu, a faixa some ate' o proximo aviso.</summary>
    [RelayCommand] private void Dispensar() => Aviso = null;
}
