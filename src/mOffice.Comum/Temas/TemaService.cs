using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Threading;
using mOffice.Nucleo;

namespace mOffice.Comum;

public enum ModoTema { Sistema = 0, Claro = 1, Escuro = 2 }

/// <summary>Troca os tokens (Cor.*) ao vivo, como o ThemeService do mPDF, mais o modo "seguir o Windows".</summary>
public sealed class TemaService
{
    private const string UriClaro = "pack://application:,,,/mOffice.Comum;component/Temas/Tokens.Claro.xaml";
    private const string UriEscuro = "pack://application:,,,/mOffice.Comum;component/Temas/Tokens.Escuro.xaml";
    private readonly Collection<ResourceDictionary> _alvo;
    private readonly Func<bool> _lerAppsClaros;
    private readonly Dispatcher _dispatcher;
    private ModoTema _modo = ModoTema.Sistema;
    private bool _assinado;

    public TemaService(Collection<ResourceDictionary> alvo) : this(alvo, TemaDoSistema.AppsClaros) { }
    internal TemaService(Collection<ResourceDictionary> alvo, Func<bool> lerAppsClaros)
    { _alvo = alvo; _lerAppsClaros = lerAppsClaros; _dispatcher = Dispatcher.CurrentDispatcher; }

    public static TemaService DoApp() => new(Application.Current.Resources.MergedDictionaries);

    public bool EscuroEfetivo { get; private set; }
    public TemaEditor TemaDoEditor => EscuroEfetivo ? TemaEditor.Escuro : TemaEditor.Claro;

    /// <summary>Disparado depois de Reaplicar() trocar o dicionario de tokens. SEMPRE na thread do
    /// Dispatcher capturado no construtor (a thread de UI que criou o TemaService) — mesmo quando
    /// SistemaMudou() chega de uma thread de worker do SystemEvents, o evento e' reencaminhado (ver
    /// SistemaMudou) antes de disparar aqui, entao quem assina pode mexer em UI direto no handler.</summary>
    public event Action<bool>? EfetivoAlterado;

    public void Aplicar(ModoTema modo)
    {
        _modo = modo;
        if (modo == ModoTema.Sistema && !_assinado) { TemaDoSistema.Alterado += SistemaMudou; _assinado = true; }
        if (modo != ModoTema.Sistema && _assinado) { TemaDoSistema.Alterado -= SistemaMudou; _assinado = false; }
        Reaplicar();
    }

    /// <summary>O que SystemEvents.UserPreferenceChanged dispara; publico para teste e para a Sonda.
    /// SystemEvents chama isto numa thread de WORKER dele, nunca a thread de UI — mexer em
    /// _alvo (MergedDictionaries de verdade em producao) fora da thread do Dispatcher lanca
    /// InvalidOperationException (cross-thread). Por isso reencaminha pro Dispatcher capturado no
    /// construtor quando necessario; se ja estamos nela (CheckAccess), roda direto (sem post,
    /// sem trocar a ordem de execucao pro caso mais comum dos testes/Sonda).</summary>
    public void SistemaMudou()
    {
        if (_modo != ModoTema.Sistema) return;
        if (_dispatcher.CheckAccess()) Reaplicar();
        else _dispatcher.BeginInvoke(Reaplicar);
    }

    private void Reaplicar()
    {
        var escuro = _modo switch { ModoTema.Escuro => true, ModoTema.Claro => false, _ => !_lerAppsClaros() };
        for (var i = _alvo.Count - 1; i >= 0; i--)
            if (_alvo[i].Source?.OriginalString.Contains("Tokens.", StringComparison.Ordinal) == true) _alvo.RemoveAt(i);
        _alvo.Insert(0, new ResourceDictionary { Source = new Uri(escuro ? UriEscuro : UriClaro) });
        var mudou = EscuroEfetivo != escuro || !_jaAplicado;
        EscuroEfetivo = escuro; _jaAplicado = true;
        if (mudou) EfetivoAlterado?.Invoke(escuro);
    }
    private bool _jaAplicado;
}
