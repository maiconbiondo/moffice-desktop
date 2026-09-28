namespace mOffice.Nucleo;

/// <summary>
/// O que um pedido de salvamento do host REALMENTE fez. Existe por causa do achado Critico da revisao final
/// do Plano 2: <see cref="JanelaEditor.SalvarComoAsync"/> devolvia <c>Task</c> nua, entao o caminho de
/// FECHAMENTO ("Salvar" no dialogo de fechar) nao tinha como distinguir "gravou o documento" de "exportou
/// para PDF" — e uma exportacao nao reponta a sessao nem preserva o documento editavel, mas o fechamento
/// seguia em frente e apagava a pasta de recuperacao. O documento do usuario sumia, restando so' o PDF.
///
/// <paramref name="Gravou"/> e' false so' quando o usuario CANCELOU o dialogo de destino (nao e' erro: erro
/// sai por excecao). <paramref name="Exportou"/> e' true quando o formato gravado e' de exportacao
/// (<see cref="mOffice.Conversor.FormatoDocumentoExtensoes.EhExportacao"/>: PDF/PDF-A) — nesse caso a sessao
/// continua no documento de origem e o evento disparado foi <see cref="JanelaEditor.Exportado"/>, nao
/// <see cref="JanelaEditor.Salvo"/>.
/// </summary>
public readonly record struct ResultadoSalvar(bool Gravou, bool Exportou)
{
    /// <summary>O usuario fechou o dialogo de destino sem escolher nada.</summary>
    public static ResultadoSalvar Cancelado => new(false, false);
}
