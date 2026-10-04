using mOffice.Conversor;

namespace mOffice.Nucleo;

/// <summary>
/// O que o usuario escolheu no dialogo "Salvar como" do host: o caminho, o formato e — so' na linha
/// "PDF assinado com mPDF" — o pedido de assinar o PDF depois de exportado (Plano 4, Task 2).
///
/// O formato sai do INDICE da linha do filtro, nao da extensao: as tres linhas de PDF (PDF, PDF/A e
/// "PDF assinado com mPDF") tem a mesma extensao <c>.pdf</c>. Ver <see cref="JanelaEditor.DestinoDoDialogo"/>.
/// </summary>
/// <param name="Caminho">Arquivo de destino (com extensao).</param>
/// <param name="Formato">Formato que o x2t grava.</param>
/// <param name="Assinar">Depois de exportar o PDF com sucesso, pedir a assinatura pelo mPDF
/// (<see cref="JanelaEditor.AssinarPedido"/>). Com formato de DOCUMENTO (a linha "PDF assinado" com extensao de documento
/// digitada) nada e' assinado: o host so' avisa (<see cref="JanelaEditor.MensagemAssinarSalvouDocumento"/>).</param>
/// <param name="Texto">Codificacao e delimitador do CSV/TXT (Plano 6, Task 3): vem da linha do dialogo (ANSI/UTF-8) ou do padrao do
/// formato quando a extensao foi digitada; <c>null</c> em qualquer outro formato.</param>
public sealed record DestinoSalvar(string Caminho, FormatoDocumento Formato, bool Assinar = false, OpcoesTexto? Texto = null);
