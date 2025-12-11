using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using RtfPipe;

namespace HelloBlazor.Components.Services
{
    public class ResultadoExtracao
    {
        public List<string> ImagensBase64 { get; set; } = new List<string>();
        public List<string> Logs { get; set; } = new List<string>();
        public string HtmlConteudo { get; set; } = "";
    }

    public class Extrator
    {
        public static ResultadoExtracao ProcessarRtf(string rtfConteudo)
        {
            var resultado = new ResultadoExtracao();
            void Log(string msg) => resultado.Logs.Add($"[{DateTime.Now:HH:mm:ss}] {msg}");

            if (string.IsNullOrWhiteSpace(rtfConteudo)) return resultado;

            string rtfLimpo = rtfConteudo.Replace("\0", "");

            // ==============================================================================
            // PASSO 1: EXTRAIR O MIOLO (TEXTO) - COMO NO PROJETO 2
            // ==============================================================================
            // Usamos o RTF original. O RtfPipe vai ignorar header/footer mas vai ler o texto perfeito.
            string htmlMiolo = "";
            try
            {
                string htmlBruto = Rtf.ToHtml(rtfLimpo);

                // LIMPEZA CRÍTICA: Removemos <html>, <head> e <body> para sobrar só o conteúdo.
                // Se não fizer isso, o navegador cria um "buraco" branco ao renderizar HTML aninhado.
                htmlMiolo = LimparHtmlContainer(htmlBruto);

                Log($"Miolo extraído com sucesso ({htmlMiolo.Length} chars).");
            }
            catch (Exception ex)
            {
                htmlMiolo = "<p style='color:red'>[Falha ao ler o texto do laudo]</p>";
                Log($"Erro no miolo: {ex.Message}");
            }

            // ==============================================================================
            // PASSO 2: EXTRAIR A ESTRUTURA (CABEÇALHO E RODAPÉ) - COMO NO PROJETO 1
            // ==============================================================================
            // Hackeamos o RTF para transformar Header e Footer em texto comum.
            string htmlEstrutura = "";
            try
            {
                string rtfHack = rtfLimpo;

                // Força cabeçalho a aparecer
                if (rtfHack.Contains(@"\header"))
                    rtfHack = rtfHack.Replace(@"\header", @"\pard");

                // Força rodapé a aparecer (com quebra de linha visual)
                if (rtfHack.Contains(@"\footer"))
                    rtfHack = rtfHack.Replace(@"\footer", @"\pard\brdrt\brdrs\brdrw10\brsp100 ");

                htmlEstrutura = Rtf.ToHtml(rtfHack);
                Log("Estrutura (Paciente/Assinatura) gerada.");
            }
            catch (Exception ex) { Log($"Erro na estrutura: {ex.Message}"); }

            // ==============================================================================
            // PASSO 3: A FUSÃO (FRANKENSTEIN)
            // ==============================================================================
            if (!string.IsNullOrEmpty(htmlEstrutura))
            {
                // O Cabeçalho do paciente é sempre a primeira tabela.
                // Injetamos o Miolo logo depois dela.
                int fimTabelaCabecalho = htmlEstrutura.IndexOf("</table>");

                if (fimTabelaCabecalho > 0)
                {
                    Log("Fusão: Inserindo texto do laudo após o cabeçalho.");
                    // Injeta o miolo limpo dentro da estrutura
                    resultado.HtmlConteudo = htmlEstrutura.Insert(fimTabelaCabecalho + 8,
                        $"<div class='miolo-do-laudo' style='margin: 20px 0; padding: 10px;'>{htmlMiolo}</div>");
                }
                else
                {
                    // Se não achou tabela, concatena (Header + Miolo)
                    resultado.HtmlConteudo = htmlEstrutura + "<hr/>" + htmlMiolo;
                }
            }
            else
            {
                resultado.HtmlConteudo = htmlMiolo; // Se falhou a estrutura, mostra só o texto
            }

            // ==============================================================================
            // PASSO 4: IMAGENS (OLE / BITMAP)
            // ==============================================================================
            Log("Iniciando extração de imagens...");
            try
            {
                using (StringReader sr = new StringReader(rtfConteudo))
                {
                    RtfReader reader = new RtfReader(sr);
                    var enumerator = reader.Read().GetEnumerator();

                    while (enumerator.MoveNext())
                    {
                        // OLE (Anexos)
                        if (enumerator.Current is RtfControlWord cw && cw.Text == "object")
                        {
                            if (RtfReader.MoveToNextControlWord(enumerator, "objclass") &&
                                RtfReader.GetNextText(enumerator) == "Package" &&
                                RtfReader.MoveToNextControlWord(enumerator, "objdata"))
                            {
                                byte[] oleData = RtfReader.GetNextTextAsByteArray(enumerator, resultado.Logs);
                                ProcessarOlePackage(oleData, resultado, Log);
                            }
                        }
                        // PICT (Assinaturas/QR)
                        else if (enumerator.Current is RtfControlWord cwPict && cwPict.Text == "pict")
                        {
                            bool isDib = false;
                            string type = "jpeg";

                            while (enumerator.MoveNext())
                            {
                                if (enumerator.Current is RtfControlWord typeCw)
                                {
                                    if (typeCw.Text.StartsWith("dibitmap")) { isDib = true; type = "bmp"; break; }
                                    if (typeCw.Text.StartsWith("png")) { type = "png"; break; }
                                    if (typeCw.Text.StartsWith("jpeg")) { type = "jpeg"; break; }
                                    if (typeCw.Text.StartsWith("wmetafile")) { type = "wmf"; break; }
                                }
                                if (enumerator.Current is RtfText) break;
                            }

                            if (type != "wmf")
                            {
                                byte[] imgData = RtfReader.GetNextTextAsByteArray(enumerator, resultado.Logs);
                                if (imgData != null && imgData.Length > 0)
                                {
                                    if (isDib) imgData = ConvertDibToBmp(imgData);
                                    string base64 = Convert.ToBase64String(imgData);

                                    string mime = type == "png" ? "image/png" : "image/bmp";
                                    if (type == "jpeg") mime = "image/jpeg";

                                    resultado.ImagensBase64.Add($"data:{mime};base64,{base64}");
                                    Log($"Imagem ({type}) recuperada.");
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex) { Log($"Erro imagens: {ex.Message}"); }

            // ==============================================================================
            // PASSO 5: TRANSPLANTE VISUAL
            // ==============================================================================
            if (resultado.ImagensBase64.Count > 0)
            {
                Log($"Injetando {resultado.ImagensBase64.Count} imagens...");
                resultado.HtmlConteudo = InjetarImagensNoHtml(resultado.HtmlConteudo, resultado.ImagensBase64);
            }

            return resultado;
        }

        // --- MÉTODOS AUXILIARES ---

        // Remove as tags de estrutura HTML para permitir aninhamento
        private static string LimparHtmlContainer(string html)
        {
            if (string.IsNullOrEmpty(html)) return "";

            // Remove tudo antes da abertura do body
            int bodyStart = html.IndexOf("<body>");
            if (bodyStart >= 0)
            {
                html = html.Substring(bodyStart + 6); // Pula "<body>"
            }

            // Remove o fechamento do body e html
            html = html.Replace("</body>", "").Replace("</html>", "");

            return html;
        }

        private static string InjetarImagensNoHtml(string html, List<string> imagensBase64)
        {
            int indexImagem = 0;
            return Regex.Replace(html, "<img[^>]+src=[\"']([^\"']*)[\"'][^>]*>", match =>
            {
                if (indexImagem < imagensBase64.Count)
                {
                    string tagOriginal = match.Value;
                    string srcNovo = imagensBase64[indexImagem];
                    indexImagem++;

                    string tagNova = tagOriginal.Replace(match.Groups[1].Value, srcNovo);
                    if (!tagNova.Contains("style"))
                        tagNova = tagNova.Replace("<img", "<img style='max-width: 150px; display:block;'");
                    return tagNova;
                }
                return match.Value;
            });
        }

        private static void ProcessarOlePackage(byte[] oleData, ResultadoExtracao resultado, Action<string> Log)
        {
            if (oleData == null || oleData.Length == 0) return;
            try
            {
                using (MemoryStream oleStream = new MemoryStream(oleData))
                using (MemoryStream dataStream = new MemoryStream())
                {
                    PackagedObject.ExtractObjectData(oleStream, dataStream);
                    if (dataStream.Length > 0)
                    {
                        PackagedObject arquivoFinal = PackagedObject.Extract(dataStream);
                        resultado.ImagensBase64.Add("data:image/jpeg;base64," + Convert.ToBase64String(arquivoFinal.Data));
                    }
                }
            }
            catch { }
        }

        private static byte[] ConvertDibToBmp(byte[] dibData)
        {
            int headerSize = 14;
            int fileSize = headerSize + dibData.Length;
            byte[] bmp = new byte[fileSize];
            bmp[0] = 0x42; bmp[1] = 0x4D;
            BitConverter.GetBytes(fileSize).CopyTo(bmp, 2);
            bmp[10] = 54;
            if (dibData.Length > 4)
            {
                int dibHeaderSize = BitConverter.ToInt32(dibData, 0);
                BitConverter.GetBytes(14 + dibHeaderSize).CopyTo(bmp, 10);
            }
            Array.Copy(dibData, 0, bmp, 14, dibData.Length);
            return bmp;
        }
    }
}