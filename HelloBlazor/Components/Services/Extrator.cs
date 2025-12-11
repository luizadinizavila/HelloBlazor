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

            // Limpeza inicial de nulos
            string rtfLimpo = rtfConteudo.Replace("\0", "");

            // ==============================================================================
            // FASE 1: O MIOLO (TEXTO PERFEITO) - Lógica do "Projeto 2"
            // ==============================================================================
            // Usamos o RTF original. O RtfPipe vai ignorar o cabeçalho/rodapé, 
            // mas vai ler o texto do exame ("TOMOGRAFIA...") perfeitamente.
            string htmlMiolo = "";
            try
            {
                // Gera o HTML padrão
                string htmlBruto = Rtf.ToHtml(rtfLimpo);

                // Remove as tags de estrutura (html, body) para sobrar apenas o conteúdo puro
                htmlMiolo = LimparHtmlContainer(htmlBruto);

                Log("Miolo (Texto do Exame) extraído.");
            }
            catch (Exception ex)
            {
                Log($"Erro no miolo: {ex.Message}");
                htmlMiolo = "<p>[Erro ao ler texto]</p>";
            }

            // ==============================================================================
            // FASE 2: A MOLDURA (CABEÇALHO + RODAPÉ) - Lógica do "Projeto 1" (Melhorada)
            // ==============================================================================
            string htmlMoldura = "";
            try
            {
                string rtfHack = rtfLimpo;

                // --- FIX DO "y2693" ---
                // Remove comandos de posicionamento vertical que vazam para o texto
                // Remove \footeryXXXX e \headeryXXXX
                rtfHack = Regex.Replace(rtfHack, @"\\(header|footer)[yx]\d+", "");
                // ----------------------

                // Transforma Header em texto normal (\pard)
                if (rtfHack.Contains(@"\header"))
                    rtfHack = rtfHack.Replace(@"\header", @"\pard");

                // Transforma Footer em texto normal (com linha separadora)
                if (rtfHack.Contains(@"\footer"))
                    rtfHack = rtfHack.Replace(@"\footer", @"\pard\brdrt\brdrs\brdrw10\brsp100 ");

                htmlMoldura = Rtf.ToHtml(rtfHack);
                Log("Moldura (Paciente/Assinatura) extraída.");
            }
            catch (Exception ex) { Log($"Erro na moldura: {ex.Message}"); }

            // ==============================================================================
            // FASE 3: A FUSÃO (FRANKENSTEIN)
            // ==============================================================================
            // Vamos inserir o Miolo (Fase 1) dentro da Moldura (Fase 2)
            if (!string.IsNullOrEmpty(htmlMoldura))
            {
                // Procura o fim da tabela do cabeçalho (</table>)
                int fimTabelaCabecalho = htmlMoldura.IndexOf("</table>");

                if (fimTabelaCabecalho > 0)
                {
                    // INSERÇÃO CIRÚRGICA: Coloca o texto logo após os dados do paciente
                    resultado.HtmlConteudo = htmlMoldura.Insert(fimTabelaCabecalho + 8,
                        $"<div class='miolo-do-laudo' style='margin: 25px 0; padding: 10px 0;'>{htmlMiolo}</div>");
                    Log("Fusão realizada: Texto inserido após o cabeçalho.");
                }
                else
                {
                    // Se não achar tabela, concatena (Header + Texto)
                    resultado.HtmlConteudo = htmlMoldura + "<hr/>" + htmlMiolo;
                }
            }
            else
            {
                resultado.HtmlConteudo = htmlMiolo;
            }

            // ==============================================================================
            // FASE 4: IMAGENS (OLE / BITMAP)
            // ==============================================================================
            Log("Extraindo imagens binárias...");
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
                                    Log($"Imagem recuperada ({type}).");
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex) { Log($"Erro imagens: {ex.Message}"); }

            // ==============================================================================
            // FASE 5: INJEÇÃO DAS IMAGENS
            // ==============================================================================
            if (resultado.ImagensBase64.Count > 0)
            {
                Log($"Aplicando {resultado.ImagensBase64.Count} imagens no layout...");
                resultado.HtmlConteudo = InjetarImagensNoHtml(resultado.HtmlConteudo, resultado.ImagensBase64);
            }

            return resultado;
        }

        // --- MÉTODOS AUXILIARES ---

        private static string LimparHtmlContainer(string html)
        {
            if (string.IsNullOrEmpty(html)) return "";
            // Remove tudo antes e depois do body para permitir aninhamento
            int bodyStart = html.IndexOf("<body>");
            if (bodyStart >= 0) html = html.Substring(bodyStart + 6);
            return html.Replace("</body>", "").Replace("</html>", "");
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
                        tagNova = tagNova.Replace("<img", "<img style='max-width: 180px; max-height: 120px; display:block;'");
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