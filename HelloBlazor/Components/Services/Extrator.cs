using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions; // <--- NECESSÁRIO PARA O TRANSPLANTE
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

            // ==============================================================================
            // PASSO 1: LIMPEZA E GERAÇÃO DO HTML (ESQUELETO)
            // ==============================================================================
            string rtfParaTexto = rtfConteudo;

            // Corrige o bug do cabeçalho oculto (para aparecer nome do paciente)
            if (rtfParaTexto.Contains(@"\header"))
            {
                rtfParaTexto = rtfParaTexto.Replace(@"\header", @"\pard");
            }

            try
            {
                // Gera o HTML com os placeholders quebrados
                resultado.HtmlConteudo = Rtf.ToHtml(rtfParaTexto);
                Log("HTML gerado (com placeholders).");
            }
            catch (Exception ex)
            {
                resultado.HtmlConteudo = $"<p>Erro no layout: {ex.Message}</p>";
            }

            // ==============================================================================
            // PASSO 2: EXTRAÇÃO DAS IMAGENS REAIS (ÓRGÃOS)
            // ==============================================================================
            try
            {
                using (StringReader sr = new StringReader(rtfConteudo))
                {
                    RtfReader reader = new RtfReader(sr);
                    var enumerator = reader.Read().GetEnumerator();

                    while (enumerator.MoveNext())
                    {
                        // Procura \object (OLE)
                        if (enumerator.Current is RtfControlWord cw && cw.Text == "object")
                        {
                            if (RtfReader.MoveToNextControlWord(enumerator, "objclass"))
                            {
                                if (RtfReader.GetNextText(enumerator) == "Package")
                                {
                                    if (RtfReader.MoveToNextControlWord(enumerator, "objdata"))
                                    {
                                        byte[] oleData = RtfReader.GetNextTextAsByteArray(enumerator, resultado.Logs);
                                        ProcessarOlePackage(oleData, resultado, Log);
                                    }
                                }
                            }
                        }
                        // Procura \pict (NATIVO/QRCODE/ASSINATURA)
                        else if (enumerator.Current is RtfControlWord cwPict && cwPict.Text == "pict")
                        {
                            bool isDib = false;
                            string type = "jpeg"; // Chute padrão

                            // Descobre o tipo
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
                                    if (isDib) imgData = ConvertDibToBmp(imgData); // Fix do Bitmap
                                    string base64 = Convert.ToBase64String(imgData);

                                    // Adiciona prefixo para HTML direto
                                    string prefixo = type == "png" ? "data:image/png;base64," : "data:image/bmp;base64,";
                                    resultado.ImagensBase64.Add(prefixo + base64);

                                    Log($"Imagem ({type}) recuperada.");
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex) { Log($"Erro extrator img: {ex.Message}"); }

            // ==============================================================================
            // PASSO 3: O TRANSPLANTE (INJETAR IMAGENS NO HTML)
            // ==============================================================================
            if (resultado.ImagensBase64.Count > 0)
            {
                Log($"Injetando {resultado.ImagensBase64.Count} imagens no documento...");
                resultado.HtmlConteudo = InjetarImagensNoHtml(resultado.HtmlConteudo, resultado.ImagensBase64);
            }

            return resultado;
        }

        // --- SUBSTITUI AS TAGS <IMG> QUEBRADAS PELAS NOSSAS IMAGENS ---
        private static string InjetarImagensNoHtml(string html, List<string> imagensBase64)
        {
            int indexImagem = 0;

            // Esta Regex procura qualquer tag <img ... src="..." ... >
            // O RtfPipe gera tags img quando encontra \pict, mas o src geralmente vem vazio ou quebrado para bitmaps
            return Regex.Replace(html, "<img[^>]+src=[\"']([^\"']*)[\"'][^>]*>", match =>
            {
                // Se ainda temos imagens na nossa lista extraída, usamos a próxima
                if (indexImagem < imagensBase64.Count)
                {
                    string tagOriginal = match.Value;
                    string srcAntigo = match.Groups[1].Value; // O que está dentro das aspas do src
                    string srcNovo = imagensBase64[indexImagem];

                    indexImagem++;

                    // Se a tag original não tinha style, vamos adicionar um tamanho padrão seguro
                    // para evitar que o QR Code fique gigante
                    string tagNova = tagOriginal.Replace(srcAntigo, srcNovo);

                    if (!tagNova.Contains("style"))
                    {
                        tagNova = tagNova.Replace("<img", "<img style='max-width: 150px; height: auto;'");
                    }

                    return tagNova;
                }

                // Se acabaram as imagens, retorna a tag original (placeholder)
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
                        // Assume JPEG para OLE
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