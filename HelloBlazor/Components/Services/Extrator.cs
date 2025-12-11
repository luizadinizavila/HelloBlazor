using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using RtfPipe; // <--- ADICIONE ISSO (O pacote que você instalou)

namespace HelloBlazor.Components.Services
{
    public class ResultadoExtracao
    {
        public List<string> ImagensBase64 { get; set; } = new List<string>();
        public List<string> Logs { get; set; } = new List<string>();
        public string HtmlConteudo { get; set; } = ""; // <--- NOVO: Para guardar o HTML
    }

    public class Extrator
    {
        public static ResultadoExtracao ProcessarRtf(string rtfConteudo)
        {
            var resultado = new ResultadoExtracao();
            void Log(string msg) => resultado.Logs.Add($"[{DateTime.Now:HH:mm:ss}] {msg}");

            if (string.IsNullOrWhiteSpace(rtfConteudo))
            {
                Log("ERRO: Conteúdo RTF vazio.");
                return resultado;
            }

            // 1. TENTA CONVERTER O TEXTO PARA HTML (Para leitura humana)
            Log("Convertendo estrutura de texto RTF para HTML...");
            try
            {
                // O RtfPipe faz a mágica de transformar tabelas e texto em HTML
                resultado.HtmlConteudo = Rtf.ToHtml(rtfConteudo);
                Log("Conversão HTML realizada com sucesso.");
            }
            catch (Exception ex)
            {
                Log($"Aviso: Não foi possível converter o layout para HTML: {ex.Message}");
                resultado.HtmlConteudo = "<p class='text-danger'>Erro ao renderizar layout.</p>";
            }

            // 2. EXTRAÇÃO AVANÇADA DE IMAGENS (Sua lógica que já funciona)
            Log("Iniciando extração HÍBRIDA de imagens (OLE + PICT)...");

            try
            {
                using (StringReader sr = new StringReader(rtfConteudo))
                {
                    RtfReader reader = new RtfReader(sr);
                    var enumerator = reader.Read().GetEnumerator();

                    while (enumerator.MoveNext())
                    {
                        // === ESTRATÉGIA 1: OLE OBJECT (SIMON MOURIER) ===
                        if (enumerator.Current is RtfControlWord cw && cw.Text == "object")
                        {
                            if (RtfReader.MoveToNextControlWord(enumerator, "objclass"))
                            {
                                string className = RtfReader.GetNextText(enumerator);
                                if (className == "Package")
                                {
                                    if (RtfReader.MoveToNextControlWord(enumerator, "objdata"))
                                    {
                                        byte[] oleData = RtfReader.GetNextTextAsByteArray(enumerator, resultado.Logs);
                                        ProcessarOlePackage(oleData, resultado, Log);
                                    }
                                }
                            }
                        }
                        // === ESTRATÉGIA 2: PICTURE TAG (NATIVO) ===
                        else if (enumerator.Current is RtfControlWord cwPict && cwPict.Text == "pict")
                        {
                            // Tenta identificar o tipo
                            string imageType = "unknown";
                            bool isDib = false;

                            while (enumerator.MoveNext())
                            {
                                if (enumerator.Current is RtfControlWord typeCw)
                                {
                                    if (typeCw.Text.StartsWith("png")) { imageType = "png"; break; }
                                    if (typeCw.Text.StartsWith("jpeg")) { imageType = "jpeg"; break; }
                                    if (typeCw.Text.StartsWith("dibitmap")) { imageType = "bmp"; isDib = true; break; }
                                    if (typeCw.Text.StartsWith("wmetafile")) { imageType = "wmf"; break; }
                                }
                                if (enumerator.Current is RtfText) break;
                            }

                            if (imageType != "wmf") // Ignora WMF pois navegadores não abrem
                            {
                                byte[] imgData = RtfReader.GetNextTextAsByteArray(enumerator, resultado.Logs);

                                if (imgData != null && imgData.Length > 0)
                                {
                                    if (isDib) imgData = ConvertDibToBmp(imgData); // Seu fix mágico

                                    string base64 = Convert.ToBase64String(imgData);
                                    resultado.ImagensBase64.Add(base64);
                                    Log($"SUCESSO! Imagem nativa ({imageType}) extraída.");
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"ERRO FATAL NA EXTRAÇÃO: {ex.Message}");
            }

            return resultado;
        }

        // --- MÉTODOS AUXILIARES (IGUAIS AO ANTERIOR) ---
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
                        Log($"[OLE] Arquivo extraído: {arquivoFinal.DisplayName}");
                        resultado.ImagensBase64.Add(Convert.ToBase64String(arquivoFinal.Data));
                    }
                }
            }
            catch (Exception) { /* Ignora erros de OLE silenciosamente para não poluir log */ }
        }

        private static byte[] ConvertDibToBmp(byte[] dibData)
        {
            int headerSize = 14;
            int fileSize = headerSize + dibData.Length;
            byte[] bmp = new byte[fileSize];

            bmp[0] = 0x42; bmp[1] = 0x4D; // BM
            BitConverter.GetBytes(fileSize).CopyTo(bmp, 2);
            bmp[10] = 54; // Offset padrão seguro

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