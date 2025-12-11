using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace HelloBlazor.Components.Services
{
    public class ResultadoExtracao
    {
        public List<string> ImagensBase64 { get; set; } = new List<string>();
        public List<string> Logs { get; set; } = new List<string>();
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

            Log("Iniciando análise HÍBRIDA (OLE + PICT)...");

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
                            Log("--> [OLE] Objeto encontrado. Analisando...");
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
                            Log("--> [PICT] Imagem nativa encontrada.");

                            // Tenta identificar o tipo de imagem pelas próximas control words
                            string imageType = "unknown";
                            bool isDib = false;

                            // Avança para descobrir o tipo (wmetafile, dibitmap, pngblip, jpegblip)
                            // Loop simples para pegar a próxima control word relevante
                            while (enumerator.MoveNext())
                            {
                                if (enumerator.Current is RtfControlWord typeCw)
                                {
                                    if (typeCw.Text.StartsWith("png")) { imageType = "png"; break; }
                                    if (typeCw.Text.StartsWith("jpeg")) { imageType = "jpeg"; break; }
                                    if (typeCw.Text.StartsWith("dibitmap")) { imageType = "bmp"; isDib = true; break; }
                                    if (typeCw.Text.StartsWith("wmetafile")) { imageType = "wmf"; break; }
                                }
                                // Se achou texto (hex), para o loop de tipo, pois já são os dados
                                if (enumerator.Current is RtfText) break;
                            }

                            Log($"    Formato detectado: {imageType}");

                            if (imageType == "wmf")
                            {
                                Log("    AVISO: Formato WMF não é suportado por navegadores modernos.");
                            }
                            else
                            {
                                // Extrai o Hex dos dados
                                byte[] imgData = RtfReader.GetNextTextAsByteArray(enumerator, resultado.Logs);

                                if (imgData != null && imgData.Length > 0)
                                {
                                    if (isDib)
                                    {
                                        // Truque: Converter DIB cru em BMP válido (adicionando cabeçalho)
                                        Log("    Aplicando correção de cabeçalho BMP para DIB...");
                                        imgData = ConvertDibToBmp(imgData);
                                    }

                                    string base64 = Convert.ToBase64String(imgData);

                                    // Adiciona prefixo correto para o navegador entender
                                    string mimeType = imageType == "png" ? "image/png" : "image/jpeg";
                                    // Se for BMP corrigido, manda como png ou bmp
                                    if (isDib) mimeType = "image/bmp";

                                    // Nota: Para exibir no <img> do HTML, não precisamos do prefixo aqui 
                                    // pois o Home.razor já coloca "data:image/jpeg...".
                                    // Mas se for PNG, o prefixo fixo no razor pode atrapalhar. 
                                    // O ideal seria retornar o objeto completo, mas vamos simplificar:
                                    // O navegador costuma ser esperto com Base64.

                                    resultado.ImagensBase64.Add(base64);
                                    Log($"    SUCESSO! Imagem nativa extraída ({imgData.Length} bytes).");
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"ERRO FATAL: {ex.Message}");
            }

            Log($"Fim. Total: {resultado.ImagensBase64.Count} imagens.");
            return resultado;
        }

        // Lógica separada para OLE (Simon Mourier)
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
                        Log($"    [OLE] Arquivo dentro do pacote: {arquivoFinal.DisplayName}");
                        resultado.ImagensBase64.Add(Convert.ToBase64String(arquivoFinal.Data));
                    }
                }
            }
            catch (Exception ex) { Log($"    Erro ao abrir pacote OLE: {ex.Message}"); }
        }

        // Método auxiliar para consertar DIB (Device Independent Bitmap) para BMP
        private static byte[] ConvertDibToBmp(byte[] dibData)
        {
            // Um BMP arquivo é = [BITMAPFILEHEADER] + [DIB Data]
            // DIB Data começa com o tamanho do header (geralmente 40 bytes - 0x28)

            // BITMAPFILEHEADER (14 bytes):
            // 0-1: 'BM'
            // 2-5: Tamanho total do arquivo
            // 6-9: Reservado (0)
            // 10-13: Offset onde começam os pixels

            int headerSize = 14;
            int fileSize = headerSize + dibData.Length;
            byte[] bmp = new byte[fileSize];

            // Assinatura 'BM'
            bmp[0] = 0x42;
            bmp[1] = 0x4D;

            // Tamanho do arquivo
            BitConverter.GetBytes(fileSize).CopyTo(bmp, 2);

            // Reservado
            bmp[6] = 0;
            bmp[7] = 0;
            bmp[8] = 0;
            bmp[9] = 0;

            // Offset dos pixels
            // Precisamos ler o DIB Header para saber cores, etc.
            // Geralmente DIB Header tem 40 bytes.
            // Se for 32bpp, offset é 14 + 40 = 54.
            // Vamos ler o tamanho do DIB header do próprio dado (primeiros 4 bytes)
            if (dibData.Length > 4)
            {
                int dibHeaderSize = BitConverter.ToInt32(dibData, 0);

                // Calculando offset grosseiro (assumindo sem paleta de cores para 16/24/32 bits)
                // Se precisar ser exato, teríamos que ler bitcount, clrused, etc.
                // Mas para a maioria dos casos simples de RTF (32bpp/24bpp), 14 + dibHeaderSize funciona.
                int pixelOffset = 14 + dibHeaderSize;
                BitConverter.GetBytes(pixelOffset).CopyTo(bmp, 10);
            }
            else
            {
                // Fallback
                BitConverter.GetBytes(54).CopyTo(bmp, 10);
            }

            // Copia o DIB original
            Array.Copy(dibData, 0, bmp, 14, dibData.Length);

            return bmp;
        }
    }
}