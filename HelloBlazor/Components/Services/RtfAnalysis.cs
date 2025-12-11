using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Globalization;

namespace HelloBlazor.Components.Services
{
    // === CLASSES AUXILIARES DE LEITURA RTF (SIMON MOURIER) ===

    public class RtfObject
    {
        public string Text { get; private set; }
        public RtfObject(string text) { Text = text?.Trim(); }
    }

    public class RtfText : RtfObject { public RtfText(string text) : base(text) { } }
    public class RtfControlWord : RtfObject { public RtfControlWord(string name) : base(name) { } }

    public class RtfReader
    {
        private TextReader _reader;
        public RtfReader(TextReader reader) { _reader = reader; }

        public IEnumerable<RtfObject> Read()
        {
            StringBuilder controlWord = new StringBuilder();
            StringBuilder text = new StringBuilder();
            Stack<bool> stack = new Stack<bool>(); // Stack simples para grupos
            bool inControlWord = false;

            int i;
            while ((i = _reader.Read()) != -1)
            {
                char c = (char)i;

                // Ignora quebras de linha e nulos do arquivo bruto
                if (c == '\r' || c == '\n' || c == '\0') continue;

                if (inControlWord)
                {
                    if (char.IsLetterOrDigit(c) || c == '-')
                    {
                        controlWord.Append(c);
                    }
                    else
                    {
                        inControlWord = false;
                        if (controlWord.Length > 0) yield return new RtfControlWord(controlWord.ToString());
                        controlWord.Clear();

                        if (c == ' ') continue; // Espaço após control word é consumido
                        // Se não for espaço, processa como caractere normal abaixo (exceto se for inicio de grupo)
                    }
                }

                if (!inControlWord)
                {
                    if (c == '{')
                    {
                        if (text.Length > 0) { yield return new RtfText(text.ToString()); text.Clear(); }
                        stack.Push(true);
                    }
                    else if (c == '}')
                    {
                        if (text.Length > 0) { yield return new RtfText(text.ToString()); text.Clear(); }
                        if (stack.Count > 0) stack.Pop();
                    }
                    else if (c == '\\')
                    {
                        if (text.Length > 0) { yield return new RtfText(text.ToString()); text.Clear(); }
                        inControlWord = true;
                    }
                    else
                    {
                        text.Append(c);
                    }
                }
            }
        }

        // Utilitários de navegação do Simon Mourier
        public static bool MoveToNextControlWord(IEnumerator<RtfObject> enumerator, string word)
        {
            while (enumerator.MoveNext())
            {
                if (enumerator.Current is RtfControlWord cw && cw.Text == word) return true;
            }
            return false;
        }

        public static string GetNextText(IEnumerator<RtfObject> enumerator)
        {
            while (enumerator.MoveNext())
            {
                if (enumerator.Current is RtfText t) return t.Text;
            }
            return null;
        }

        public static byte[] GetNextTextAsByteArray(IEnumerator<RtfObject> enumerator, List<string> logs)
        {
            while (enumerator.MoveNext())
            {
                if (enumerator.Current is RtfText t)
                {
                    return HexToBytes(t.Text, logs);
                }
            }
            return null;
        }

        private static byte[] HexToBytes(string hex, List<string> logs)
        {
            if (string.IsNullOrEmpty(hex)) return null;
            var bytes = new List<byte>();
            for (int i = 0; i < hex.Length; i++)
            {
                char c = hex[i];
                if (char.IsWhiteSpace(c)) continue;

                // Pega par de caracteres
                if (i + 1 < hex.Length)
                {
                    string byteStr = hex.Substring(i, 2);
                    if (byte.TryParse(byteStr, NumberStyles.HexNumber, null, out byte b))
                    {
                        bytes.Add(b);
                    }
                    i++; // Pula o próximo pois já usamos
                }
            }
            return bytes.ToArray();
        }
    }

    // === CLASSE PARA LER O PACOTE OLE (SIMON MOURIER) ===
    public class PackagedObject
    {
        public string DisplayName { get; private set; }
        public string FilePath { get; private set; }
        public byte[] Data { get; private set; }

        // Remove o cabeçalho "Object Header" OLE 1.0
        public static void ExtractObjectData(Stream inputStream, Stream outputStream)
        {
            BinaryReader reader = new BinaryReader(inputStream);
            reader.ReadInt32(); // OLEVersion
            int formatId = reader.ReadInt32(); // FormatID

            if (formatId != 2) return; // 2 = EmbeddedObject

            // Lê strings prefixadas com tamanho
            ReadLengthPrefixedAnsiString(reader); // ClassName (ex: Package)
            ReadLengthPrefixedAnsiString(reader); // TopicName
            ReadLengthPrefixedAnsiString(reader); // ItemName

            int nativeDataSize = reader.ReadInt32();
            byte[] bytes = reader.ReadBytes(nativeDataSize);
            outputStream.Write(bytes, 0, bytes.Length);
        }

        private static string ReadLengthPrefixedAnsiString(BinaryReader reader)
        {
            if (reader.BaseStream.Position >= reader.BaseStream.Length) return "";
            int length = reader.ReadInt32();
            if (length <= 0) return "";
            byte[] bytes = reader.ReadBytes(length);
            // Retorna string removendo o null terminator se existir
            return Encoding.ASCII.GetString(bytes, 0, bytes.Length > 0 ? bytes.Length - 1 : 0);
        }

        // Lê o pacote final (Package) para extrair o arquivo real
        public static PackagedObject Extract(Stream inputStream)
        {
            inputStream.Position = 0;
            BinaryReader reader = new BinaryReader(inputStream);

            // Assinatura
            reader.ReadUInt16();

            PackagedObject po = new PackagedObject();
            po.DisplayName = ReadAnsiString(reader);
            ReadAnsiString(reader); // Icon path
            reader.ReadUInt16();    // Icon index

            int type = reader.ReadUInt16();
            // Tipo 3 = Arquivo Embutido (File)
            if (type != 3) throw new Exception("O pacote OLE não é um arquivo embutido (Tipo encontrado: " + type + ")");

            reader.ReadInt32(); // Next size
            po.FilePath = ReadAnsiString(reader);

            int dataSize = reader.ReadInt32();
            po.Data = reader.ReadBytes(dataSize);

            return po;
        }

        private static string ReadAnsiString(BinaryReader reader)
        {
            StringBuilder sb = new StringBuilder();
            while (reader.BaseStream.Position < reader.BaseStream.Length)
            {
                byte b = reader.ReadByte();
                if (b == 0) return sb.ToString();
                sb.Append((char)b);
            }
            return sb.ToString();
        }
    }
}