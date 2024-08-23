using System;
using System.Collections.Generic;
using System.IO;
using System.Security;
using System.Text.RegularExpressions;

using DocumentFormat.OpenXml.Packaging;

using namasdev.Core.IO;

namespace namasdev.Word
{
    public class WordHelper
    {
        private const string WORD_NEW_LINE = "<w:br/>";

        public static Archivo GenerarWordDesdeTemplateReemplazandoCampos(
            Archivo template, Dictionary<string,string> campos)
        {
            using (var stream = new MemoryStream())
            {
                // NOTA (ML): necesitamos inicializar/cargar el MemoryStream de esta manera para que sea resizable
                stream.Write(template.Contenido, 0, template.Contenido.Length);

                using (var doc = WordprocessingDocument.Open(stream, isEditable: true, new OpenSettings { AutoSave = false }))
                {
                    string docText = null;
                    using (var sr = new StreamReader(doc.MainDocumentPart.GetStream()))
                    {
                        docText = sr.ReadToEnd();
                    }

                    docText = Regex.Replace(
                        docText, 
                        @"{{(\w+)}}", 
                        (m) => 
                        {
                            string valor;
                            return campos.TryGetValue(m.Groups[1].Value, out valor)
                                ? FormatearValor(valor)
                                : String.Empty;
                        });

                    using (var wr = new StreamWriter(doc.MainDocumentPart.GetStream(FileMode.Create)))
                    {
                        wr.Write(docText);
                    }
                    
                    template.Contenido = stream.ToArray();

                    return template;
                }
            }
        }

        private static string FormatearValor(string valor)
        {
            return SecurityElement.Escape(valor)
                .Replace(Environment.NewLine, WORD_NEW_LINE);
        }
    }
}
