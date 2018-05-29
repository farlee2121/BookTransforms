using Mono.Options;
using OpenXmlPowerTools;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Clients.Console
{
    class Program
    {
        static void Main(string[] args)
        {
            // https://stackoverflow.com/questions/18351829/merge-multiple-word-documents-into-one-open-xml/18352219#18352219
            // OpenXMLPowerTools - (compose, template, convert to html, split word docs, convert html to docx, search&replace, and more)

            //https://stackoverflow.com/questions/491595/best-way-to-parse-command-line-arguments-in-c

            string outDocName = @".\out.docx"; // default value
            string formatDocName = @"";

            var p = new OptionSet() {
                { "o|OutputFile=", "the name of the generated docx.",
                    v => outDocName = v.Trim() },
                { "f|TemplateFile=", "outline file",
                    v => {
                        formatDocName = v.Trim();
                    } }
            };

            IEnumerable<string> meow = p.Parse(args);
            IEnumerable<string> docNameList = GetOutlineComponents(formatDocName);
            var sources = new List<Source>();
            //Document Streams (File Streams) of the documents to be merged.
            foreach (var docName in docNameList)
            {
                sources.Add(new Source(new WmlDocument(docName), true));
            }

            var mergedDoc = DocumentBuilder.BuildDocument(sources);
            mergedDoc.SaveAs(outDocName);

        }

        static IEnumerable<string> GetOutlineComponents(string filePath) // should this take a stream?
        {
            IEnumerable<string> lines = File.ReadAllLines(filePath);

            return lines;
        }
    }
}
