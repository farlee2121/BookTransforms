using OpenXmlPowerTools;
using System;
using System.Collections.Generic;
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

            string outDocName = @"C:\Users\Spencer Farley\Desktop\moo.docx";
            List<string> docNameList = new List<string>() {
                @"C:\Users\Spencer Farley\Desktop\Roslyn Notes.docx",
                @"C:\Users\Spencer Farley\Desktop\Closing items.docx",
            };
            var sources = new List<Source>();
            //Document Streams (File Streams) of the documents to be merged.
            foreach (var docName in docNameList)
            {
                sources.Add(new Source(new WmlDocument(docName), true));
            }

            var mergedDoc = DocumentBuilder.BuildDocument(sources);
            mergedDoc.SaveAs(outDocName);

        }
    }
}
