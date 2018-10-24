using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    class Program
    {
        static void Main(string[] args)
        {
            DirectoryInfo di = new DirectoryInfo(@"C:\Users\k.komarov\source\ad\Open-Xml-PowerTools\Open-Xml-PowerTools\OpenXmlPowerToolsExamples\DocumentAssembler");
            args = "TemplateDocument.docx Data.xml AssembledDoc.docx".Split(' ');
            if (args.Length != 3)
            {
                PrintUsage();
                Environment.Exit(0);
            }

            FileInfo templateDoc = new FileInfo(Path.Combine(di.FullName, args[0]));
            if (!templateDoc.Exists)
            {
                Console.WriteLine("Error, {0} does not exist.", args[0]);
                PrintUsage();
                Environment.Exit(0);
            }
            FileInfo dataFile = new FileInfo(Path.Combine(di.FullName, args[1]));
            if (!dataFile.Exists)
            {
                Console.WriteLine("Error, {0} does not exist.", args[1]);
                PrintUsage();
                Environment.Exit(0);
            }
            FileInfo assembledDoc = new FileInfo(Path.Combine(di.FullName, args[2]));
            if (assembledDoc.Exists)
            {
                Console.WriteLine("Error, {0} exists.", args[2]);
                PrintUsage();
                Environment.Exit(0);
            }

            WmlDocument wmlDoc = new WmlDocument(templateDoc.FullName);
            XElement data = XElement.Load(dataFile.FullName);
            bool templateError;
            WmlDocument wmlAssembledDoc = DocumentAssembler.AssembleDocument(wmlDoc, data, out templateError);
            if (templateError)
            {
                Console.WriteLine("Errors in template.");
                Console.WriteLine("See {0} to determine the errors in the template.", assembledDoc.Name);
            }

            wmlAssembledDoc.SaveAs(assembledDoc.FullName);
        }

        static void PrintUsage()
        {
            Console.WriteLine("Usage: DocumentAssembler TemplateDocument.docx Data.xml AssembledDoc.docx");
        }
    }
}