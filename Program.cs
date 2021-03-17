using System;
using System.IO;
using Xceed.Words.NET;

namespace ExtractHyperlinksFromWord
{
    class Program
    {
        static void Main()
        {
            const string extension = "*.docx";

            var files = Directory.GetFiles(".", extension);

            foreach (var fileName in files)
            {
                using var doc = DocX.Load($"{fileName}");

                var fileNameWithoutExtension = fileName[2..^(extension.Length-1)];
                var finalFilePath = $".//{fileNameWithoutExtension}.txt";

                if(File.Exists(finalFilePath)) File.Delete(finalFilePath);

                using var sw = File.CreateText(finalFilePath);

                foreach (var hyperlink in doc.Hyperlinks)
                {
                    sw.WriteLine(hyperlink.Uri.AbsoluteUri);
                }
            }

            Console.WriteLine("All set");
        }
    }
}
