using System;
using System.IO;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Core;

namespace ppt2pdf
{
    class Program
    {
        static void doc2pdf(string[] args)
        {
            var objWord = new Word.Application();
            objWord.WindowState = Word.WdWindowState.wdWindowStateMinimize;
            convert2pdf(args[0], "*.doc?", (inputPath, outputPath) =>
            {
                var wordDoc = objWord.Documents.Open(inputPath,ReadOnly: MsoTriState.msoTrue);
                wordDoc.ExportAsFixedFormat(outputPath, Word.WdExportFormat.wdExportFormatPDF);
                wordDoc.Close();
                return true;
            });
        }
        static void convert2pdf(string rootPath, string pat, Func<string, string, bool> conv)
        {
            foreach (var file in Directory.GetFiles(rootPath, pat, SearchOption.AllDirectories))
            {
                if (file.Contains("reference")) continue;
                var inputPath = Path.GetFullPath(file);
                var dir_name = Path.GetDirectoryName(inputPath);
                var file_name = Path.GetFileNameWithoutExtension(inputPath);
                var outputPath = Path.Combine(dir_name, file_name + ".pdf");
                // skipping documents that have been already converted
                if (File.Exists(outputPath) && File.GetLastWriteTime(outputPath) >= File.GetLastWriteTime(inputPath))
                {
                    Console.WriteLine($"skipping {inputPath}");
                    continue;
                }
                else
                {
                    Console.WriteLine($"Converting: {inputPath}");
                    conv(inputPath, outputPath);
                    Console.WriteLine($"Completed: {outputPath}.");
                }
            }
        }
        static void ppt2pdf(string[] args)
        {
            var objPowerPoint = new PowerPoint.Application();
            objPowerPoint.WindowState = PowerPoint.PpWindowState.ppWindowMinimized;
            convert2pdf(args[0], "*.ppt?", (inputPath, outputPath) =>
            {
                var pptDoc = objPowerPoint.Presentations.Open(inputPath, ReadOnly : MsoTriState.msoTrue);
                pptDoc.ExportAsFixedFormat(outputPath, PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF, OutputType: PowerPoint.PpPrintOutputType.ppPrintOutputNotesPages, PrintHiddenSlides : MsoTriState.msoTrue);
                pptDoc.Close();
                return true;
            });
        }
        static void Main(string[] args)
        {
            ppt2pdf(args);
            doc2pdf(args);
        }
    }
}
