using System;
using System.IO;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;

namespace ppt2pdf
{
    class Program
    {
        /// <summary>
        /// Converts Excel documents to PDFs
        /// </summary>
        /// <param name="rootPath">The path of the directory that contains all Excel documents to be converted</param>
        static void excel2pdf(string rootPath)
        {
            var objExcel = new Excel.Application(){
                DisplayAlerts = false
            };
            convert2pdf(rootPath, "*.xls?", (inputPath, outputPath) =>
            {
                var excelBook = objExcel.Workbooks.Open(inputPath, ReadOnly: MsoTriState.msoTrue);
                excelBook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, outputPath);
                excelBook.Close(SaveChanges:false);
                return true;
            });
            objExcel.Quit();
        }

        /// <summary>
        /// Converts Word documents to PDFs
        /// </summary>
        /// <param name="rootPath">The path of the directory that contains all Word documents to be converted</param>
        static void doc2pdf(string rootPath)
        {
            var objWord = new Word.Application(){
                DisplayAlerts = Word.WdAlertLevel.wdAlertsNone
            };
            objWord.WindowState = Word.WdWindowState.wdWindowStateMinimize;
            convert2pdf(rootPath, "*.doc?", (inputPath, outputPath) =>
            {
                var wordDoc = objWord.Documents.Open(inputPath, ReadOnly: MsoTriState.msoTrue);
                wordDoc.ExportAsFixedFormat(outputPath, Word.WdExportFormat.wdExportFormatPDF);
                wordDoc.Close(SaveChanges:false);
                return true;
            });
            objWord.Quit();
        }

        /// <summary>
        /// Applies the function <param name="conv"/> for each file found under rootPath
        /// </summary>
        /// <param name="rootPath">The path of the directory that contains all office files</param>
        /// <param name="pat">The pattern of the names of the files to be processed</param>
        /// <param name="conv">The function that converts each office file to a pdf document</param>
        static void convert2pdf(string rootPath, string pat, Func<string, string, bool> conv)
        {
            foreach (var file in Directory.GetFiles(rootPath, pat, SearchOption.AllDirectories))
            {
                // Skipping temporary files and the files under "reference" (sub-)directories
                if (
                    file.Contains("reference")
                    || ((File.GetAttributes(file) & FileAttributes.Temporary) == FileAttributes.Temporary)
                )
                    continue;
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

        /// <summary>
        /// Converts PowerPoint documents to PDFs
        /// </summary>
        /// <param name="rootPath">The path of the directory that contains all PowerPoint documents to be converted</param>
        static void ppt2pdf(string rootPath)
        {
            var objPowerPoint = new PowerPoint.Application(){
                DisplayAlerts = PowerPoint.PpAlertLevel.ppAlertsNone
            };
            objPowerPoint.WindowState = PowerPoint.PpWindowState.ppWindowMinimized;
            convert2pdf(rootPath, "*.ppt?", (inputPath, outputPath) =>
            {
                var pptDoc = objPowerPoint.Presentations.Open(inputPath, ReadOnly: MsoTriState.msoTrue);
                pptDoc.ExportAsFixedFormat2(outputPath, PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF, PowerPoint.PpFixedFormatIntent.ppFixedFormatIntentPrint, OutputType: PowerPoint.PpPrintOutputType.ppPrintOutputNotesPages, PrintHiddenSlides: MsoTriState.msoTrue);
                //Using printers to generate pdfs
                /*
                pptDoc.PrintOptions.ActivePrinter = "Microsoft Print to PDF";
                pptDoc.PrintOptions.OutputType = PowerPoint.PpPrintOutputType.ppPrintOutputNotesPages;
                pptDoc.PrintOptions.PrintHiddenSlides = MsoTriState.msoTrue;
                pptDoc.PrintOptions.HighQuality = MsoTriState.msoTrue;
                pptDoc.PrintOptions.FitToPage = MsoTriState.msoTrue;
                pptDoc.PrintOut(PrintToFile: outputPath);
                */
                pptDoc.Close();
                return true;
            });
            objPowerPoint.Quit();
        }

        static void Main(string[] args)
        {
            string rootPath = args[0];
            ppt2pdf(rootPath);
            doc2pdf(rootPath);
            excel2pdf(rootPath);
        }
    }
}
