using System;
using System.IO;
using BCL.easyPDF.PDFProcessor;
using BCL.easyPDF.Printer;

namespace DotNetEasyPdfSample.ConsoleApp.DotNetFw
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // 形式変換のサンプル
            var printer = new Printer();
            var bytes = File.ReadAllBytes(@".\Sample.docx");

            var converted = printer.WordPrintJobEx.PrintOut3(bytes, ".docx");
            // NOTE 汎用の変換メソッドを利用するとOfficeオブジェクトが残存してしまい、連続実行した場合に例外が発生する
            // COM版、ネイティブ版、どちらのオブジェクトを利用しても挙動は変わらず・・・
            // var converted = printer.PrintJob.PrintOut3(bytes, ".docx");

            File.WriteAllBytes($@".\Sample{DateTime.Now.Ticks}.pdf", converted);

            printer.Dispose();

            // マージのサンプル
            var pdfBytes = File.ReadAllBytes(@".\Sample-1.pdf");
            var pdfBBytes2 = File.ReadAllBytes(@".\Sample-2.pdf");
            var processor = new PDFProcessor();
            var merged = processor.MergeMem(pdfBytes, pdfBBytes2);

            File.WriteAllBytes($@".\Sample{DateTime.Now.Ticks}.pdf", merged);
        }
    }
}
