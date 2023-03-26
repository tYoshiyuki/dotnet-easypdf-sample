using System.Net;
using BCL.easyPDF.PDFProcessor;
using BCL.easyPDF.Printer;
using DotNetEasyPdfSample.Web.Models;
using Microsoft.AspNetCore.Mvc;

namespace DotNetEasyPdfSample.Web.Controllers
{
    /// <summary>
    /// PDFコントローラ
    /// </summary>
    [Route("api/[controller]")]
    [ApiController]
    public class PdfController : ControllerBase
    {
        /// <summary>
        /// PDFファイルへ変換します。
        /// </summary>
        /// <param name="request"></param>
        /// <returns></returns>
        [HttpPost("ConvertToPdf")]
        [Produces("application/octet-stream", Type = typeof(FileResult))]
        public async Task<IActionResult> ConvertToPdf([FromForm]ConvertToPdfRequest request)
        {
            try
            {
                var file = request.File;
                var extension = Path.GetExtension(file.FileName);

                // NOTE .pdfは変換失敗するので、バリデーションにてフィルタリングする
                // その他、変換不可能な形式は必要に応じてフィルタリングする
                if (extension.Contains(".pdf", StringComparison.InvariantCultureIgnoreCase))
                {
                    return Problem($"Not supported file type:{extension}.", statusCode: (int) HttpStatusCode.BadRequest);
                }

                using var printer = new Printer();

                using var memoryStream = new MemoryStream();
                await file.CopyToAsync(memoryStream);

                byte[] result;

                // NOTE 汎用の変換メソッドを利用するとOfficeオブジェクトが残存してしまい、連続実行した場合に例外が発生する
                // COM版、ネイティブ版、どちらのオブジェクトを利用しても挙動は変わらず・・・
                // 従って拡張子に応じて個別メソッドを呼び分けた方が良い
                if (extension.Contains(".doc", StringComparison.OrdinalIgnoreCase))
                {
                    result = printer.WordPrintJobEx.PrintOut3(memoryStream.ToArray(), Path.GetExtension(file.FileName));
                }
                else if (extension.Contains(".xls", StringComparison.OrdinalIgnoreCase))
                {
                    result = printer.ExcelPrintJobEx.PrintOut3(memoryStream.ToArray(), Path.GetExtension(file.FileName));
                }
                else
                {
                    result = printer.PrintJob.PrintOut3(memoryStream.ToArray(), Path.GetExtension(file.FileName));
                }

                return File(result, "application/octet-stream", fileDownloadName: Path.GetFileNameWithoutExtension(file.FileName) + DateTime.Now.Ticks + ".pdf");
            }
            catch (PrinterException ex)
            {
                Console.WriteLine(ex.Message + "ErrorCode: " + ex.ErrorCode);
                throw;
            }
        }

        /// <summary>
        /// PDFファイルをマージします。
        /// </summary>
        /// <param name="request"></param>
        /// <returns></returns>
        [HttpPost("MergePdf")]
        [Produces("application/octet-stream", Type = typeof(FileResult))]
        public async Task<IActionResult> MergePdf([FromForm]MergePdfRequest request)
        {
            try
            {
                var file1 = request.File1;
                var file2 = request.File2;

                var extension1 = Path.GetExtension(file1.FileName);
                var extension2 = Path.GetExtension(file2.FileName);

                if (!extension1.Contains(".pdf", StringComparison.InvariantCultureIgnoreCase) || !extension2.Contains(".pdf", StringComparison.InvariantCultureIgnoreCase))
                {
                    return Problem("File type must be pdf.", statusCode: (int)HttpStatusCode.BadRequest);
                }

                var processor = new PDFProcessor();

                using var memoryStream1 = new MemoryStream();
                await file1.CopyToAsync(memoryStream1);

                using var memoryStream2 = new MemoryStream();
                await file2.CopyToAsync(memoryStream2);

                var result = processor.MergeMem(memoryStream1.ToArray(), memoryStream2.ToArray());

                return File(result, "application/octet-stream", fileDownloadName: "Merged" + DateTime.Now.Ticks + ".pdf");
            }
            catch (PrinterException ex)
            {
                Console.WriteLine(ex.Message + "ErrorCode: " + ex.ErrorCode);
                throw;
            }
        }
    }
}
