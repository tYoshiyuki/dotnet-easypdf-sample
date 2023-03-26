using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web.Http;
using BCL.easyPDF.PDFProcessor;
using BCL.easyPDF.Printer;

namespace DotNetEasyPdfSample.Web.DotNetFw.Controllers
{
    /// <summary>
    /// PDFコントローラ
    /// </summary>
    [RoutePrefix("api/Pdf")]
    public class PdfController : ApiController
    {
        /// <summary>
        /// PDFファイルへ変換します。
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        [Route("ConvertToPdf")]
        public HttpResponseMessage ConvertToPdf()
        {
            try
            {
                if (!Request.Content.IsMimeMultipartContent())
                {
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Must be multipart content.");
                }

                var provider = Request.Content.ReadAsMultipartAsync().Result;

                var fileContent = provider.Contents.FirstOrDefault();
                var fileName = fileContent?.Headers.ContentDisposition.FileName.TrimStart('\"').TrimEnd('\"');

                if (string.IsNullOrEmpty(fileName))
                {
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "File is empty.");
                }

                var extension = Path.GetExtension(fileName);

                // NOTE .pdfは変換失敗するので、バリデーションにてフィルタリングする
                // その他、変換不可能な形式は必要に応じてフィルタリングする
                if (extension.Contains(".pdf"))
                {
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, $"Not supported file type:{extension}.");
                }

                using (var printer = new Printer())
                {
                    var bytes = fileContent.ReadAsByteArrayAsync().Result;

                    byte[] result;

                    // NOTE 汎用の変換メソッドを利用するとOfficeオブジェクトが残存してしまい、連続実行した場合に例外が発生する
                    // COM版、ネイティブ版、どちらのオブジェクトを利用しても挙動は変わらず・・・
                    // 従って拡張子に応じて個別メソッドを呼び分けた方が良い
                    if (extension.Contains(".doc"))
                    {
                        result = printer.WordPrintJobEx.PrintOut3(bytes, Path.GetExtension(fileName));
                    }
                    else if (extension.Contains(".xls"))
                    {
                        result = printer.ExcelPrintJobEx.PrintOut3(bytes, Path.GetExtension(fileName));
                    }
                    else
                    {
                        result = printer.PrintJob.PrintOut3(bytes, Path.GetExtension(fileName));
                    }

                    var response = Request.CreateResponse(HttpStatusCode.OK);
                    response.Content = new ByteArrayContent(result);
                    response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
                    response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
                    {
                        FileName = Path.GetFileNameWithoutExtension(fileName) + DateTime.Now.Ticks + ".pdf"
                    };

                    return response;
                }
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
        /// <returns></returns>
        [HttpPost]
        [Route("MergePdf")]
        public HttpResponseMessage MergePdf()
        {
            try
            {
                if (!Request.Content.IsMimeMultipartContent())
                {
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Must be multipart content.");
                }

                var provider = Request.Content.ReadAsMultipartAsync().Result;

                var fileContent1 = provider.Contents.FirstOrDefault();
                var fileName1 = fileContent1?.Headers.ContentDisposition.FileName.TrimStart('\"').TrimEnd('\"');

                var fileContent2 = provider.Contents.Skip(1).FirstOrDefault();
                var fileName2 = fileContent2?.Headers.ContentDisposition.FileName.TrimStart('\"').TrimEnd('\"');

                if (string.IsNullOrEmpty(fileName1) || string.IsNullOrEmpty(fileName2))
                {
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "File is empty.");
                }

                var extension1 = Path.GetExtension(fileName1);
                var extension2 = Path.GetExtension(fileName2);

                if (!extension1.Contains(".pdf") || !extension2.Contains(".pdf"))
                {
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "File type must be pdf.");
                }

                var processor = new PDFProcessor();

                var bytes1 = fileContent1.ReadAsByteArrayAsync().Result;
                var bytes2 = fileContent2.ReadAsByteArrayAsync().Result;

                var result = processor.MergeMem(bytes1, bytes2);

                var response = Request.CreateResponse(HttpStatusCode.OK);
                response.Content = new ByteArrayContent(result);
                response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
                response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
                {
                    FileName = "Merged" + DateTime.Now.Ticks + ".pdf"
                };

                return response;
            }
            catch (PrinterException ex)
            {
                Console.WriteLine(ex.Message + "ErrorCode: " + ex.ErrorCode);
                throw;
            }
        }
    }
}
