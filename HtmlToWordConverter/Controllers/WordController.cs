using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using NotesFor.HtmlToOpenXml;
using System;
using System.IO;
using System.Web.Mvc;

namespace HtmlToWordConverter.Controllers
{
    public class WordController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult GenerateWord()
        {
            Download();
            return Json(true, JsonRequestBehavior.AllowGet);
        }

        private static void Download()
        {
            var guid = Guid.NewGuid().ToString();
            string filename = @"D:\test" + guid + ".docx";
            string html = Resource1.sample;

            if (System.IO.File.Exists(filename)) System.IO.File.Delete(filename);

            using (MemoryStream generatedDocument = new MemoryStream())
            {
                using (WordprocessingDocument package = WordprocessingDocument.Create(generatedDocument, WordprocessingDocumentType.Document))
                {
                    MainDocumentPart mainPart = package.MainDocumentPart;
                    if (mainPart == null)
                    {
                        mainPart = package.AddMainDocumentPart();
                        new Document(new Body()).Save(mainPart);
                    }

                    HtmlConverter converter = new HtmlConverter(mainPart)
                    {
                        BaseImageUrl = new Uri("http://localhost:52798/")
                    };
                    converter.ParseHtml(html);
                    
                    mainPart.Document.Save();
                }

                System.IO.File.WriteAllBytes(filename, generatedDocument.ToArray());
            }

            System.Diagnostics.Process.Start(filename);
        }
    }
}