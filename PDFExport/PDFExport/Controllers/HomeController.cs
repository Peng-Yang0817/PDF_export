using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using iTextSharp;
using PDFExport.Models;

namespace PDFExport.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult madePDF() {
            itextSharpService temp = new itextSharpService();
            return File(temp.testPDF().ToArray(), "Application/pdf", "test.pdf");
        }
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}