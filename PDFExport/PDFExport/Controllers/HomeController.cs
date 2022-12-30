using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using iTextSharp;
using PDFExport.Models;

using System.Net.Http;
using System.Threading.Tasks;

using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;


using HtmlAgilityPack;

// 這是要而外載的
using NPOI.HSSF.UserModel;
using System.IO;

namespace PDFExport.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult madePDF()
        {
            itextSharpService temp = new itextSharpService();
            return File(temp.testPDF().ToArray(), "Application/pdf", "test.pdf");
        }
        public ActionResult Index()
        {
            return View();
        }

        public async Task<ActionResult> Index2()
        {
            // 建立String List 來預備等等塞值
            List<string> DataList = new List<string>();

            // 產生 HttpClient 物件
            var client = new HttpClient();

            // 發出 HTTP GET 請求並取得回應
            var response = await client.GetAsync("https://oaa.ntut.edu.tw/p/403-1008-11-1.php?Lang=zh-tw");

            // 若回應的狀態碼為 200 OK，則取得網站的 HTML 內容
            if (response.IsSuccessStatusCode)
            {
                var html = await response.Content.ReadAsStringAsync();

                // 將 HTML 字串載入 HtmlDocument 物件
                var htmlDoc = new HtmlDocument();
                htmlDoc.LoadHtml(html);

                // 使用 XPath 語法提取所有 class 為 "article-title" 的標題元素
                //var titles = htmlDoc.DocumentNode.SelectNodes("//*[@class='article-title']");

                //foreach (var title in titles)
                //{
                //    Console.WriteLine(title.InnerText);
                //}

                // 使用 XPath 語法抓取所有 class 為 "d-txt" 底下的 div 元素中，且 class 為 "mtitle" 的 div 元素底下的 a 標籤的內容
                var nodes = htmlDoc.DocumentNode.SelectNodes("//div[@class='d-txt']//div[@class='mtitle']/a");

                // 若有找到標籤，則將所有標籤的內容印出
                if (nodes != null)
                {
                    string data = "";
                    foreach (var node in nodes)
                    {
                        data = node.InnerText;
                        DataList.Add(data);
                    }
                }

            }

            ViewBag.DataList = DataList;
            return View();
        }

        public async Task<ActionResult> Index3()
        {
            // 建立String List 來預備等等塞值
            List<string> DataList = new List<string>();
            List<string> DataListHref = new List<string>();
            List<string> DataListDate = new List<string>();

            // 產生 HttpClient 物件
            var client = new HttpClient();

            // 發出 HTTP GET 請求並取得回應
            var response = await client.GetAsync("https://oaa.ntut.edu.tw/p/403-1008-11-1.php?Lang=zh-tw");

            // 若回應的狀態碼為 200 OK，則取得網站的 HTML 內容
            if (response.IsSuccessStatusCode)
            {
                var html = await response.Content.ReadAsStringAsync();

                // 將 HTML 字串載入 HtmlDocument 物件
                var htmlDoc = new HtmlDocument();
                htmlDoc.LoadHtml(html);

                // 使用 XPath 語法提取所有 class 為 "article-title" 的標題元素
                //var titles = htmlDoc.DocumentNode.SelectNodes("//*[@class='article-title']");

                //foreach (var title in titles)
                //{
                //    Console.WriteLine(title.InnerText);
                //}

                // 使用 XPath 語法抓取所有 class 為 "d-txt" 底下的 div 元素中，且 class 為 "mtitle" 的 div 元素底下的 a 標籤的內容
                var nodes = htmlDoc.DocumentNode.SelectNodes("//div[@class='d-txt']//div[@class='mtitle']/a");

                // 若有找到標籤，則將所有標籤的內容印出
                if (nodes != null)
                {
                    string data = "";
                    foreach (var node in nodes)
                    {
                        data = node.InnerText;
                        DataList.Add(data);

                        data = node.Attributes["href"].Value;
                        DataListHref.Add(data);
                    }
                }

            }
            ViewBag.DataListHref = DataListHref;
            ViewBag.DataList = DataList;
            return View();
        }


        public async Task<ActionResult> Index4()
        {
            // 建立String List 來預備等等塞值
            List<string> DataList = new List<string>();
            List<string> DataListHref = new List<string>();
            List<string> DataListDate = new List<string>();
            List<string> DataListAdministration = new List<string>();

            // 產生 HttpClient 物件
            var client = new HttpClient();

            // 發出 HTTP GET 請求並取得回應
            var response = await client.GetAsync("https://oaa.ntut.edu.tw/p/403-1008-11-1.php?Lang=zh-tw");

            // 若回應的狀態碼為 200 OK，則取得網站的 HTML 內容
            if (response.IsSuccessStatusCode)
            {
                var html = await response.Content.ReadAsStringAsync();

                // 將 HTML 字串載入 HtmlDocument 物件
                var htmlDoc = new HtmlDocument();
                htmlDoc.LoadHtml(html);

                // 使用 XPath 語法提取所有 class 為 "article-title" 的標題元素
                //var titles = htmlDoc.DocumentNode.SelectNodes("//*[@class='article-title']");

                //foreach (var title in titles)
                //{
                //    Console.WriteLine(title.InnerText);
                //}

                // 使用 XPath 語法抓取所有 class 為 "d-txt" 底下的 div 元素中，且 class 為 "mtitle" 的 div 元素底下的 a 標籤的內容
                var nodes = htmlDoc.DocumentNode.SelectNodes("//div[@class='d-txt']//div[@class='mtitle']/a");

                // 若有找到標籤，則將所有標籤的內容印出
                if (nodes != null)
                {
                    string data = "";
                    foreach (var node in nodes)
                    {
                        data = node.InnerText;
                        DataList.Add(data);

                        data = node.Attributes["href"].Value;
                        DataListHref.Add(data);
                    }
                    // 使用 XPath 語法抓取所有 data-th 屬性為 "日期" 的 td 元素底下的，且 class 為 "d-txt" 的 div 元素
                    nodes = htmlDoc.DocumentNode.SelectNodes("//td[@data-th='日期']//div[@class='d-txt']");
                    foreach (var node in nodes)
                    {
                        data = node.InnerText;
                        DataListDate.Add(data);
                    }

                    // 使用 XPath 語法抓取所有 data-th 屬性為 "發佈單位" 的 td 元素底下的，且 class 為 "d-txt" 的 div 元素
                    nodes = htmlDoc.DocumentNode.SelectNodes("//td[@data-th='發佈單位']//div[@class='d-txt']");
                    foreach (var node in nodes)
                    {
                        data = node.InnerText;
                        DataListAdministration.Add(data);
                    }
                }

            }
            ViewBag.DataListAdministration = DataListAdministration;
            ViewBag.DataListDate = DataListDate;
            ViewBag.DataListHref = DataListHref;
            ViewBag.DataList = DataList;
            return View();
        }

        public async Task<ActionResult> Index5()
        {
            // 建立String List 來預備等等塞值
            List<string> DataList = new List<string>();

            // 建立 Chrome 瀏覽器物件
            var chromeDriver = new ChromeDriver();

            // 前往網頁
            chromeDriver.Navigate().GoToUrl("https://www.twitch.tv/");

            // 等待異步內容完全載入
            chromeDriver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);

            // 取得網頁 HTML 內容
            var html = await Task.Run(() => chromeDriver.PageSource);

            // 使用 HtmlAgilityPack 載入 HTML 內容
            var htmlDoc = new HtmlDocument();
            htmlDoc.LoadHtml(html);


            // 使用 XPath 語法抓取所有 class 為 "mtitle" 底下的 a 標籤
            var nodes = htmlDoc.DocumentNode.SelectNodes("//div[@class='ScTower-sc-1sjzzes-0 kwNEqQ tw-tower']/div");

            if (nodes != null)
            {
                foreach (var node in nodes)
                {
                    DataList.Add(node.InnerHtml);
                }
            }

            ViewBag.DataList = DataList;

            // 關閉 ChromeDriver 物件
            chromeDriver.Quit();

            return View();
        }

        public void ExportToExcel()
        {
            // 取得靜態資源檔案位置，並轉呈資料流
            FileStream fs = new FileStream(string.Concat(Server.MapPath("~/Models/Excel/"), "基本資料下載.xls")
            , FileMode.Open, FileAccess.ReadWrite);

            // 建立 HSSFWorkbook 物件，代表 Excel 檔案
            HSSFWorkbook templateWorkbook = new HSSFWorkbook(fs);

            // 取得第一個工作表 (Sheet)
            HSSFSheet ws = (HSSFSheet)templateWorkbook.GetSheetAt(0);

            // 複製資料行次數
            int copyCount = 3;
            for (int i = 0; i < copyCount; i++)
            {
                // 用for迴圈複製第二行，插入到第三行
                ws.CopyRow(1, 2);
            }

            // 在第二列的第一、二、三、四、五、六欄填入資料
            ws.GetRow(1).GetCell(0).SetCellValue("一號");
            ws.GetRow(1).GetCell(1).SetCellValue("力大山");
            ws.GetRow(1).GetCell(2).SetCellValue("23歲");
            ws.GetRow(1).GetCell(3).SetCellValue("女性");
            ws.GetRow(1).GetCell(4).SetCellValue("0911111111");
            ws.GetRow(1).GetCell(5).SetCellValue("新北市大寮區");

            // 設定輸出的 Content Type、檔名和下載檔案的標頭
            Response.Clear();
            Response.ContentType = "application/vnd.openxmlformats";
            Response.AddHeader("Content-Disposition", "attachment; filename=" + Server.UrlPathEncode(
                "基本資料匯出結果.xls"));

            // 將 HSSFWorkbook 寫入記憶體
            MemoryStream ms = new MemoryStream();
            templateWorkbook.Write(ms);

            // 將記憶體的資料寫回 Response.OutputStream
            ms.WriteTo(Response.OutputStream);

            // 結束回應
            Response.End();
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