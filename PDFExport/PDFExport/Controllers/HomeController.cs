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

// 這是要而外載的
using HtmlAgilityPack;

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