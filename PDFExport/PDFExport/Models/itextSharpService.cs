using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace PDFExport.Models
{
    public class itextSharpService
    {
        public MemoryStream testPDF() {
            MemoryStream memory = new MemoryStream();
            Document document = new Document(PageSize.A4, -30, -30, 5, 0);
            //Document document = new Document(PageSize.A4);
            PdfWriter pdfWriter = PdfWriter.GetInstance(document, memory);

            pdfWriter.Open();
            document.Open();
            //標題
            document.Add(setStringTH(" 履 歷 表 ", 25, Element.ALIGN_CENTER, Font.NORMAL));
            document.Add(setString(" ", 15, Element.ALIGN_CENTER, Font.NORMAL));

            //基本資料
            PdfPTable pTableSchool_PS = new PdfPTable(new float[] { 1f });
            Font font = new Font(BaseFont.CreateFont("C:\\Windows\\Fonts\\kaiu.ttf", BaseFont.IDENTITY_H, BaseFont.EMBEDDED));
            font.SetColor(255,255,255);
            font.Size=13;
            PdfPCell cell = new PdfPCell(new Phrase("基 本 資 料", font));
            cell.BackgroundColor = new BaseColor(22, 54, 93);
            cell.HorizontalAlignment = Element.ALIGN_CENTER;
            cell.MinimumHeight = 25f;
            //cell.Rowspan = 1;
            //cell.Colspan = 1;
            //pTableSchool_PS.AddCell(setCellH("基 本 資 料", 1, 1, Element.ALIGN_CENTER, Font.NORMAL, 12, 15f));
            pTableSchool_PS.AddCell(cell);
            document.Add(pTableSchool_PS);

            //基本資料_內容
            //列1
            PdfPTable pdfTable01 =new PdfPTable(new float[] { 1.5f, 3f, 1.5f, 2f, 1.6f});
            pdfTable01.AddCell(setCell("姓　　　　　名", 1, 1, Element.ALIGN_CENTER, Font.NORMAL, 10, 15f));
            pdfTable01.AddCell(setCell("孫鵬洋", 1, 1, Element.ALIGN_LEFT, Font.NORMAL, 10, 15f));
            pdfTable01.AddCell(setCell("性　　　　　別", 1, 1, Element.ALIGN_CENTER, Font.NORMAL, 10, 15f));
            pdfTable01.AddCell(setCell("■男 □女", 1, 1, Element.ALIGN_LEFT, Font.NORMAL, 10, 15f));

            Image img01 = iTextSharp.text.Image.GetInstance("C:/Users/412/Desktop/PDFExport/PDF_export/PDFExport/PDFExport/Models/Test.jpg");
            img01.ScalePercent(7.8f);
            PdfPCell cell01 = new PdfPCell(img01);
            cell01.Colspan = 1;//占用欄數
            cell01.Rowspan = 5;//占用列數
            cell01.HorizontalAlignment = Element.ALIGN_CENTER;//水平位置
            cell01.VerticalAlignment = Element.ALIGN_MIDDLE;//垂直位置
            cell01.MinimumHeight = 120f;
            pdfTable01.AddCell(cell01);
            //pdfTable01.AddCell(setCell(" ", 5, 1, Element.ALIGN_CENTER, Font.NORMAL, 10, 15f));
            //列2
            pdfTable01.AddCell(setCell("生　　　　　日", 1, 1, Element.ALIGN_CENTER, Font.NORMAL, 10, 15f));
            pdfTable01.AddCell(setCell("87 年 08 月 17 日", 1, 1, Element.ALIGN_LEFT, Font.NORMAL, 10, 15f));
            pdfTable01.AddCell(setCell("身　高　體　重", 1, 1, Element.ALIGN_CENTER, Font.NORMAL, 10, 15f));
            pdfTable01.AddCell(setCell("177公分 90公斤", 1, 1, Element.ALIGN_LEFT, Font.NORMAL, 10, 15f));
            //列3
            pdfTable01.AddCell(setCell("駕　　　　　照", 1, 1, Element.ALIGN_CENTER, Font.NORMAL, 10, 15f));
            pdfTable01.AddCell(setCell("汽車、機車", 1, 1, Element.ALIGN_LEFT, Font.NORMAL, 10, 15f));
            pdfTable01.AddCell(setCell("婚　姻　狀　況", 1, 1, Element.ALIGN_CENTER, Font.NORMAL, 10, 15f));
            pdfTable01.AddCell(setCell("■未婚 □已婚", 1, 1, Element.ALIGN_LEFT, Font.NORMAL, 10, 15f));
            //列4
            pdfTable01.AddCell(setCell("兵　役　狀　況", 1, 1, Element.ALIGN_CENTER, Font.NORMAL, 10, 15f));
            pdfTable01.AddCell(setCell("未服兵役", 1, 1, Element.ALIGN_LEFT, Font.NORMAL, 10, 15f));
            pdfTable01.AddCell(setCell("聯　絡　電　話", 1, 1, Element.ALIGN_CENTER, Font.NORMAL, 10, 15f));
            pdfTable01.AddCell(setCell("家中:(02)2203_6493\n手機:0961336515", 1, 1, Element.ALIGN_LEFT, Font.NORMAL, 10, 15f));
            //列5
            pdfTable01.AddCell(setCell("聯　絡　住　址", 1, 1, Element.ALIGN_CENTER, Font.NORMAL, 10, 15f));
            pdfTable01.AddCell(setCell("新北市新莊區豐年街51巷", 1, 1, Element.ALIGN_LEFT, Font.NORMAL, 10, 15f));
            pdfTable01.AddCell(setCell("電　子　郵　件", 1, 1, Element.ALIGN_CENTER, Font.NORMAL, 10, 15f));
            pdfTable01.AddCell(setCellB("tcher.2009@gmail.com", 1, 1, Element.ALIGN_LEFT, Font.NORMAL, 10, 15f));
            //列6
            pdfTable01.AddCell(setCell("興　　　　　趣", 1, 1, Element.ALIGN_CENTER, Font.NORMAL, 10, 15f));
            pdfTable01.AddCell(setCell("\n聽音樂、運動、畫畫、看電影、打球\n ", 1, 4, Element.ALIGN_LEFT, Font.NORMAL, 10, 15f));
            document.Add(pdfTable01);
            document.Add(setString(" ", 10, Element.ALIGN_CENTER, Font.NORMAL));

            //學歷
            PdfPTable pTableSchool_ED = new PdfPTable(new float[] { 1f });
            pTableSchool_ED.AddCell(setCellH("學  歷", 1, 1, Element.ALIGN_CENTER, Font.NORMAL, 12, 20f));
            document.Add(pTableSchool_ED);

            //學歷_內容
            //列1
            PdfPTable pdfTable02 = new PdfPTable(new float[] { 1f, 2f, 2f, 2f});
            pdfTable02.AddCell(setCell("最　高　學　歷", 1, 1, Element.ALIGN_CENTER, Font.NORMAL, 10, 15f));
            pdfTable02.AddCell(setCell("起　始　日", 1, 1, Element.ALIGN_CENTER, Font.NORMAL, 10, 15f));
            pdfTable02.AddCell(setCell("學　校　名　稱", 1, 1, Element.ALIGN_CENTER, Font.NORMAL, 10, 15f));
            pdfTable02.AddCell(setCell("科　系　名　稱", 1, 1, Element.ALIGN_CENTER, Font.NORMAL, 10, 15f));
            //列2
            pdfTable02.AddCell(setCell("學    士", 1, 1, Element.ALIGN_CENTER, Font.NORMAL, 10, 15f));
            pdfTable02.AddCell(setCell("2014~2021", 1, 1, Element.ALIGN_CENTER, Font.NORMAL, 10, 15f));
            pdfTable02.AddCell(setCell("黎明技術學院", 1, 1, Element.ALIGN_CENTER, Font.NORMAL, 10, 15f));
            pdfTable02.AddCell(setCell("電機工程系", 1, 1, Element.ALIGN_CENTER, Font.NORMAL, 10, 15f));
            document.Add(pdfTable02);
            document.Add(setString(" ", 10, Element.ALIGN_CENTER, Font.NORMAL));

            //專長與專業
            PdfPTable pTableSchool_PT = new PdfPTable(new float[] { 1f });
            pTableSchool_PT.AddCell(setCellH("專長 與 專業訓練", 1, 1, Element.ALIGN_CENTER, Font.NORMAL, 12, 20f));
            document.Add(pTableSchool_PT);

            //專長與專業訓練_內容
            //列1
            PdfPTable pdfTable03 = new PdfPTable(new float[] { 1f, 6f });
            pdfTable03.AddCell(setCell("繪　　　　　畫", 1, 1, Element.ALIGN_CENTER, Font.NORMAL, 10, 15f));
            pdfTable03.AddCell(setCell(" DrowIO、autocad", 1, 1, Element.ALIGN_LEFT, Font.NORMAL, 10, 15f));
            //列2
            pdfTable03.AddCell(setCell("資　　料　　庫", 1, 1, Element.ALIGN_CENTER, Font.NORMAL, 10, 15f));
            pdfTable03.AddCell(setCell(" MySSQL、MS SQL Server", 1, 1, Element.ALIGN_LEFT, Font.NORMAL, 10, 15f));
            //列3
            pdfTable03.AddCell(setCell("程　式　語　言", 1, 1, Element.ALIGN_CENTER, Font.NORMAL, 10, 15f));
            pdfTable03.AddCell(setCell(" HTML、CSS、JS、Java、C#", 1, 1, Element.ALIGN_LEFT, Font.NORMAL, 10, 15f));
            //列4
            pdfTable03.AddCell(setCell("軟　體　操　作", 1, 1, Element.ALIGN_CENTER, Font.NORMAL, 10, 15f));
            pdfTable03.AddCell(setCell(" Windeow、Office", 1, 1, Element.ALIGN_LEFT, Font.NORMAL, 10, 15f));
            //列5
            pdfTable03.AddCell(setCell("自　我　進　修", 1, 1, Element.ALIGN_CENTER, Font.NORMAL, 10, 15f));
            pdfTable03.AddCell(setCell(" ", 1, 1, Element.ALIGN_LEFT, Font.NORMAL, 10, 15f));
            document.Add(pdfTable03);
            document.Add(setString(" ", 10, Element.ALIGN_CENTER, Font.NORMAL));

            //工作經歷
            PdfPTable pTableSchool_WE = new PdfPTable(new float[] { 1f });
            pTableSchool_WE.AddCell(setCellH("工 作 經 歷", 1, 1, Element.ALIGN_CENTER, Font.NORMAL, 12, 20f));
            document.Add(pTableSchool_WE);

            //專長與專業訓練_內容
            //列1
            PdfPTable pdfTable04 = new PdfPTable(new float[] { 2f, 1.2f,3f });
            pdfTable04.AddCell(setCell("服務單位與規模", 1, 1, Element.ALIGN_CENTER, Font.NORMAL, 10, 15f));
            pdfTable04.AddCell(setCell("職稱 / 服務時間", 1, 1, Element.ALIGN_CENTER, Font.NORMAL, 10, 15f));
            pdfTable04.AddCell(setCell("主要工作", 1, 1, Element.ALIGN_CENTER, Font.NORMAL, 10, 15f));
            //列2
            pdfTable04.AddCell(setCell("嘉展冷凍空調設備\n有限公司: 約18人", 1, 1, Element.ALIGN_LEFT, Font.NORMAL, 10, 15f));
            pdfTable04.AddCell(setCell("助手\n2017/6 ~ 2021/9\n2018/6 ~ 2018/8", 1, 1, Element.ALIGN_CENTER, Font.NORMAL, 10, 15f));
            pdfTable04.AddCell(setCell("冷氣安裝及維修", 1, 1, Element.ALIGN_LEFT, Font.NORMAL, 10, 15f));
            //列3
            pdfTable04.AddCell(setCell("\n順嘉電機\n規模: 約8人\n ", 1, 1, Element.ALIGN_LEFT, Font.NORMAL, 10, 15f));
            pdfTable04.AddCell(setCell("產線人員\n2018/9 ~ 2019/8", 1, 1, Element.ALIGN_CENTER, Font.NORMAL, 10, 15f));
            pdfTable04.AddCell(setCell("配盤配線", 1, 1, Element.ALIGN_LEFT, Font.NORMAL, 10, 15f));
            //列4
            pdfTable04.AddCell(setCell("\n鑫堡\n股份有限公司: 約30人\n ", 1, 1, Element.ALIGN_LEFT, Font.NORMAL, 10, 15f));
            pdfTable04.AddCell(setCell("產線人員\n2020/11 ~ 2021/8", 1, 1, Element.ALIGN_CENTER, Font.NORMAL, 10, 15f));
            pdfTable04.AddCell(setCell("馬達軸卡控制檢測、維修電路板", 1, 1, Element.ALIGN_LEFT, Font.NORMAL, 10, 15f));
            document.Add(pdfTable04);
            document.Add(setString(" ", 10, Element.ALIGN_CENTER, Font.NORMAL));
            document.Close();
            pdfWriter.Close();
            

            return memory;
        }
        private PdfPCell setCell(string str, int rowspan, int colspan, int hAlign, int style, int size, float height)
        {
            PdfPCell tempCell = new PdfPCell(setString(str, size, Element.ALIGN_CENTER, style));
            tempCell.HorizontalAlignment = hAlign;//水平位置
            tempCell.VerticalAlignment = Element.ALIGN_MIDDLE;//垂直位置
            tempCell.Colspan = colspan;//占用欄數
            tempCell.Rowspan = rowspan;//占用列數
            tempCell.MinimumHeight = 22;//最小高度

            return tempCell;
        }

        //設定文字
        private Paragraph setString(string str, float size, int align, int style)
        {
            //設定字形(標楷體)
            BaseFont bfChinese = BaseFont.CreateFont(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts),
                "kaiu.ttf"), BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);

            Font chFont = new Font(bfChinese);
            chFont.SetColor(35, 35, 35);
            chFont.Size = size;//字體大小
            chFont.SetStyle(style);//字體(粗體....)
            

            Paragraph paragraph = new Paragraph(str, chFont);//建立paragraph
            paragraph.Alignment = align;//字體位置

            return paragraph;
        }
        //藍字風格
        private PdfPCell setCellB(string str, int rowspan, int colspan, int hAlign, int style, int size, float height)
        {
            PdfPCell tempCell = new PdfPCell(setStringB(str, size, Element.ALIGN_CENTER, style));
            tempCell.HorizontalAlignment = hAlign;//水平位置
            tempCell.VerticalAlignment = Element.ALIGN_MIDDLE;//垂直位置
            tempCell.Colspan = colspan;//占用欄數
            tempCell.Rowspan = rowspan;//占用列數
            tempCell.MinimumHeight = height;//最小高度

            return tempCell;
        }
        private Paragraph setStringB(string str, float size, int align, int style)
        {
            //設定字形(標楷體)
            BaseFont bfChinese = BaseFont.CreateFont(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts),
                "kaiu.ttf"), BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);

            Font chFont = new Font(bfChinese);
            chFont.SetColor(0, 95, 234);
            chFont.Size = size;//字體大小
            chFont.SetStyle(style);//字體(粗體....)


            Paragraph paragraph = new Paragraph(str, chFont);//建立paragraph
            paragraph.Alignment = align;//字體位置

            return paragraph;
        }
        //標題風格
        private PdfPCell setCellH(string str, int rowspan, int colspan, int hAlign, int style, int size, float height)
        {
            PdfPCell tempCell = new PdfPCell(setStringH(str, size, Element.ALIGN_CENTER, style));
            tempCell.HorizontalAlignment = hAlign;//水平位置
            //tempCell.VerticalAlignment = Element.ALIGN_MIDDLE;//垂直位置
            tempCell.Colspan = colspan;//占用欄數
            tempCell.Rowspan = rowspan;//占用列數
            tempCell.MinimumHeight = 25;//最小高度
            tempCell.BackgroundColor = new BaseColor(22, 54, 93);
            return tempCell;
        }
        private Paragraph setStringH(string str, float size, int align, int style)
        {
            //設定字形(標楷體)
            BaseFont bfChinese = BaseFont.CreateFont(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts),
                "kaiu.ttf"), BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);

            Font chFont = new Font(bfChinese);
            chFont.SetColor(255,255,255);
            chFont.Size = 13;//字體大小
            chFont.SetStyle(style);//字體(粗體....)


            Paragraph paragraph = new Paragraph(str, chFont);//建立paragraph
            paragraph.Alignment = align;//字體位置

            return paragraph;
        }
        private Paragraph setStringBH(string str, float size, int align, int style)
        {
            //設定字形(標楷體)
            BaseFont bfChinese = BaseFont.CreateFont(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts),
                "kaiu.ttf"), BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);

            Font chFont = new Font(bfChinese);
            chFont.SetColor(0,0,0);
            chFont.Size = size;//字體大小
            chFont.SetStyle("bold");//字體(粗體....)


            Paragraph paragraph = new Paragraph(str, chFont);//建立paragraph
            paragraph.Alignment = align;//字體位置

            return paragraph;
        }
        private Paragraph setStringTH(string str, float size, int align, int style)
        {
            //設定字形(標楷體)
            BaseFont bfChinese = BaseFont.CreateFont(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts),
                "kaiu.ttf"), BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);

            Font chFont = new Font(bfChinese);
            chFont.SetColor(35, 35, 35);
            chFont.Size = size;//字體大小
            chFont.SetStyle("bold");//字體(粗體....)


            Paragraph paragraph = new Paragraph(str, chFont);//建立paragraph
            paragraph.Alignment = align;//字體位置

            return paragraph;
        }
    }
}



