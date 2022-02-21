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
            //Document document = new Document(PageSize.A4.Rotate(), -70, -70, 5, 0);利用rotate()實現橫向A4文件
            Document document = new Document();
            PdfWriter pdfWriter = PdfWriter.GetInstance(document, memory);

            pdfWriter.Open();
            document.Open();
            //標題
            document.Add(setString(" 履 歷 表 ", 25, Element.ALIGN_CENTER, Font.NORMAL));
            document.Add(setString(" ", 15, Element.ALIGN_CENTER, Font.NORMAL));

            //學校_PS
            PdfPTable pTableSchool_PS = new PdfPTable(new float[] { 1f });
            pTableSchool_PS.AddCell(setCell("基本資料", 1, 1, Element.ALIGN_CENTER, Font.NORMAL, 12, 20f));
            document.Add(pTableSchool_PS);

            //資料表
            PdfPTable pdfTable1 = new PdfPTable(new float[] { 2f, 3f, 2f, 2f, 2f, 3f });
            //th 第一列
            pdfTable1.AddCell(setCell("流水\n編號", 1, 1, Element.ALIGN_CENTER, Font.NORMAL, 10, 15f));
            pdfTable1.AddCell(setCell("科組", 1, 1, Element.ALIGN_CENTER, Font.NORMAL, 10, 15f));
            pdfTable1.AddCell(setCell("核定\n招生\n名額", 1, 1, Element.ALIGN_CENTER, Font.NORMAL, 10, 15f));
            //下行寫法代表會佔兩列的空間，也就是會多項右多佔一列
            pdfTable1.AddCell(setCell("內含名額", 2, 1, Element.ALIGN_CENTER, Font.NORMAL, 10, 15f));
            //下行寫法代表會佔兩行的空間，也就是會多向右多佔一列
            pdfTable1.AddCell(setCell("一般生外加名額", 1, 2, Element.ALIGN_CENTER, Font.NORMAL, 10, 15f));
            //pdfTable1.AddCell(setCell("五專特種生（外加）", 1, 1, Element.ALIGN_CENTER, Font.NORMAL, 10, 15f));
            //th 第二列
            pdfTable1.AddCell(setCell("test1", 1, 1, Element.ALIGN_CENTER, Font.NORMAL, 10, 15f));
            pdfTable1.AddCell(setCell("test2", 1, 1, Element.ALIGN_CENTER, Font.NORMAL, 10, 15f));
            pdfTable1.AddCell(setCell("test3", 1, 1, Element.ALIGN_CENTER, Font.NORMAL, 10, 15f));
            //pdfTable1.AddCell(setCell("test4", 1, 1, Element.ALIGN_CENTER, Font.NORMAL, 10, 15f));
            pdfTable1.AddCell(setCell("test5", 1, 1, Element.ALIGN_CENTER, Font.NORMAL, 10, 15f));
            pdfTable1.AddCell(setCell("test6", 1, 1, Element.ALIGN_CENTER, Font.NORMAL, 10, 15f));
            document.Add(pdfTable1);

            //尾部簽核部分
            PdfPTable pTableBottom = new PdfPTable(new float[] { 1f });
            pTableBottom.AddCell(setCell("核章:", 1, 1, Element.ALIGN_LEFT, Font.NORMAL, 10, 15f));
            pTableBottom.AddCell(setCell("", 1, 1, Element.ALIGN_LEFT, Font.NORMAL, 10, 15f));
            pTableBottom.AddCell(setCell("填表人:                              填表單位主管:                                                                                          " + DateTime.Now.Year + "年" + DateTime.Now.Month + "月" + DateTime.Now.Day + "日", 1, 1, Element.ALIGN_LEFT, Font.NORMAL, 10, 15f));
            document.Add(pTableBottom);
            document.Add(iTextSharp.text.Image.GetInstance("D:/Codding/PDFexprot/PDFExport/PDFExport/Models/Test.jpg"));
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
            tempCell.MinimumHeight = height;//最小高度

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
    }
}



