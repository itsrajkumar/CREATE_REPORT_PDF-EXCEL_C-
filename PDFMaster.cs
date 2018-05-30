using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using iTextSharp;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Data;
using System.IO;

namespace DasTec.Factory.LycheeGST.Model
{
    class PDFMaster
    {

        public PDFMaster()
        {

        }
        
        public void createPDF(String strHeader, Dictionary<string, string> dict,String FromDate,String EndDate, DataTable dt) 
        {


            Random rnd = new Random();
            int randomNo = rnd.Next(1, 9999);
            string path = Path.GetTempPath() + "tempPdf" + randomNo + ".pdf";
            if (File.Exists(path))
                File.Delete(path);
            System.IO.FileStream fs = new FileStream(path, FileMode.Create);
            // Create an instance of the document class which represents the PDF document itself.
            Document document = new Document(PageSize.A4, 25, 25, 30, 30);

            // Create an instance to the PDF file by creating an instance of the PDF 
            // Writer class using the document and the filestrem in the constructor.
            PdfWriter writer = PdfWriter.GetInstance(document, fs);

            PdfPCell cell, cell1, cell2, cell3;  

            // Add meta information to the document
            document.AddAuthor("DasTec");
            document.AddCreator("Supported By DasTec Solution pvt. Ltd.");
            document.AddKeywords("PDF Master");
            document.AddSubject("Document subject - Describing the steps creating a PDF document");
            document.AddTitle(strHeader);

            // Open the document to enable you to write to the document
            document.Open();

            PdfPTable pt = new PdfPTable(1);
            PdfPTable headert = new PdfPTable(2);


            //Report Header
            BaseFont bfntHead = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            Font fntHead = new Font(bfntHead, 15, 1, BaseColor.GRAY);
            Paragraph prgHeading = new Paragraph();
            prgHeading.Alignment = Element.ALIGN_CENTER;
            prgHeading.Add(new Chunk(strHeader.ToUpper(), fntHead));
            document.Add(prgHeading);

            //Author
            Paragraph prgAuthor = new Paragraph();
            BaseFont btnAuthor = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            Font fntAuthor = new Font(btnAuthor,10, 2, BaseColor.GRAY);
            prgAuthor.Alignment = Element.ALIGN_CENTER;

            for (int i = 0; i < dict.Count; i++)
            {
               /* Console.WriteLine("Key: {0}, Value: {1}",
                                                        dict.Keys.ElementAt(i),
                                                        dict[dict.Keys.ElementAt(i)]);*/

                prgAuthor.Add(new Chunk(dict.Keys.ElementAt(i)+" : "+ dict[dict.Keys.ElementAt(i)]+"\n", fntAuthor));
                
            }
            prgAuthor.Add(new Chunk("Date : " + FromDate+" TO "+EndDate, fntAuthor));

            document.Add(prgAuthor);

            //Add a line seperation
            Paragraph p = new Paragraph(new Chunk(new iTextSharp.text.pdf.draw.LineSeparator(0.0F, 100.0F, BaseColor.BLACK, Element.ALIGN_LEFT, 1)));
            document.Add(p);

            //Add line break
            document.Add(new Chunk("\n", fntHead));

            //Write the table
            PdfPTable table = new PdfPTable(dt.Columns.Count);
            table.WidthPercentage = 100;

            if (dt.Columns.Count == 6)
            {
                var colWidthPercentages = new[] { 15f, 20f, 35f, 10f, 10f, 10f };
                table.SetWidths(colWidthPercentages);
            }
            //Table header
            BaseFont btnColumnHeader = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
          
            Font fntColumnHeader = new Font(btnColumnHeader, 8, 1, BaseColor.WHITE);
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                cell = new PdfPCell();
                cell.BackgroundColor = BaseColor.GRAY;
                cell.AddElement(new Chunk(dt.Columns[i].ColumnName.ToUpper(), fntColumnHeader));
                table.AddCell(cell);
            }
            //table Data
            Font fntColumnDetail = new Font(btnColumnHeader, 7, 1, BaseColor.BLACK);
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    cell = new PdfPCell();
                    cell.BackgroundColor = BaseColor.WHITE;
                    cell.AddElement(new Chunk(dt.Rows[i][j].ToString(), fntColumnDetail));
                    table.AddCell(cell);
                    //table.AddCell(dt.Rows[i][j].ToString());
                }
            }

            document.Add(table);













            pt.WidthPercentage = 100;
            float[] width = { 100f };
            pt.SetWidths(width);
            pt.SpacingAfter = 0;
            pt.SpacingBefore = 0;

            //OnEndPage(writer, document);

            cell2 = new PdfPCell(new Phrase(" *** This s a system generated Invoice *** ", fntColumnDetail));
            cell2.HorizontalAlignment = Element.ALIGN_CENTER;
            //cell2.Colspan = 14;
            cell2.Border = 0;
            cell2.Padding = 0;
            pt.AddCell(cell2);
            document.Add(pt);

            
            // Close the document
            document.Close();
            // Close the writer instance
            writer.Close();
            // Always close open filehandles explicity
            fs.Close();







            System.Diagnostics.Process.Start(path);

        }


        public void OnEndPage(PdfWriter writer, Document document)
        {
            //base.OnEndPage(writer, document);

            var content = writer.DirectContent;
            var pageBorderRect = new Rectangle(document.PageSize);

            pageBorderRect.Left += document.LeftMargin;
            pageBorderRect.Right -= document.RightMargin;
            pageBorderRect.Top -= document.TopMargin;
            pageBorderRect.Bottom += document.BottomMargin;

            content.SetColorStroke(BaseColor.BLACK);
            content.Rectangle(pageBorderRect.Left, pageBorderRect.Bottom, pageBorderRect.Width, pageBorderRect.Height);
            content.Stroke();
        }
    }
}
