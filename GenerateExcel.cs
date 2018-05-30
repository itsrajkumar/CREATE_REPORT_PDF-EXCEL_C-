using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace DasTec.Factory.LycheeGST.Model.ExportExcel
{
    class GenerateExcel
    {
        Excel.Application xlApp;
        Excel.Workbook xlWorkBook;
        Excel.Worksheet xlWorkSheet;
        object misValue = System.Reflection.Missing.Value;
        Excel.Range chartRange;
        int worksheetCount = 1;

        List<GenerateExcelData> data { get; set; }

        public GenerateExcel(List<GenerateExcelData> data)
        {
            this.data = data;
        }


        public void getExcel()
        {
           // xlApp = new Excel.Application();
           // xlWorkBook = xlApp.Workbooks.Add(misValue);

            createWorksheet();

          //  System.Windows.Forms.SaveFileDialog saveDlg = new System.Windows.Forms.SaveFileDialog();
           // saveDlg.InitialDirectory = @"C:\";
          //  saveDlg.Filter = "Excel files (*.xlsx)|*.xlsx";
          //  saveDlg.FilterIndex = 0;
           // saveDlg.RestoreDirectory = true;
          //  saveDlg.Title = "Export Excel File To";
          //  if (saveDlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
          //  {
           //     string path = saveDlg.FileName;
            //    xlWorkBook.SaveCopyAs(path);
             //   xlWorkBook.Saved = true;
            //    xlWorkBook.Close(true, misValue, misValue);
            //    xlApp.Quit();
            //    MessageBox.Show("Excel file created and saved");
           // }
           // releaseObject(xlWorkSheet);
           // releaseObject(xlWorkBook);
           // releaseObject(xlApp);
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;

            }
            finally
            {
                GC.Collect();
            }
        }


        void createWorksheet()
        {
            using (ExcelPackage objExcelPackage = new ExcelPackage())
            {
                  
            foreach (GenerateExcelData d in data)
            {
               // if (worksheetCount > xlWorkBook.Worksheets.Count)
              //  {
              //      xlWorkBook.Worksheets.Add(After: xlWorkBook.Sheets[xlWorkBook.Sheets.Count]);
            //    }
               // xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(worksheetCount);
              //  xlWorkSheet.Name = d.worksheetName;
               // if(!String.IsNullOrEmpty(d.excelHeader))
                 //   createHeader(d, "YELLOW", true, 10, "n");
             //   mapData(d);
                   worksheetCount++;

                  //Create the worksheet    
                        ExcelWorksheet objWorksheet = objExcelPackage.Workbook.Worksheets.Add(d.worksheetName);
                        //Load the datatable into the sheet, starting from cell A1. Print the column names on row 1 

                        objWorksheet.Name = d.worksheetName;
                        // objWorksheet.Cells[1, 1] =dtSrc;
                        objWorksheet.Cells["A1"].LoadFromDataTable(d.worksheetData, true);
                        // objWorksheet.Cells.Style.Font.SetFromFont(new Font("Calibri", 10));    
                       // objWorksheet.Cells.AutoFitColumns();    
                        //Format the header   
                        if (d.worksheetData.Columns[4].DataType.Name.ToString() != "Object" && d.worksheetData.Columns[4].DataType.Name.ToString() != "Int64")
                        objWorksheet.Column(4).Style.Numberformat.Format = "yyyy-mm-dd";
                        if (d.worksheetData.Columns[5].DataType.Name.ToString() != "Object" && d.worksheetData.Columns[5].DataType.Name.ToString() != "Int64")
                        objWorksheet.Column(5).Style.Numberformat.Format = "yyyy-mm-dd";
                        using (ExcelRange objRange = objWorksheet.Cells["A1:XFD1"])
                        {
                            objRange.Style.Font.Bold = true;
                            objRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            objRange.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                           // objRange.Style.Fill.PatternType = ExcelFillStyle.LightUp;
                            objRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            objRange.Style.Fill.BackgroundColor.SetColor(Color.SkyBlue);
                            //objRange.Style.Fill.BackgroundColor.SetColor(Color.FromA#eaeaea);    
                        }

                    
                }
           

           // System.Windows.Forms.SaveFileDialog saveDlg = new System.Windows.Forms.SaveFileDialog();
          //  saveDlg.InitialDirectory = @"C:\";
          //  saveDlg.Filter = "Excel files (*.xlsx)|*.xlsx";
          //  saveDlg.FilterIndex = 0;
          //  saveDlg.RestoreDirectory = true;
          //  saveDlg.Title = "Export Excel File To";
          FileStream objFileStrm;
          //  if (saveDlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
           // {
            //    string path = saveDlg.FileName;
           //     if (File.Exists(path))
             //       File.Delete(path);

                //Create excel file on physical disk    
             //   objFileStrm = File.Create(path);
            //    objFileStrm.Close();

                //Write content to excel file    
             //   File.WriteAllBytes(path, objExcelPackage.GetAsByteArray());
             //   MessageBox.Show("Excel file created and saved");
                
             //       System.Diagnostics.Process.Start(path);
            //    
           // }
           // else 
          //  {
                     Random rnd = new Random();
                     int randomNo = rnd.Next(1, 9999);
                     string path = Path.GetTempPath() + "temp" + randomNo + ".xlsx";
                if (File.Exists(path))
                    File.Delete(path);
                //Create excel file on physical disk    
                objFileStrm = File.Create(path);
                objFileStrm.Close();

                //Write content to excel file    
                File.WriteAllBytes(path, objExcelPackage.GetAsByteArray());

                System.Diagnostics.Process.Start(path);

          //  }



            }

         


        }

        void createHeader(GenerateExcelData d, string b, bool font, int size, string fcolor)
        {
            xlWorkSheet.Cells[1, 1] = d.excelHeader;
            int unicode = 65 + d.worksheetData.Columns.Count-1;
            char character = (char)unicode;
            chartRange = xlWorkSheet.get_Range("A1", character.ToString()+"1");
            chartRange.Merge(d.worksheetData.Columns.Count-1);
            switch (b)
            {
                case "YELLOW":
                    chartRange.Interior.Color = System.Drawing.Color.Yellow.ToArgb();
                    break;
                case "GRAY":
                    chartRange.Interior.Color = System.Drawing.Color.Gray.ToArgb();
                    break;
                case "GAINSBORO":
                    chartRange.Interior.Color =
            System.Drawing.Color.Gainsboro.ToArgb();
                    break;
                case "Turquoise":
                    chartRange.Interior.Color =
            System.Drawing.Color.Turquoise.ToArgb();
                    break;
                case "PeachPuff":
                    chartRange.Interior.Color =
            System.Drawing.Color.PeachPuff.ToArgb();
                    break;
                default:
                    //  workSheet_range.Interior.Color = System.Drawing.Color..ToArgb();
                    break;
            }

            chartRange.Borders.Color = System.Drawing.Color.Black.ToArgb();
            chartRange.Font.Bold = font;
            chartRange.ColumnWidth = size;
            if (fcolor.Equals(""))
            {
                chartRange.Font.Color = System.Drawing.Color.White.ToArgb();
            }
            else
            {
                chartRange.Font.Color = System.Drawing.Color.Black.ToArgb();
            }    
        
        }

        void mapData(GenerateExcelData d)
        {
            int j = 2;
            int k = 1;
            int colCount = d.worksheetData.Columns.Count;
            int unicode = 65 + colCount - 1;
            char character = (char)unicode;
            

            if (String.IsNullOrEmpty(d.excelHeader))
                j = 1;
            foreach (DataColumn col in d.worksheetData.Columns)
            {

                object cell = col.ToString();
                xlWorkSheet.Cells[j, k++] = cell;
                

            }
            chartRange = xlWorkSheet.get_Range("A" + j.ToString(), character.ToString() + j.ToString());

            //var columnHeadingsRange = xlWorkSheet.Range[xlWorkSheet.Cells[j], xlWorkSheet.Cells[colCount]];
            chartRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightPink);
            chartRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            chartRange.Font.Bold = true;
           
            //j = startPoint;
            foreach (DataRow row in d.worksheetData.Rows)
            {
                ++j;
                for (int i = 0; i < colCount; i++)
                {
                    object cell = row[i];
                    xlWorkSheet.Cells[j, i + 1] = cell;
                }

            }
            
        }

    }

    public struct GenerateExcelData
    {
        public String excelHeader { get; set; }
        public String worksheetName { get; set; }
        public DataTable worksheetData { get; set; }
        
    }


}
