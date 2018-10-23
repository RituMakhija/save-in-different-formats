using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Drop_Down.Models;
using System.Data.Entity;
using System.Data;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Reflection;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Microsoft.Office.Interop.Word;

namespace Drop_Down.Controllers
{
    public class Drop_Down_To_SaveController : Controller
    {
        // GET: Drop_Down_To_Save

        public ActionResult Index()
        {
            SaveType sv = new SaveType();
            testEntities ent = new testEntities();
            sv.savedFormatlst = ent.tbl_saveFormat;
            
            //List<SelectListItem> lst = new List<SelectListItem>();
            //foreach(var a in sv.savedFormatlst.ToList())
            //{
            //    SelectListItem lsti = new SelectListItem
            //    {
            //        Text = a.FormatType,
            //        Value = a.id.ToString()
            //    };
            //    lst.Add(lsti);
            //}

            ViewBag.ExemploList = new SelectList(sv.savedFormatlst,"id","FormatType");
            return View();
        }
        [HttpPost]
        public ActionResult Index(SaveType sv)
        {
            int sel = sv.saved;
            if (sel == 1)
            {
                testEntities te = new testEntities();
                IEnumerable<tbl_entryDate> ied = te.tbl_entryDate.ToList();
                System.Data.DataTable dataTable = new System.Data.DataTable(typeof(tbl_entryDate).Name);
                PropertyInfo[] Props = typeof(tbl_entryDate).GetProperties(BindingFlags.Public | BindingFlags.Instance);
                foreach (PropertyInfo prop in Props)
                {
                    var type = (prop.PropertyType.IsGenericType && prop.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>) ? Nullable.GetUnderlyingType(prop.PropertyType) : prop.PropertyType);
                    dataTable.Columns.Add(prop.Name, type);
                }
                foreach (tbl_entryDate item in ied)
                {
                    var values = new object[Props.Length];
                    for (int i = 0; i < Props.Length; i++)
                    {
                        values[i] = Props[i].GetValue(item, null);
                    }
                    dataTable.Rows.Add(values);
                }
                string destinationPath = @"D:\Diksha\Drop_Down\Drop_Down\Files\" + System.DateTime.Now.ToString("dd-MMM-yyyy-hh-mm-ss") + ".pdf";
                iTextSharp.text.Document document = new iTextSharp.text.Document();
                PdfWriter writer = PdfWriter.GetInstance(document, new FileStream(destinationPath, FileMode.Create));
                document.Open();

                PdfPTable table = new PdfPTable(dataTable.Columns.Count);
                table.WidthPercentage = 100;

                //Set columns names in the pdf file
                for (int k = 0; k < dataTable.Columns.Count; k++)
                {
                    PdfPCell cell = new PdfPCell(new Phrase(dataTable.Columns[k].ColumnName));

                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.BackgroundColor = new iTextSharp.text.BaseColor(51, 102, 102);

                    table.AddCell(cell);
                }

                //Add values of DataTable in pdf file
                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    for (int j = 0; j < dataTable.Columns.Count; j++)
                    {
                        PdfPCell cell = new PdfPCell(new Phrase(dataTable.Rows[i][j].ToString()));

                        //Align the cell in the center
                        cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                        cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;

                        table.AddCell(cell);
                    }
                }

                document.Add(table);
                document.Close();
            }
            else if (sel == 2)
            {
                testEntities te = new testEntities();
                IEnumerable<tbl_entryDate> ied = te.tbl_entryDate.ToList();
                System.Data.DataTable dt = new System.Data.DataTable(typeof(tbl_entryDate).Name);
                PropertyInfo[] Props = typeof(tbl_entryDate).GetProperties(BindingFlags.Public | BindingFlags.Instance);
                foreach (PropertyInfo prop in Props)
                {
                    var type = (prop.PropertyType.IsGenericType && prop.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>) ? Nullable.GetUnderlyingType(prop.PropertyType) : prop.PropertyType);
                    dt.Columns.Add(prop.Name, type);
                }
                foreach (tbl_entryDate item in ied)
                {
                    var values = new object[Props.Length];
                    for (int i = 0; i < Props.Length; i++)
                    {
                        values[i] = Props[i].GetValue(item, null);
                    }
                    dt.Rows.Add(values);
                }
                string ExcelFilePath = @"D:\Diksha\Drop_Down\Drop_Down\Files\" + System.DateTime.Now.ToString("dd-MMM-yyyy-hh-mm-ss") + ".xlsx";
                try
                {
                    int ColumnsCount;

                    if (dt == null || (ColumnsCount = dt.Columns.Count) == 0)
                        throw new Exception("ExportToExcel: Null or empty input table!\n");
                    Microsoft.Office.Interop.Excel.Application Excel = new Microsoft.Office.Interop.Excel.Application();
                    Excel.Workbooks.Add();
                    Microsoft.Office.Interop.Excel._Worksheet Worksheet = Excel.ActiveSheet;
                    object[] Header = new object[ColumnsCount];
                    for (int i = 0; i < ColumnsCount; i++)
                        Header[i] = dt.Columns[i].ColumnName;
                    Microsoft.Office.Interop.Excel.Range HeaderRange = Worksheet.get_Range((Microsoft.Office.Interop.Excel.Range)(Worksheet.Cells[1, 1]), (Microsoft.Office.Interop.Excel.Range)(Worksheet.Cells[1, ColumnsCount]));
                    HeaderRange.Value = Header;
                    HeaderRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                    HeaderRange.Font.Bold = true;
                    int RowsCount = dt.Rows.Count;
                    object[,] Cells = new object[RowsCount, ColumnsCount];
                    for (int j = 0; j < RowsCount; j++)
                        for (int i = 0; i < ColumnsCount; i++)
                            Cells[j, i] = dt.Rows[j][i];

                    Worksheet.get_Range((Microsoft.Office.Interop.Excel.Range)(Worksheet.Cells[2, 1]), (Microsoft.Office.Interop.Excel.Range)(Worksheet.Cells[RowsCount + 1, ColumnsCount])).Value = Cells;
                    if (ExcelFilePath != null && ExcelFilePath != "")
                    {
                        try
                        {
                            Worksheet.SaveAs(ExcelFilePath);
                            Excel.Quit();
                        }
                        catch (Exception ex)
                        {
                            throw new Exception("ExportToExcel: Excel file could not be saved! Check filepath.\n"
                                + ex.Message);
                        }
                    }
                    else
                    {
                        Excel.Visible = true;
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception("ExportToExcel: \n" + ex.Message);
                }
            }
            else
            {
                testEntities te = new testEntities();
                IEnumerable<tbl_entryDate> ied = te.tbl_entryDate.ToList();
                System.Data.DataTable dt = new System.Data.DataTable(typeof(tbl_entryDate).Name);
                PropertyInfo[] Props = typeof(tbl_entryDate).GetProperties(BindingFlags.Public | BindingFlags.Instance);
                foreach (PropertyInfo prop in Props)
                {
                    var type = (prop.PropertyType.IsGenericType && prop.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>) ? Nullable.GetUnderlyingType(prop.PropertyType) : prop.PropertyType);
                    dt.Columns.Add(prop.Name, type);
                }
                foreach (tbl_entryDate item in ied)
                {
                    var values = new object[Props.Length];
                    for (int i = 0; i < Props.Length; i++)
                    {
                        values[i] = Props[i].GetValue(item, null);
                    }
                    dt.Rows.Add(values);
                }
                try
                {
                    //Create an instance for word app
                    Microsoft.Office.Interop.Word.Application winword = new Microsoft.Office.Interop.Word.Application();

                    //Set animation status for word application
                    winword.ShowAnimation = false;

                    //Set status for word application is to be visible or not.
                    winword.Visible = false;

                    //Create a missing variable for missing value
                    object missing = System.Reflection.Missing.Value;

                    //Create a new document
                    Microsoft.Office.Interop.Word.Document document = winword.Documents.Add(ref missing, ref missing, ref missing, ref missing);

                    //Add header into the document
                    foreach (Microsoft.Office.Interop.Word.Section section in document.Sections)
                    {
                        //Get the header range and add the header details.
                        Microsoft.Office.Interop.Word.Range headerRange = section.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                        headerRange.Fields.Add(headerRange, Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage);
                        headerRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        headerRange.Font.ColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdBlue;
                        headerRange.Font.Size = 10;
                        headerRange.Text = "Header text goes here";
                    }

                    //Add the footers into the document
                    foreach (Microsoft.Office.Interop.Word.Section wordSection in document.Sections)
                    {
                        //Get the footer range and add the footer details.
                        Microsoft.Office.Interop.Word.Range footerRange = wordSection.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                        footerRange.Font.ColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdDarkRed;
                        footerRange.Font.Size = 10;
                        footerRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        footerRange.Text = "Footer text goes here";
                    }

                    //adding text to document
                    document.Content.SetRange(0, 0);
                    document.Content.Text = "This is test document " + Environment.NewLine;

                    //Add paragraph with Heading 1 style
                    Microsoft.Office.Interop.Word.Paragraph para1 = document.Content.Paragraphs.Add(ref missing);
                    object styleHeading1 = "Heading 1";
                    para1.Range.set_Style(ref styleHeading1);
                    para1.Range.Text = "Para 1 text";
                    para1.Range.InsertParagraphAfter();

                    ////Add paragraph with Heading 2 style
                    //Microsoft.Office.Interop.Word.Paragraph para2 = document.Content.Paragraphs.Add(ref missing);
                    //object styleHeading2 = "Heading 2";
                    //para2.Range.set_Style(ref styleHeading2);
                    //para2.Range.Text = "Para 2 text";
                    //para2.Range.InsertParagraphAfter();

                    //Create a 5X5 table and insert some dummy record
                    Table firstTable = document.Tables.Add(para1.Range, dt.Rows.Count+1, dt.Columns.Count, ref missing, ref missing);

                    firstTable.Borders.Enable = 1;
                    //foreach (Row row in firstTable.Rows)
                    //{

                    //        foreach (Cell cell in row.Cells)
                    //        {

                    //            //Header row
                    //            if (cell.RowIndex == 1)
                    //            {

                    //                cell.Range.Text =dt.Columns[0].ColumnName;
                    //                cell.Range.Font.Bold = 1;
                    //                //other format properties goes here
                    //                cell.Range.Font.Name = "verdana";
                    //                cell.Range.Font.Size = 10;
                    //                //cell.Range.Font.ColorIndex = WdColorIndex.wdGray25;                            
                    //                cell.Shading.BackgroundPatternColor = WdColor.wdColorGray25;
                    //                //Center alignment for the Header cells
                    //                cell.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                    //                cell.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    //            }

                    //            //Data row
                    //            else
                    //            {
                    //                for (int i = 0; i < dt.Rows.Count; i++)
                    //                {
                    //                    for (int j = 0; j < dt.Columns.Count; j++)
                    //                    {
                    //                        cell.Range.Text = dt.Rows[i][j].ToString();
                    //                    }
                    //                }
                    //            }
                    //        }                                                      

                    //}


                    for (int i = 0; i < dt.Rows.Count+1; i++)
                    {
                        Row row = firstTable.Rows[i+1];
                        for (int k = 0; k < dt.Columns.Count; k++)
                        {
                            Cell cell = row.Cells[k+1];
                            //Header row
                            if (i == 0)
                            {

                                cell.Range.Text = dt.Columns[k].ColumnName;
                                cell.Range.Font.Bold = 1;
                                //other format properties goes here
                                cell.Range.Font.Name = "verdana";
                                cell.Range.Font.Size = 10;
                                //cell.Range.Font.ColorIndex = WdColorIndex.wdGray25;                            
                                cell.Shading.BackgroundPatternColor = WdColor.wdColorGray25;
                                //Center alignment for the Header cells
                                cell.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                                cell.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                            }
                            else
                            {
                                int j = i - 1;


                                cell.Range.Text = dt.Rows[j][k].ToString();
                            }
                            
                        }
                    }



                    //Save the document
                    object filename = @"D:\Diksha\Drop_Down\Drop_Down\Files\" + System.DateTime.Now.ToString("dd - MMM - yyyy - hh - mm - ss") + ".doc";
                    document.SaveAs2(ref filename);
                    document.Close(ref missing, ref missing, ref missing);
                    document = null;
                    winword.Quit(ref missing, ref missing, ref missing);
                    winword = null;
                   
                }
                catch (Exception ex)
                {
                    throw new Exception("ExportToDocument: \n" + ex.Message);
                }

            }
            return RedirectToAction("Index");
        }      
    }
}