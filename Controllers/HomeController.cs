using System;
using System.Linq;
using System.Web.Mvc;
using TestNagis.Models.Manager;
using System.IO;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Diagnostics;
using TestNagis.Models;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using Syncfusion.XlsIO;
using System.Drawing;
using System.Collections.Generic;
using System.ComponentModel;
using OfficeOpenXml;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Reflection;

namespace TestNagis.Controllers
{


    public class HomeController : Controller
    {


        DatabaseContext db = new DatabaseContext();


        public ActionResult Index()
        {
            var model1 = db.Transections.ToList();
            var model2 = db.Downloads.ToList();
            
/*
            //Fake data localde çalışyor ama uzak sql serverde çalışmıyor
            //muhtemelen connection stringde ekstra bir parametre eklemek gerekıyor bulamadıgım ıcın bu sekılde transection datası olusturdum
            for (int i = 0; i < 998; i++)
            {   

            
                Transection islem2 = new Transection();
                islem2.Amount = FakeData.NumberData.GetDouble();
                islem2.Buyer = FakeData.NameData.GetFirstName();
                islem2.Seller = FakeData.NameData.GetFirstName();
                islem2.Date = FakeData.DateTimeData.GetDatetime();
            db.Transections.Add(islem2);
            db.SaveChanges();
            }

    */

            //return RedirectToAction("Listele", "Home");
            return View();    
        }

        //private void ImportToExcel(DataSet ds, string strCurrentDir, string strFile)
        //{
        //    Application oXL;
        //    _Workbook oWB;
        //    _Worksheet oSheet;
        //    Range oRng;
        //    //AS we getting the Directory as a parameter it not required
        //    string strCurrentDir = Server.MapPath(".") + "\\reports\\";
        //    try
        //    {
        //        oXL = new Application();
        //        oXL.Visible = false;
        //        //Get a new workbook.
        //        oWB = (_Workbook)(oXL.Workbooks.Add(Missing.Value));
        //        oSheet = (_Worksheet)oWB.ActiveSheet;
        //        //System.Data.DataTable dtGridData=ds.Tables[0];
        //        int iRow = 2;
        //        if (ds.Tables[0].Rows.Count > 0)
        //        {
        //            //     for(int j=0;j<ds.Tables[0].Columns.Count;j++)
        //            //     {
        //            //      oSheet.Cells[1,j+1]=ds.Tables[0].Columns[j].ColumnName;
        //            //
        //            for (int j = 0; j < ds.Tables[0].Columns.Count; j++)
        //            {
        //                oSheet.Cells[1, j + 1] = ds.Tables[0].Columns[j].ColumnName;
        //            }
        //            // For each row, print the values of each column.
        //            for (int rowNo = 0; rowNo < ds.Tables[0].Rows.Count; rowNo++)
        //            {
        //                for (int colNo = 0; colNo < ds.Tables[0].Columns.Count; colNo++)
        //                {
        //                    oSheet.Cells[iRow, colNo + 1] = ds.Tables[0].Rows[rowNo][colNo].ToString();
        //                }
        //            }
        //            iRow++;
        //        }
        //        oRng = oSheet.get_Range("A1", "IV1");
        //        oRng.EntireColumn.AutoFit();
        //        oXL.Visible = false;
        //        oXL.UserControl = false;
        //        //AS we are getting the FileName as a parameter it not required
        //        //string strFile ="report"+ DateTime.Now.Ticks.ToString() +".xls";//+
        //        oWB.SaveAs(strCurrentDir +
        //   strFile, XlFileFormat.xlWorkbookNormal, null, null, false, false, XlSaveAsAccessMode.xlShared, false, false, null, null);
        //        // Need all following code to clean up and remove all references!!!
        //        oWB.Close(null, null, null);
        //        oXL.Workbooks.Close();
        //        oXL.Quit();
        //        Marshal.ReleaseComObject(oRng);
        //        Marshal.ReleaseComObject(oXL);
        //        Marshal.ReleaseComObject(oSheet);
        //        Marshal.ReleaseComObject(oWB);
        //        string strMachineName = Request.ServerVariables["SERVER_NAME"];
        //        Response.Redirect("http://" + strMachineName + "/" + "ViewNorthWindSample/reports/" + strFile);
        //    }
        //    catch (Exception theException)
        //    {
        //        Response.Write(theException.Message);
        //    }
        //}

        public void ExportListUsingEPPlus()
        {
            var data = new[]{
                               new{ Name="Ram", Email="ram@techbrij.com", Phone="111-222-3333" },
                               new{ Name="Shyam", Email="shyam@techbrij.com", Phone="159-222-1596" },
                               new{ Name="Mohan", Email="mohan@techbrij.com", Phone="456-222-4569" },
                               new{ Name="Sohan", Email="sohan@techbrij.com", Phone="789-456-3333" },
                               new{ Name="Karan", Email="karan@techbrij.com", Phone="111-222-1234" },
                               new{ Name="Brij", Email="brij@techbrij.com", Phone="111-222-3333" }
                      };

            ExcelPackage excel = new ExcelPackage();
            var workSheet = excel.Workbook.Worksheets.Add("Sheet1");
            workSheet.Cells[1, 1].LoadFromCollection(data, true);
            using (var memoryStream = new MemoryStream())
            {
                 string workingDirectory = Request.Params["APPL_PHYSICAL_PATH"];
                string workbookPath = workingDirectory + @"\ReportRepository\filename.xls";
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=Contact.xlsx");
                /*
                 string workingDirectory = Request.Params["APPL_PHYSICAL_PATH"];
                string workbookPath = workingDirectory + @"\ReportRepository\filename.xls";
                   workbook.SaveAs(workbookPath, ...);
                 */
                //excel.SaveAs(memoryStream);

                memoryStream.WriteTo(Response.OutputStream);
                Response.Flush();
                Response.End();
            }
        }

        public void WriteTsv<T>(IEnumerable<T> data, TextWriter output)
        {
            PropertyDescriptorCollection props = TypeDescriptor.GetProperties(typeof(T));
            foreach (PropertyDescriptor prop in props)
            {
                output.Write(prop.DisplayName); // header
                output.Write("\t");
            }
            output.WriteLine();
            foreach (T item in data)
            {
                foreach (PropertyDescriptor prop in props)
                {
                    output.Write(prop.Converter.ConvertToString(
                         prop.GetValue(item)));
                    output.Write("\t");
                }
                output.WriteLine();
            }
        }

        public void ExportListFromTsv()
        {
            var data = new[]{
                               new{ Name="Ram", Email="ram@techbrij.com", Phone="111-222-3333" },
                               new{ Name="Shyam", Email="shyam@techbrij.com", Phone="159-222-1596" },
                               new{ Name="Mohan", Email="mohan@techbrij.com", Phone="456-222-4569" },
                               new{ Name="Sohan", Email="sohan@techbrij.com", Phone="789-456-3333" },
                               new{ Name="Karan", Email="karan@techbrij.com", Phone="111-222-1234" },
                               new{ Name="Brij", Email="brij@techbrij.com", Phone="111-222-3333" }
                      };

            Response.ClearContent();
            Response.AddHeader("content-disposition", "attachment;filename=Contact.xls");
            Response.AddHeader("Content-Type", "application/vnd.ms-excel");
            WriteTsv(data, Response.Output);
            Response.End();
        }


        public FileResult Download(string guid)
        {
            string path = Path.Combine(Server.MapPath("~/"), ("Reports\\"+guid));
            //public static byte[] ReadAllBytes (string path);
            byte[] fileBytes = System.IO.File.ReadAllBytes(path);
            string fileName = guid;
            return File(fileBytes, System.Net.Mime.MediaTypeNames.Application.Octet, fileName);
        }
        public void TestEEP()
        {
            /*
              using (var excelFile = new ExcelPackage(targetFile))
    {
        var worksheet = excelFile.Workbook.Worksheets.Add("Sheet1");
        worksheet.Cells["A1"].LoadFromCollection(Collection: employees, PrintHeaders: true);
        excelFile.Save();
    }
             */
            using (ExcelPackage excel = new ExcelPackage())
            {
                excel.Workbook.Worksheets.Add("Worksheet1");
                //var headerrow = new List<string[]>()
                //{
                //  new string[] { "id", "first name", "last name", "dob" }
                //};
                var temp = db.Transections.ToList();
                for (int i = 0; i < temp.Count(); i++)
                {

                }
                // determine the header range (e.g. a1:d1)
                //string headerrange = "a1:" + char.ConvertFromUtf32(headerrow[0].Length + 64) + "1";

                // target a worksheet
                var worksheet = excel.Workbook.Worksheets["worksheet1"];

                // popular header row data
                //worksheet.Cells[headerrange].LoadFromArrays(headerrow);
                //worksheet.Cells["A1"].Style.Font.Bold = true;
                //worksheet.Cells["A2"].Value = "asb";
                //Set Headers and make bold
                worksheet.Cells["A1"].Value = "id";
                worksheet.Cells["B1"].Value = "Buyer";
                worksheet.Cells["C1"].Value = "Seller";
                worksheet.Cells["D1"].Value = "Amount";
                worksheet.Cells["E1"].Value = "Date";

                int RowRange = temp.Count();
                //set data from list

                for (int i = 2; i < (RowRange + 2); i++)
                {
                    string index = i.ToString();
                    worksheet.Cells["A" + index].Value = temp[i - 2].Id.ToString();
                    worksheet.Cells["B" + index].Value = temp[i - 2].Buyer.ToString();
                    worksheet.Cells["C" + index].Value = temp[i - 2].Seller.ToString();
                    worksheet.Cells["D" + index].Value = temp[i - 2].Amount.ToString();
                    worksheet.Cells["E" + index].Value = temp[i - 2].Date.ToString();
                }


                //Make all text fit the cells
                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                string path = Path.Combine(Server.MapPath("~/"), ("Reports\\" + "a.xlsx"));
                FileInfo excelFile = new FileInfo(path);//new FileInfo(@"C:\Users\amir\Desktop\test.xlsx");
                excel.SaveAs(excelFile);
            }
        }
        public void CreateDocument(DateTime? start, DateTime? end)
        {
            //Create an instance of ExcelEngine
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                var temp = db.Transections.ToList();
                DateTime startdate = start ?? new DateTime(2000, 10, 10, 1, 1, 1, 1); ;

                DateTime enddate = end ?? DateTime.Now;
                DateTime now = DateTime.Now;
                string date = now.ToShortDateString();
                string time = now.ToLongTimeString();
                date = date + "-" + time;




                //for (int i = 0; i < 10000; i++)
                //{
                //    Debug.WriteLine(i);

                //}
 
                foreach (var item in temp.ToList())
                {
                    if (item.Date < startdate || item.Date > enddate)
                    {
                        temp.Remove(item);

                    }
                }
                int RowRange =  temp.Count();

                string name = "Report_";
               // string date = now.ToString("F");
                date = date.Replace(" ", "_");
                date = date.Replace(",", "_");
                date = date.Replace(":", "-");
                date = date.Replace("/", "_");
                string sonu = ".xls";
                date += sonu;
                name += date;


                Download m = new Download();
                m.IsExist = false;
                m.CreateDate = now;
                m.EndDate = now;
                m.StartDate = now;
                m.GuidName = name;
                db.Downloads.Add(m);
                db.SaveChanges();

                IApplication application = excelEngine.Excel;

                application.DefaultVersion = ExcelVersion.Excel2016;

                //Create a workbook
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Add a picture
                //IPictureShape shape = worksheet.Pictures.AddPicture(1, 1, Server.MapPath("App_Data/AdventureCycles-Logo.png"));

                //Disable gridlines in the worksheet
                //worksheet.IsGridLinesVisible = false;
             
                //Enter values to the cells from A3 to A5
                worksheet.Range["A1"].Text = "id";
                worksheet.Range["B1"].Text = "Buyer";
                worksheet.Range["C1"].Text = "Seller";
                worksheet.Range["D1"].Text = "Amount";
                worksheet.Range["E1"].Text = "Date";


                //Make the text bold
                worksheet.Range["A1"].CellStyle.Font.Bold = true;
                worksheet.Range["B1"].CellStyle.Font.Bold = true;
                worksheet.Range["C1"].CellStyle.Font.Bold = true;
                worksheet.Range["D1"].CellStyle.Font.Bold = true;
                worksheet.Range["E1"].CellStyle.Font.Bold = true;

                for (int i = 2; i < (RowRange+2); i++)
                {
                    string index = i.ToString();
                    worksheet.Range["A" + index].Text = temp[i-2].Id.ToString();
                    worksheet.Range["B" + index].Text = temp[i-2].Buyer.ToString();
                    worksheet.Range["C" + index].Text = temp[i-2].Seller.ToString();
                    worksheet.Range["D" + index].Text = temp[i-2].Amount.ToString();
                    worksheet.Range["E" + index].Text = temp[i-2].Date.ToString();
                }
                worksheet.Range["E1:"+"E"+(RowRange + 2).ToString()].ColumnWidth = 25;
                //Merge cells
                //worksheet.Range["D1:E1"].Merge();

                ////Enter text to the cell D1 and apply formatting.
                //worksheet.Range["D1"].Text = "INVOICE";
                //worksheet.Range["D1"].CellStyle.Font.Bold = true;
                //worksheet.Range["D1"].CellStyle.Font.RGBColor = Color.FromArgb(42, 118, 189);
                //worksheet.Range["D1"].CellStyle.Font.Size = 35;

                ////Apply alignment in the cell D1
                //worksheet.Range["D1"].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignRight;
                //worksheet.Range["D1"].CellStyle.VerticalAlignment = ExcelVAlign.VAlignTop;

                ////Enter values to the cells from D5 to E8
                //worksheet.Range["D5"].Text = "INVOICE#";
                //worksheet.Range["E5"].Text = "DATE";
                //worksheet.Range["D6"].Number = 1028;
                //worksheet.Range["E6"].Value = "12/31/2018";
                //worksheet.Range["D7"].Text = "CUSTOMER ID";
                //worksheet.Range["E7"].Text = "TERMS";
                //worksheet.Range["D8"].Number = 564;
                //worksheet.Range["E8"].Text = "Due Upon Receipt";

                ////Apply RGB backcolor to the cells from D5 to E8
                //worksheet.Range["D5:E5"].CellStyle.Color = Color.FromArgb(42, 118, 189);
                //worksheet.Range["D7:E7"].CellStyle.Color = Color.FromArgb(42, 118, 189);

                ////Apply known colors to the text in cells D5 to E8
                //worksheet.Range["D5:E5"].CellStyle.Font.Color = ExcelKnownColors.White;
                //worksheet.Range["D7:E7"].CellStyle.Font.Color = ExcelKnownColors.White;

                ////Make the text as bold from D5 to E8
                //worksheet.Range["D5:E8"].CellStyle.Font.Bold = true;

                ////Apply alignment to the cells from D5 to E8
                //worksheet.Range["D5:E8"].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
                //worksheet.Range["D5:E5"].CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter;
                //worksheet.Range["D7:E7"].CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter;
                //worksheet.Range["D6:E6"].CellStyle.VerticalAlignment = ExcelVAlign.VAlignTop;

                ////Enter value and applying formatting in the cell A7
                //worksheet.Range["A7"].Text = "  BILL TO";
                //worksheet.Range["A7"].CellStyle.Color = Color.FromArgb(42, 118, 189);
                //worksheet.Range["A7"].CellStyle.Font.Bold = true;
                //worksheet.Range["A7"].CellStyle.Font.Color = ExcelKnownColors.White;

                ////Apply alignment
                //worksheet.Range["A7"].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignLeft;
                //worksheet.Range["A7"].CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter;

                ////Enter values in the cells A8 to A12
                //worksheet.Range["A8"].Text = "Steyn";
                //worksheet.Range["A9"].Text = "Great Lakes Food Market";
                //worksheet.Range["A10"].Text = "20 Whitehall Rd";
                //worksheet.Range["A11"].Text = "North Muskegon,USA";
                //worksheet.Range["A12"].Text = "+1 231-654-0000";

                ////Create a Hyperlink for e-mail in the cell A13
                //IHyperLink hyperlink = worksheet.HyperLinks.Add(worksheet.Range["A13"]);
                //hyperlink.Type = ExcelHyperLinkType.Url;
                //hyperlink.Address = "Steyn@greatlakes.com";
                //hyperlink.ScreenTip = "Send Mail";

                ////Merge column A and B from row 15 to 22
                //worksheet.Range["A15:B15"].Merge();
                //worksheet.Range["A16:B16"].Merge();
                //worksheet.Range["A17:B17"].Merge();
                //worksheet.Range["A18:B18"].Merge();
                //worksheet.Range["A19:B19"].Merge();
                //worksheet.Range["A20:B20"].Merge();
                //worksheet.Range["A21:B21"].Merge();
                //worksheet.Range["A22:B22"].Merge();

                ////Enter details of products and prices
                //worksheet.Range["A15"].Text = "  DESCRIPTION";
                //worksheet.Range["C15"].Text = "QTY";
                //worksheet.Range["D15"].Text = "UNIT PRICE";
                //worksheet.Range["E15"].Text = "AMOUNT";
                //worksheet.Range["A16"].Text = "Cabrales Cheese";
                //worksheet.Range["A17"].Text = "Chocos";
                //worksheet.Range["A18"].Text = "Pasta";
                //worksheet.Range["A19"].Text = "Cereals";
                //worksheet.Range["A20"].Text = "Ice Cream";
                //worksheet.Range["C16"].Number = 3;
                //worksheet.Range["C17"].Number = 2;
                //worksheet.Range["C18"].Number = 1;
                //worksheet.Range["C19"].Number = 4;
                //worksheet.Range["C20"].Number = 3;
                //worksheet.Range["D16"].Number = 21;
                //worksheet.Range["D17"].Number = 54;
                //worksheet.Range["D18"].Number = 10;
                //worksheet.Range["D19"].Number = 20;
                //worksheet.Range["D20"].Number = 30;
                //worksheet.Range["D23"].Text = "Total";

                ////Apply number format
                //worksheet.Range["D16:E22"].NumberFormat = "$.00";
                //worksheet.Range["E23"].NumberFormat = "$.00";

                ////Apply incremental formula for column Amount by multiplying Qty and UnitPrice
                //application.EnableIncrementalFormula = true;
                //worksheet.Range["E16:E20"].Formula = "=C16*D16";

                ////Formula for Sum the total
                //worksheet.Range["E23"].Formula = "=SUM(E16:E22)";

                ////Apply borders
                //worksheet.Range["A16:E22"].CellStyle.Borders[ExcelBordersIndex.EdgeTop].LineStyle = ExcelLineStyle.Thin;
                //worksheet.Range["A16:E22"].CellStyle.Borders[ExcelBordersIndex.EdgeBottom].LineStyle = ExcelLineStyle.Thin;
                //worksheet.Range["A16:E22"].CellStyle.Borders[ExcelBordersIndex.EdgeTop].Color = ExcelKnownColors.Grey_25_percent;
                //worksheet.Range["A16:E22"].CellStyle.Borders[ExcelBordersIndex.EdgeBottom].Color = ExcelKnownColors.Grey_25_percent;
                //worksheet.Range["A23:E23"].CellStyle.Borders[ExcelBordersIndex.EdgeTop].LineStyle = ExcelLineStyle.Thin;
                //worksheet.Range["A23:E23"].CellStyle.Borders[ExcelBordersIndex.EdgeBottom].LineStyle = ExcelLineStyle.Thin;
                //worksheet.Range["A23:E23"].CellStyle.Borders[ExcelBordersIndex.EdgeTop].Color = ExcelKnownColors.Black;
                //worksheet.Range["A23:E23"].CellStyle.Borders[ExcelBordersIndex.EdgeBottom].Color = ExcelKnownColors.Black;

                ////Apply font setting for cells with product details
                //worksheet.Range["A3:E23"].CellStyle.Font.FontName = "Arial";
                //worksheet.Range["A3:E23"].CellStyle.Font.Size = 10;
                //worksheet.Range["A15:E15"].CellStyle.Font.Color = ExcelKnownColors.White;
                //worksheet.Range["A15:E15"].CellStyle.Font.Bold = true;
                //worksheet.Range["D23:E23"].CellStyle.Font.Bold = true;

                ////Apply cell color
                //worksheet.Range["A15:E15"].CellStyle.Color = Color.FromArgb(42, 118, 189);

                ////Apply alignment to cells with product details
                //worksheet.Range["A15"].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignLeft;
                //worksheet.Range["C15:C22"].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
                //worksheet.Range["D15:E15"].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;

                ////Apply row height and column width to look good
                //worksheet.Range["A1"].ColumnWidth = 36;
                //worksheet.Range["B1"].ColumnWidth = 11;
                //worksheet.Range["C1"].ColumnWidth = 8;
                //worksheet.Range["D1:E1"].ColumnWidth = 18;
                //worksheet.Range["A1"].RowHeight = 47;
                //worksheet.Range["A2"].RowHeight = 15;
                //worksheet.Range["A3:A4"].RowHeight = 15;
                //worksheet.Range["A5"].RowHeight = 18;
                //worksheet.Range["A6"].RowHeight = 29;
                //worksheet.Range["A7"].RowHeight = 18;
                //worksheet.Range["A8"].RowHeight = 15;
                //worksheet.Range["A9:A14"].RowHeight = 15;
                //worksheet.Range["A15:A23"].RowHeight = 18;
                string path = Server.MapPath("~/Reports/");
                //Save the workbook to disk in xlsx format
                //workbook.SaveAs("Output.xlsx", HttpContext.ApplicationInstance.Response, ExcelDownloadType.Open);
                workbook.SaveAs(path + Path.GetFileName(name));
                var bull = db.Downloads.SingleOrDefault(b => b.GuidName == name);


                if (bull != null)
                {
                    bull.IsExist = true;
                    db.SaveChanges();
                }
                /*
                  if (postedFile != null)
            {
                string path = Server.MapPath("~/Uploads/");
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }

                postedFile.SaveAs(path + Path.GetFileName(postedFile.FileName));
                ViewBag.Message = "File uploaded successfully.";
            }    
                 */

            }


        }


        public ActionResult Listele(DateTime? dates, DateTime? datee, string submit)
        {

            var model = db.Transections.ToList();

            DateTime startdate = dates ?? new DateTime(2000, 10, 10, 1, 1, 1, 1);
            DateTime enddate = datee ?? DateTime.Now; ;
            foreach (var item in model.ToList())
            {
                if (item.Date < startdate || item.Date > enddate)
                {
                    model.Remove(item);

                }
            }

            TestEEP();// Buna devam edilecek
            if (submit == "report")
            {
                ExportToExcel(dates, datee);
                return RedirectToAction("Listele", "Home");


                //CancellationToken a = new CancellationToken(false);
                //return RedirectToAction( "ExportToExcel", "Sec");
                //ExportToExcelasync(dates, datee);

                //ThreadStart childthreat = delegate { ExportToExcelasync(dates, datee,a); };  //new ThreadStart(childthreadcall(3));

                //Thread child = new Thread(childthreat);

                //child.Start();
                //Debug.WriteLine("GİRDİ");

                //child.Abort();

            }

            return View(model);

        }

        public void childthreadcall(DateTime? a, DateTime? b)
        {//kullanılmadı
            string text = "";
            try
            {
                text = "<br />Child thread started <br/>";
               
                text += "Child Thread: Coiunting to 10";
             
                for (int i = 0; i < 10; i++)
                {
                    Thread.Sleep(500);
                    text += "<br/> in Child thread </br>";
              
                }

                text += "<br/> child thread finished";
              

            }
            catch (ThreadAbortException e)
            {

                text += "<br /> child thread - exception";
              

            }
            finally
            {
                text += "<br /> child thread - unable to catch the  exception";
          
            }
        }


        public ActionResult DownloadListele()
        {
            var model = db.Downloads.ToList();
            return View(model);
        }


        public async Task ExportToExcelasync(DateTime? start, DateTime? end, CancellationToken cancellationToken)
        {//kullanılmadı
            await Task.Delay(1_000, cancellationToken);
            await Task.Run(() => ExportToExcel(start, end));
        }

        public void ExportToExcel(DateTime? start, DateTime? end)
        {


            try
            {//Downloads tablosuna indirmenin kaydı

                Download m = new Download();

                DateTime startdate = start ?? new DateTime(2000, 10, 10, 1, 1, 1, 1); ;

                DateTime enddate = end ?? DateTime.Now;
                DateTime now = DateTime.Now;

                m.IsExist = false;
                m.CreateDate = now;
                m.EndDate = now;
                m.StartDate = now;
                string name = "Report";
                string date = now.ToString("F");
                date = date.Replace(" ", "_");
                date = date.Replace(",", "_");
                string sonu = ".xls";
                date += sonu;
                name += date;
                m.GuidName = name;
                db.Downloads.Add(m);
                db.SaveChanges();

                //Creating Excel
                var gv = new GridView();

                var temp = db.Transections.ToList();

                foreach (var item in temp.ToList())
                {
                    if (item.Date < startdate || item.Date > enddate)
                    {
                        temp.Remove(item);

                    }
                }
         

                gv.DataSource = temp;
                gv.DataBind();
                Response.ClearContent();
                Response.Buffer = true;
                string attachment = "attachment; filename=";
                attachment += name;
                //Response.ClientDisconnectedToken = false;
                Response.AddHeader("content-disposition", attachment);
                Response.ContentType = "application/ms-excel";
                Response.Charset = "";
                StringWriter objStringWriter = new StringWriter();
                HtmlTextWriter objHtmlTextWriter = new HtmlTextWriter(objStringWriter);
                gv.RenderControl(objHtmlTextWriter);
                Response.Output.Write(objStringWriter.ToString());
                Response.Flush();
                Response.End();
                //IsExist set True
                var bull = db.Downloads.SingleOrDefault(b => b.GuidName == name);
                if (bull != null)
                {
                    bull.IsExist = true;
                    db.SaveChanges();
                }
            }
            catch (Exception ex)
            {
               
            }
        }




    }

    public class SecController : Controller
    {

        DatabaseContext db = new DatabaseContext();
        public ActionResult Index()
        {
           
            return View();
        }
        public ActionResult Deneme(DateTime? dates, DateTime? datee)
        {
            return new EmptyResult();
        }

        public ActionResult ExportToExcel(DateTime? start, DateTime? end)
        {

            try
            {//Downloads tablosuna indirmenin kaydı

                Download m = new Download();

                DateTime startdate = start ?? new DateTime(2000, 10, 10, 1, 1, 1, 1); ;

                DateTime enddate = end ?? DateTime.Now;
                DateTime now = DateTime.Now;
                m.IsExist = false;
                m.CreateDate = now;
                m.EndDate = now;
                m.StartDate = now;
                string name = "Report";
                string date = now.ToString("F");
                date = date.Replace(" ", "_");
                date = date.Replace(",", "_");
                string sonu = ".xls";
                date += sonu;
                name += date;
                m.GuidName = name;
                db.Downloads.Add(m);
                db.SaveChanges();
           

                //Creating Excel
                var gv = new GridView();

                var temp = db.Transections.ToList();

                foreach (var item in temp.ToList())
                {
                    if (item.Date < startdate || item.Date > enddate)
                    {
                        temp.Remove(item);

                    }
                }

                gv.DataSource = temp;
                gv.DataBind();
                Response.ClearContent();
                Response.Buffer = true;
                string attachment = "attachment; filename=";
                attachment = attachment + name;
                Response.AddHeader("content-disposition", attachment);
                Response.ContentType = "application/ms-excel";
                Response.Charset = "";
                StringWriter objStringWriter = new StringWriter();
                HtmlTextWriter objHtmlTextWriter = new HtmlTextWriter(objStringWriter);
                gv.RenderControl(objHtmlTextWriter);
                Response.Output.Write(objStringWriter.ToString());
                Response.Flush();
                Response.End();
                //IsExist set True
                var bull = db.Downloads.SingleOrDefault(b => b.GuidName == name);

            
                if (bull != null)
                {
                    bull.IsExist = true;
                    db.SaveChanges();
                }

                return new EmptyResult();
            }
            catch (Exception ex)
            {
               
                return new EmptyResult();
            }

        }
    }


}




