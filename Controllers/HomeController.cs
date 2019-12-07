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

namespace TestNagis.Controllers
{
    public class HomeController : Controller
    {

        DatabaseContext db = new DatabaseContext();

        public ActionResult Gel()
        {
            return Content("GelGel");
        }

        public ActionResult Index()
        {
            var model1 = db.Transections.ToList();
            var model2 = db.Downloads.ToList();

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

            if (submit == "report")
            {
                // DataSet dataSet= GetRecordsFromDatabase();
                //RedirectToAction("ExportToExcel", "Grid", new { @start = dates, @end = datee });
                ExportToExcel(dates, datee);
                //AsyncBackground();
            //DataSet dataSet = dbtest();
                //test();
                //Debug.WriteLine("Asenkron BİTTİ");
                //AsyncBackground();


            }

            return View(model);

        }


        public ActionResult DownloadListele()
        {
            var model = db.Downloads.ToList();
            return View(model);
        }
        public void test()
        {
            int k = 0;
            for (int i = 0; i < 10000; i++)
            {
                k += i;
                Debug.WriteLine(i);

            }
            Debug.WriteLine(k);
            Debug.WriteLine("test calıstı");

        }

        public void Asynctest()
        {
            Debug.WriteLine("Asynctest çalıştı");
            System.Threading.Tasks.Task.Run(() => test());
        }


        DataSet dbtest()
        {
            DataSet dataSet = new DataSet();
            string connetionString;
            SqlConnection cnn;
            // connetionString = @"Data Source=WIN-50GP30FGO75;Initial Catalog=Demodb;User ID=sa;Password=demol23";
            connetionString = "Server=DESKTOP-29NSEN8\\SQLEXPRESS;Database=TestNagisDB;Integrated Security=True";
            cnn = new SqlConnection(connetionString);
            cnn.Open();
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "Select * FROM Downloads";
            cmd.Connection = cnn;
            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter();
            sqlDataAdapter.SelectCommand = cmd;
            sqlDataAdapter.Fill(dataSet);
            cnn.Close();
            Debug.WriteLine(dataSet.ToString());
            return dataSet;
        }

        public void AsyncBackground()
        {
            Debug.WriteLine("AsyncBackground çalıştı");
            System.Threading.Tasks.Task.Run(() => ExcelCreateBackground());
        }

        [ActionName("Demo")]
        public ActionResult ExcelCreateBackground()
        {

            for (int i = 0; i < 1000; i++)
            {

                Debug.WriteLine(i);

            }
            //Fill dataset with records
            DataSet dataSet = dbtest();

            StringBuilder sb = new StringBuilder();

            sb.Append("<table>");

            //LINQ to get Column names
            var columnName = dataSet.Tables[0].Columns.Cast<DataColumn>()
                                 .Select(x => x.ColumnName)
                                 .ToArray();
            sb.Append("<tr>");
            //Looping through the column names
            foreach (var col in columnName)
                sb.Append("<td>" + col + "</td>");
            sb.Append("</tr>");

            //Looping through the records
            foreach (DataRow dr in dataSet.Tables[0].Rows)
            {
                sb.Append("<tr>");
                foreach (DataColumn dc in dataSet.Tables[0].Columns)
                {
                    sb.Append("<td>" + dr[dc] + "</td>");
                }
                sb.Append("</tr>");
            }

            sb.Append("</table>");

            //Writing StringBuilder content to an excel file.
            Response.Clear();
            Response.ClearContent();
            Response.ClearHeaders();
            Response.Charset = "";
            Response.Buffer = true;
            Response.ContentType = "application/vnd.ms-excel";
            Response.AddHeader("content-disposition", "attachment;filename=UserReport.xls");
            Response.Write(sb.ToString());
            Response.Flush();
            Response.Close();

            return RedirectToAction("Listele");
        }

        DataSet GetRecordsFromDatabase()
        {
            //string constr = ConfigurationManager.ConnectionStrings["DefaultConnection"].ToString();
            //conn = new SqlConnection(constr);
            //SqlConnection conn = null;

            string connetionString;
            SqlConnection cnn;
            //Server=DESKTOP-29NSEN8\SQLEXPRESS;Database=TestNagisDB;Integrated Security=True
            connetionString = "Server=DESKTOP-29NSEN8\\SQLEXPRESS;Database=TestNagisDB;Integrated Security=True";
            cnn = new SqlConnection(connetionString);
            cnn.Open();
            Debug.WriteLine("con acıldı ");
            cnn.Close();
            DataSet dataSet = new DataSet();
            /*
            conn.ConnectionString = ConfigurationManager.ConnectionStrings["CS"].ConnectionString;
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "Select * FROM Downloads";
            cmd.Connection = conn;

            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter();
            sqlDataAdapter.SelectCommand = cmd;
            sqlDataAdapter.Fill(dataSet);
            */

            return dataSet;
        }



        public async Task ExportToExcelasync(DateTime? start, DateTime? end)
        {
            Debug.WriteLine("async geldi");
            Debug.WriteLine(start);
            Debug.WriteLine(end);
            await Task.Run(() => ExportToExcel(start, end));
        }

        public void ExportToExcel(DateTime? start, DateTime? end)
        {
            Debug.WriteLine("Export To Excel çaılıştı");

            Debug.WriteLine(start);
            Debug.WriteLine(end);

            for (int i = 0; i < 10011; i++)
            {
                Debug.Write(i);
            }




            try
            {//Downloads tablosuna indirmenin kaydı


                Download m = new Download();



                DateTime startdate = start ?? new DateTime(2000, 10, 10, 1, 1, 1, 1); ;

                DateTime enddate = end ?? DateTime.Now;

                Debug.WriteLine(startdate);
                Debug.WriteLine(enddate);
                //each caseler eklenecek seçmezse geçersiz seçerse end date start date den buyukse startla end date degıstırme eklencek




                DateTime now = DateTime.Now;
                //string trim = text.Replace( " ", "_" );
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

                //Excelin hazırlanması

                var gv = new GridView();

                var temp = db.Transections.ToList();

                // start 2010  ///// end 2015
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
                //attachment = "\"" + attachment +"\"";
                Response.AddHeader("content-disposition", attachment);
                Response.ContentType = "application/ms-excel";
                Response.Charset = "";
                StringWriter objStringWriter = new StringWriter();
                HtmlTextWriter objHtmlTextWriter = new HtmlTextWriter(objStringWriter);
                gv.RenderControl(objHtmlTextWriter);
                Response.Output.Write(objStringWriter.ToString());
                Response.Flush();
                Response.End();
                /*
                     var author = context.Authors.First(a => a.AuthorId == 1);
    author.FirstName = "Bill";
    context.SaveChanges();
    */
                var bull  = db.Downloads.SingleOrDefault(b => b.GuidName == name);
                if (bull != null)
                {
                    bull.IsExist = true;
                    db.SaveChanges();
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }
        }

    }
}




