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

       

        public ActionResult Index()
        {
            var model1 = db.Transections.ToList();
            var model2 = db.Downloads.ToList();

            return RedirectToAction("Listele","Home");
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
                 ExportToExcel(dates, datee);
                //ExportToExcelasync(dates, datee);
            }

            return View(model);

        }


        public ActionResult DownloadListele()
        {
            var model = db.Downloads.ToList();
            return View(model);
        }
       

        public async Task ExportToExcelasync(DateTime? start, DateTime? end)
        {
           
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




