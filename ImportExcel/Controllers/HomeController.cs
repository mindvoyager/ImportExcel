using ImportExcel.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.Entity.Core.EntityClient;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace ImportExcel.Controllers
{
    public class HomeController : Controller
    {
        // GET: Home
        public ActionResult Index()
        {
            return View();
        }
        // POST: Home
        [HttpPost]
        public ActionResult Index(ReadExcel readExcel)
        {
            if (ModelState.IsValid)
            {
                //string path = Server.MapPath("~/Content/Upload" + readExcel.File.FileName);
                string folder = Server.MapPath("~/Content/Upload");
                string file = Path.GetFileName(readExcel.File.FileName);
                string path = Path.Combine(folder, file);
                readExcel.File.SaveAs(path);

                string excelConnectionString = @"Provider='Microsoft.ACE.OLEDB.12.0';Data Source='" + path + "';Extended Properties='Excel 12.0 Xml;IMEX=1'";
                OleDbConnection excelConnection = new OleDbConnection(excelConnectionString);

                //Sheet Name
                excelConnection.Open();
                string tableName = excelConnection.GetSchema("Tables").Rows[0]["TABLE_NAME"].ToString();
                excelConnection.Close();
                //End

                OleDbCommand cmd = new OleDbCommand("SELECT * FROM [" + tableName + "]", excelConnection);

                excelConnection.Open();

                OleDbDataReader dbReader = cmd.ExecuteReader();

                EntityConnectionStringBuilder entityConnectionStringBuilder = new EntityConnectionStringBuilder(ConfigurationManager.ConnectionStrings["SampleDatabaseEntities"].ConnectionString);
                string connectionString = entityConnectionStringBuilder.ProviderConnectionString;

                SqlBulkCopy sqlBulk = new SqlBulkCopy(connectionString);

                //Give destination table name
                sqlBulk.DestinationTableName = "Sale";

                //Mappings
                sqlBulk.ColumnMappings.Add("Date", "AddedOn");
                sqlBulk.ColumnMappings.Add("Region", "Region");
                sqlBulk.ColumnMappings.Add("Person", "Person");
                sqlBulk.ColumnMappings.Add("Item", "Item");
                sqlBulk.ColumnMappings.Add("Units", "Units");
                sqlBulk.ColumnMappings.Add("Unit Cost", "UnitCost");
                sqlBulk.ColumnMappings.Add("Total", "Total");

                sqlBulk.WriteToServer(dbReader);
                excelConnection.Close();

                ViewBag.Result = "Successfully Imported";
            }
            return View();
        }
    }
}