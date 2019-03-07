using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Excel = Microsoft.Office.Interop.Excel;

namespace WebApplicationDemo
{
    public partial class _Default : Page
    {
        #region [GLOBAL DECLARATION]
        public static DataSet ds;

        #endregion

        #region [PAGE LOAD]
        protected void Page_Load(object sender, EventArgs e)
        {

            if (!IsPostBack)
            {
                ds = new DataSet();
                DataTable _dataTable = new DataTable(); // create new instance of data table.
                // add columns to that data table.
                _dataTable.Columns.Add("Id", typeof(int));
                _dataTable.Columns.Add("Name", typeof(string));
                _dataTable.Columns.Add("Email", typeof(string));
                // add data rows to the data table.
                _dataTable.Rows.Add(12, "Pritey Kapoor", "pritey@user.com");
                _dataTable.Rows.Add(24, "Juhi Kapoor", "Juhi@user.com");
                _dataTable.Rows.Add(54, "Janki Kapoor", "janki@user.com");
                ds.Tables.Add(_dataTable);
            }
        }

        #endregion

        #region [EVENTS]
        protected void btnExport_Click(object sender, EventArgs e)
        {
            try
            {
                // set export excel file path
                string excelFilePath = Server.MapPath("/Content/Data/") + "ExportExcel_" + DateTime.Now.ToString("ddMMYYYY_hhmmss") + ".xls";
                // create excel file and save in the folder               
                var app = new Microsoft.Office.Interop.Excel.Application();
                var wb = app.Workbooks.Add();
                wb.SaveAs(excelFilePath);
                wb.Close();
                //Create an Excel application instance
                Excel.Application excelApp = new Excel.Application();
                string projectDirectory = Directory.GetParent(Environment.CurrentDirectory).Parent.FullName;
                //Create an Excel workbook instance and open it from the predefined location
                Excel.Workbook excelWorkBook = excelApp.Workbooks.Open(excelFilePath);

                foreach (DataTable table in ds.Tables)
                {
                    //Add a new worksheet to workbook with the Datatable name
                    Excel.Worksheet excelWorkSheet = excelWorkBook.Sheets.Add();
                    excelWorkSheet.Name = table.TableName;

                    for (int i = 1; i < table.Columns.Count + 1; i++)
                    {
                        excelWorkSheet.Cells[1, i] = table.Columns[i - 1].ColumnName;
                    }

                    for (int j = 0; j < table.Rows.Count; j++)
                    {
                        for (int k = 0; k < table.Columns.Count; k++)
                        {
                            excelWorkSheet.Cells[j + 2, k + 1] = table.Rows[j].ItemArray[k].ToString();
                        }
                    }
                }

                excelWorkBook.Save();
                excelWorkBook.Close();
                excelApp.Quit();
            }
            catch (Exception ex)
            {
                Response.Write(ex.Message);
            }
        }

        #endregion 
    }
}