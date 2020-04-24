using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.IO;
using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;

namespace Epplus_Export
{
    class DAL
    {
        SqlCommand cmd1 = null;
        SqlDataReader dataReader;
        private static IKDE_Email emailer = new IKDE_Email();
        public static AppSettingsReader apReader = new AppSettingsReader();
        SqlConnection con = null;

        string QueryString = string.Empty;
        string EntityName = "KWS";
        string fileName = string.Empty;
        string conString = ConfigurationManager.ConnectionStrings["IBSConnectionString"].ConnectionString;
        string toEmailAddress = string.Empty;
        string root = string.Empty;
        string path = string.Empty;

        string QueryStringDepotGiud = string.Empty;

        public Guid Depot = Guid.Empty;
        ExcelPackage excel = new ExcelPackage();
        //int rowCnt = 1;

        void formatSheetColumns(DataTable table, ExcelWorksheet worksheet, ExcelWorksheet worksheetTrans, int rowCnt)
        {
            int count = table.Rows.Count;
            int countCol = table.Columns.Count;
            bool isHeader = true;
            int num = 0;
            string value = string.Empty;

            worksheet.Cells["A1"].Style.Font.Bold = true;
            worksheet.Cells["A1"].Style.Font.Size = 18;

            worksheet.Cells["A1"].Value = "Organisation statement";
            worksheet.Cells["A3"].Value = "Welcome to the organisation statement report.";
            worksheet.Cells["A4"].Value = "Please navigate to the organisation statement sheet to view the report.";
            worksheet.Cells["A6"].Value = "Report parameters";
            worksheet.Cells["A7"].Value = "Organisation";
            worksheet.Cells["A8"].Value = "Account financed by organisation";
            worksheet.Cells["A9"].Value = "Organisation balance";
            worksheet.Cells["A10"].Value = "Vehicle";
            worksheet.Cells["A11"].Value = "Driver";
            worksheet.Cells["A12"].Value = "Depot";
            worksheet.Cells["A13"].Value = "Captured date start";
            worksheet.Cells["A14"].Value = "Captured date end";
            worksheet.Cells["A15"].Value = "Include child accounts";
            worksheet.Cells["A16"].Value = "Plaza";

            worksheet.Cells["B7"].Value = "KWS Carriers (Pty) Ltd(8BAU01)";
            worksheet.Cells["B8"].Value = "Yes";
            worksheet.Cells["B9"].Value = "[Not filtered]";
            worksheet.Cells["B10"].Value = "[Not filtered]";
            worksheet.Cells["B11"].Value = "[Not filtered]";
            worksheet.Cells["B12"].Value = "[Not filtered]";
            worksheet.Cells["B13"].Value = DateTime.Now.ToString();
            worksheet.Cells["B14"].Value = DateTime.Now.ToString();
            worksheet.Cells["B15"].Value = "No";
            worksheet.Cells["B16"].Value = "[Not filtered]";



            worksheetTrans.Cells["A1"].Value = "Customer";
            worksheetTrans.Cells["A2"].Value = "KWS Carriers (Pty) Ltd(8BAU01)";
            worksheetTrans.Cells.AutoFitColumns();

            for (int i = 0; i < count; i++)
            {
                // Set Border
                for (int j = 1; j <= countCol; j++)
                {
                    // Set Border For Header. Run once only
                    if (isHeader) worksheetTrans.Cells[1, j].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    worksheetTrans.Cells[rowCnt, j].Style.Font.Bold = true;
                    worksheetTrans.Cells[rowCnt, j].Style.Font.Size = 10;
                    worksheetTrans.Cells[rowCnt, j].Style.Font.Color.SetColor(Color.Navy);
                    worksheetTrans.Cells[rowCnt, j].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheetTrans.Cells[rowCnt, j].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                    worksheetTrans.Cells[rowCnt, j].Style.Font.Name = "Tohoma";
                }

                // Set to false
                isHeader = false;

                value = worksheetTrans.Cells[2 + i, 2].Value + ""; // You can use .ToString() if u sure is not null

                // Parse
                int.TryParse(value, out num);

                // Check
                if (num >= 1 && num <= 11)
                    worksheetTrans.Cells[2 + i, 2].Style.Font.Color.SetColor(System.Drawing.Color.Green);
                else if (num >= 12 && num <= 39)
                    worksheetTrans.Cells[2 + i, 2].Style.Font.Color.SetColor(System.Drawing.Color.Orange);
            }
        }


        public DataSet RetriveTransactions()
        {
            int rowCnt = 1;
            int rowCnt2 = 1;

            string ltrTot, ltrTot2, amnt, amnt2, vatAmnt, vatAmnt2, chldCust, chldCust2, rebate, rebate2;

            DataSet ds = new DataSet();
            using (con = new SqlConnection(this.conString))
            {
                try
                {
                    con.Open();
                    fileName = (string)apReader.GetValue("FileName", typeof(string)) + "_" + DateTime.Now.ToString("yyyyMMdd") + ".xls";

                    int counter = 0;

                    counter += 1;

                    QueryString = "sp_KWS_Transactions";

                    SqlCommand cmd = new SqlCommand(QueryString, con);
                    cmd.CommandType = CommandType.StoredProcedure;
                    // cmd.Parameters.AddWithValue("depoGuid", Depot);

                    SqlDataAdapter dscmd = new SqlDataAdapter(cmd);
                    DataSet ds1 = new DataSet();
                    dscmd.Fill(ds1);

                    foreach (DataTable table in ds1.Tables)
                    {
                        using (ExcelPackage excel = new ExcelPackage())
                        {
                            excel.Workbook.Worksheets.Add("Introduction");
                            excel.Workbook.Worksheets.Add("Organisation statement");
                            //excel.Workbook.Worksheets.Add("Purchase Orders");
                            // Target a worksheet
                            var worksheet = excel.Workbook.Worksheets["Introduction"];
                            var worksheetTransSalesOrder = excel.Workbook.Worksheets["Organisation statement"];

                            worksheetTransSalesOrder.Cells[rowCnt, 2].LoadFromDataTable(table, true);
                            formatSheetColumns(table, worksheet, worksheetTransSalesOrder, rowCnt);
                            rowCnt = worksheetTransSalesOrder.Dimension.End.Row;
                            rowCnt2 = worksheetTransSalesOrder.Dimension.End.Row;
                            rowCnt += 1;

                            string rCount = Convert.ToString(rowCnt);

                            worksheetTransSalesOrder.Cells["O1:O" + rCount].Style.Numberformat.Format = "#,##0.00";
                            worksheetTransSalesOrder.Cells["R1:R" + rCount].Style.Numberformat.Format = "R#,##0.00";
                            worksheetTransSalesOrder.Cells["S1:S" + rCount].Style.Numberformat.Format = "R#,##0.00";
                            worksheetTransSalesOrder.Cells["T1:T" + rCount].Style.Numberformat.Format = "R#,##0.00";
                            worksheetTransSalesOrder.Cells["U1:U" + rCount].Style.Numberformat.Format = "R#,##0.00";
                            worksheetTransSalesOrder.Cells["AA1:AA" + rCount].Style.Numberformat.Format = "R#,##0.00";

                            ltrTot = 'O' + rowCnt.ToString();
                            ltrTot2 = 'O' + rowCnt2.ToString();

                            amnt = 'R' + rowCnt.ToString();
                            amnt2 = 'R' + rowCnt2.ToString();

                            vatAmnt = 'S' + rowCnt.ToString();
                            vatAmnt2 = 'S' + rowCnt2.ToString();

                            chldCust = 'T' + rowCnt.ToString();
                            chldCust2 = 'T' + rowCnt2.ToString();

                            rebate = 'U' + rowCnt.ToString();
                            rebate2 = 'U' + rowCnt2.ToString();

                            worksheetTransSalesOrder.Cells[ltrTot].Formula = "=SUM(O2:" + ltrTot2 + ")";
                            worksheetTransSalesOrder.Cells[amnt].Formula = "=SUM(R2:" + amnt2 + ")";
                            worksheetTransSalesOrder.Cells[vatAmnt].Formula = "=SUM(S2:" + vatAmnt2 + ")";
                            worksheetTransSalesOrder.Cells[chldCust].Formula = "=SUM(T2:" + chldCust2 + ")";
                            worksheetTransSalesOrder.Cells[rebate].Formula = "=SUM(U2:" + rebate2 + ")";

                            worksheetTransSalesOrder.Cells["O1:O" + rCount].AutoFitColumns();

                            worksheetTransSalesOrder.Cells[ltrTot].Style.Border.Top.Style = ExcelBorderStyle.Double;
                            worksheetTransSalesOrder.Cells[amnt].Style.Border.Top.Style = ExcelBorderStyle.Double;
                            worksheetTransSalesOrder.Cells[vatAmnt].Style.Border.Top.Style = ExcelBorderStyle.Double;
                            worksheetTransSalesOrder.Cells[chldCust].Style.Border.Top.Style = ExcelBorderStyle.Double;
                            worksheetTransSalesOrder.Cells[rebate].Style.Border.Top.Style = ExcelBorderStyle.Double;


                            worksheetTransSalesOrder.Cells[ltrTot].Style.Font.Color.SetColor(Color.Navy);
                            worksheetTransSalesOrder.Cells[ltrTot].Style.Font.Bold = true;
                            worksheetTransSalesOrder.Cells[ltrTot].Style.Font.Size = 12;

                            worksheetTransSalesOrder.Cells[amnt].Style.Font.Color.SetColor(Color.Navy);
                            worksheetTransSalesOrder.Cells[amnt].Style.Font.Bold = true;
                            worksheetTransSalesOrder.Cells[amnt].Style.Font.Size = 12;

                            worksheetTransSalesOrder.Cells[vatAmnt].Style.Font.Color.SetColor(Color.Navy);
                            worksheetTransSalesOrder.Cells[vatAmnt].Style.Font.Bold = true;
                            worksheetTransSalesOrder.Cells[vatAmnt].Style.Font.Size = 12;

                            worksheetTransSalesOrder.Cells[chldCust].Style.Font.Color.SetColor(Color.Navy);
                            worksheetTransSalesOrder.Cells[chldCust].Style.Font.Bold = true;
                            worksheetTransSalesOrder.Cells[chldCust].Style.Font.Size = 12;

                            worksheetTransSalesOrder.Cells[rebate].Style.Font.Color.SetColor(Color.Navy);
                            worksheetTransSalesOrder.Cells[rebate].Style.Font.Bold = true;
                            worksheetTransSalesOrder.Cells[rebate].Style.Font.Size = 12;

                            root = @"C:\IBS_Dev\IBS_KWS_DailyTransaction_Extract_Console\" + EntityName + "";

                            // If directory does not exist, create it. 
                            if (!Directory.Exists(root))
                            {
                                Directory.CreateDirectory(root);
                            }

                            path = root + @"\" + fileName;

                            FileInfo excelFile = new FileInfo(path);
                            excel.SaveAs(excelFile);
                        }
                    }

                    con.Close();
                    //Get All billed transaction for the previos business day 
                    //End

                    //Send email Start
                    emailer.SendNotifyMailUser(path, toEmailAddress);

                    //con.Open();
                    //QueryString = "sp_Update_Billed_Transactions";

                    //SqlCommand cmd1 = new SqlCommand(QueryString, con);
                    //cmd1.CommandType = CommandType.StoredProcedure;
                    //cmd1.ExecuteNonQuery();

                    //con.Close();
                    ////Flag all sent Records. To avoid Duplicates


                }
                catch (Exception ex)
                {
                    emailer.SendNotifyErrorMail("Error", ex.Message);
                }
            }
            return ds;
        }
    }
}
