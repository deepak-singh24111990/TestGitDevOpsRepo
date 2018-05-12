using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

using System.Data;
using System.IO;
using System.Net;
using DocumentFormat.OpenXml.Packaging;
using System.Collections.Concurrent;
using System.Data.SqlClient;


namespace BasePointDownload
{
    class Program
    {

        public void FetchData()
        {

            SqlConnection sqcon = new SqlConnection(@"Data Source=119.81.116.156;Initial Catalog=D_RedClub;User ID=DSDlocal;Password=Hy@tt%43;Connect Timeout=1000000;");
            //SqlConnection sqcon = new SqlConnection(@"Data Source=119.81.116.148;Initial Catalog=D_RedClubNew;User ID=DSDlocal;Password=Hy@tt%43;Connect Timeout=1000000;");
            try
            {
                sqcon.Open();

                SqlDataAdapter daFetch = new SqlDataAdapter("[spFetchBasePointExecution]", sqcon);
                DataTable dt = new DataTable();
                daFetch.Fill(dt);

                if (dt.Rows.Count > 0)
                {

                    SqlCommand sqcmd = new SqlCommand("[spCalculateBasePoints]", sqcon);
                    sqcmd.CommandType = CommandType.StoredProcedure;
                    sqcmd.Parameters.AddWithValue("@SchemeID", dt.Rows[0]["SchemeId"].ToString());
                    sqcmd.Parameters.AddWithValue("@UserId", dt.Rows[0]["UserId"].ToString());
                    sqcmd.Parameters.AddWithValue("@FDate", dt.Rows[0]["FDate"].ToString());
                    sqcmd.Parameters.AddWithValue("@TDate", dt.Rows[0]["TDate"].ToString());
                    sqcmd.Parameters.AddWithValue("@SD", dt.Rows[0]["SD"].ToString());
                   
                    sqcmd.CommandTimeout = 0;
                    SqlDataAdapter da = new SqlDataAdapter(sqcmd);
                    DataSet ds = new DataSet();
                    da.Fill(ds);

                    Guid g;
                    g = Guid.NewGuid();
                    string filename = dt.Rows[0]["SchemeName"].ToString();
                    string downloadfilename = dt.Rows[0]["FileName"].ToString();


                   
                  //  ExportDataSet(ds,  "D:\\test\\" + downloadfilename, downloadfilename, filename);
                    ExportDataSet(ds, "C:\\Run_Codes\\IIS Root Folder\\RedClub\\PointCalcFile\\" + downloadfilename, downloadfilename, filename);
                  //  ExportDataSet(ds, "E:\\Run_Codes\\IIS Root Folder\\RC\\RedClubPortal\\PointCalcFile\\" + downloadfilename, downloadfilename, filename);
                }
            }
            catch (Exception ex)
            {
                if (File.Exists("C:\\Run_Codes\\IIS Root Folder\\RedClub\\PointCalcFile\\log1.txt"))
                {
                    File.Delete("C:\\Run_Codes\\IIS Root Folder\\RedClub\\PointCalcFile\\log1.txt");
                }
                // Create a file to write to.
                using (StreamWriter sw = File.CreateText("C:\\Run_Codes\\IIS Root Folder\\RedClub\\PointCalcFile\\log1.txt"))
                {
                    sw.WriteLine(ex.Message.ToString());

                }
            }
            finally
            {
                sqcon.Close();
            }
        }


        private void ExportDataSet(DataSet ds, string destination, string filename, string sheetname)
        {

            using (var workbook = SpreadsheetDocument.Create(destination, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
            {

                var workbookPart = workbook.AddWorkbookPart();

                workbook.WorkbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();

                workbook.WorkbookPart.Workbook.Sheets = new DocumentFormat.OpenXml.Spreadsheet.Sheets();
                ds.Tables[0].TableName = sheetname;
                ds.Tables[1].TableName = "Brand Wise NRV";
                ds.AcceptChanges();
                foreach (System.Data.DataTable table in ds.Tables)
                {

                    var sheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();
                    var sheetData = new DocumentFormat.OpenXml.Spreadsheet.SheetData();
                    sheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(sheetData);

                    DocumentFormat.OpenXml.Spreadsheet.Sheets sheets = workbook.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>();
                    string relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);

                    uint sheetId = 1;
                    if (sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Count() > 0)
                    {
                        sheetId =
                            sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                    }

                    DocumentFormat.OpenXml.Spreadsheet.Sheet sheet = new DocumentFormat.OpenXml.Spreadsheet.Sheet() { Id = relationshipId, SheetId = sheetId, Name = table.TableName };
                    sheets.Append(sheet);

                    DocumentFormat.OpenXml.Spreadsheet.Row headerRow = new DocumentFormat.OpenXml.Spreadsheet.Row();

                    List<String> columns = new List<string>();
                    foreach (System.Data.DataColumn column in table.Columns)
                    {
                        columns.Add(column.ColumnName);

                        DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                        cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                        cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(column.ColumnName);



                        headerRow.AppendChild(cell);
                    }


                    sheetData.AppendChild(headerRow);

                    foreach (System.Data.DataRow dsrow in table.Rows)
                    {
                        DocumentFormat.OpenXml.Spreadsheet.Row newRow = new DocumentFormat.OpenXml.Spreadsheet.Row();
                        foreach (String col in columns)
                        {
                            DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                            Int32 val;
                            Decimal val1;
                            if (Int32.TryParse(dsrow[col].ToString(), out val))
                            {
                                cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.Number;
                                cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(dsrow[col].ToString()); //
                            }
                            else if (Decimal.TryParse(dsrow[col].ToString(), out val1))
                            {
                                cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.Number;
                                cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(dsrow[col].ToString()); //
                            }
                            else
                            {
                                cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                                cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(dsrow[col].ToString()); //
                            }
                            newRow.AppendChild(cell);
                        }

                        sheetData.AppendChild(newRow);
                    }



                    
                }

                workbook.Close();
                workbook.Dispose();

            }

            // byte[] byteArray = File.ReadAllBytes(destination);
            // ViewState["sourceFile"] = "RPLScheme.xlsx";
            // Do work here
            //HttpContext.Current.Response.AppendHeader("content-disposition", "attachment; filename=" + filename);
            //HttpContext.Current.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            //HttpContext.Current.Response.BinaryWrite(byteArray);
            //HttpContext.Current.Response.Flush(); // Sends all currently buffered output to the client.
            //HttpContext.Current.Response.SuppressContent = true;  // Gets or sets a value indicating whether to send HTTP content to the client.
            //HttpContext.Current.ApplicationInstance.CompleteRequest(); 

        }

        static void Main(string[] args)
        {
            Program obj = new Program();
            obj.FetchData();
        }
    }
}
