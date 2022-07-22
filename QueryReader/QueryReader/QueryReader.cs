using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks; 
using System.Runtime.InteropServices;
using GemBox.Spreadsheet;
using ClosedXML.Excel;

namespace QueryReader
{
    public static class QueryReader
    {
        public static int indexSheet = 1;
        public static int indexFile = 1;
        public static List<DataTable> dataTables = new List<DataTable>();

       
        public static void ReadFile(string path)
        {
            


            string script = File.ReadAllText(path);
            string regexSemicolon = ";(?=(?:[^\"]*\"[^\"]*\")*(?![^\"]*\"))";
            string regexGo = @"^\s*GO\s*$";
            Console.WriteLine("Reading data");
            IEnumerable<string> commandStrings = Regex.Split(script, regexSemicolon, RegexOptions.Multiline | RegexOptions.IgnoreCase);
            Console.WriteLine("Executing querys");
            //Integrated Security=SSPI;
            SqlConnection connection = new SqlConnection(string.Format("Data Source={0};database={1};Integrated Security=SSPI; User ID={2};Password={3}", ConfigurationManager.AppSettings["ServerDatabase"], ConfigurationManager.AppSettings["Database"], ConfigurationManager.AppSettings["User"], ConfigurationManager.AppSettings["Pass"]));
            connection.Open();
            foreach (string commandString in commandStrings)
            {                
                if (!string.IsNullOrWhiteSpace(commandString.Trim()))
                {
                    using (var command = new SqlCommand(commandString, connection))
                    {
                        var adapter = new SqlDataAdapter(command);
                        var dataset = new DataSet();
                        adapter.Fill(dataset);
                        foreach (DataTable table in dataset.Tables)
                        {
                            dataTables.Add(table);
                        }
                        Console.WriteLine("Completed query");
                    }
                }
            }
            connection.Close();

            while (File.Exists(@ConfigurationManager.AppSettings["FileOutput"]+".xlsx"))
            {
                string pathFile = ConfigurationManager.AppSettings["FileOutput"];
                pathFile = pathFile.Remove(pathFile.Length-1,1) + indexFile;
                ConfigurationManager.AppSettings["FileOutput"] = pathFile;
                indexFile++;
            }
            Console.WriteLine("Choose output (t: txt file, e: excel file)");
            var response = Console.ReadKey();
            if(response.Key == ConsoleKey.T)
            {
                WriteTxtFile(dataTables);
            }else
            {
                WriteExcelFileClosed(@ConfigurationManager.AppSettings["FileOutput"] + ".xlsx", dataTables);
            }
            Console.WriteLine("Process Done. press any key to exit");
            Console.ReadKey();
            
        }

        #region ClosedXML
        private static void WriteExcelFileClosed(string path, List<DataTable> tables)
        {
            Console.WriteLine("Writing file");
            XLWorkbook workbook = new XLWorkbook();
            foreach(var table in tables)
            {
                workbook.Worksheets.Add(table, "Sheet" + indexSheet);                
                indexSheet++;
            }
            workbook.SaveAs(path);
        }
        #endregion


        #region GemBox
        private static void WriteExcelFile(string path, List<DataTable> tables)
        {
            Console.WriteLine("Writing file");           
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
            var workbook = new ExcelFile();
            foreach (var table in tables)
            {
                var worksheet = workbook.Worksheets.Add("Sheet" + indexSheet);
                worksheet.InsertDataTable(table, new InsertDataTableOptions() {
                    ColumnHeaders = true,
                    StartRow = 0,
                });                
                indexSheet++;
            }
            workbook.Save(path);
        }
        #endregion

        #region Excel Interop
        private static void WriteExcelFileInterop(string path, List<DataTable> tables)
        {
            Microsoft.Office.Interop.Excel.Application xlApp;
            Microsoft.Office.Interop.Excel.Workbook worKbooK;
            Microsoft.Office.Interop.Excel.Worksheet worKsheeT;
            Microsoft.Office.Interop.Excel.Range celLrangE;

            xlApp = new Microsoft.Office.Interop.Excel.Application();
            worKbooK = xlApp.Workbooks.Add(Type.Missing);

            foreach (var table in tables)
            {  
                worKsheeT = (Microsoft.Office.Interop.Excel.Worksheet)worKbooK.Worksheets.Add();
                worKsheeT.Name = "Sheet" + indexSheet;
                worKsheeT = (Microsoft.Office.Interop.Excel.Worksheet)worKbooK.ActiveSheet; worKsheeT.Name = "Sheet" + indexSheet;
                for (int col = 1; col < table.Columns.Count + 1; col++)
                {
                    worKsheeT.Cells[1, col] = table.Columns[col - 1].ColumnName;
                    for (int row = 2; row < table.Rows.Count + 2; row++)
                    {
                        worKsheeT.Cells[row, col] = table.Rows[row - 2][col - 1];
                    }
                }
                indexSheet++;
            }

            worKbooK.SaveAs(path); ;
            worKbooK.Close();
            xlApp.Quit();
        }
        #endregion

        #region OpenXML (alternative)
        private static void WriteExcelOpenXml(string path, List<DataTable> tables)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                foreach(var table in tables)
                {
                    WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    var sheetData = new SheetData();
                    worksheetPart.Worksheet = new Worksheet(sheetData);

                    Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
                    Sheet sheet = new Sheet()
                    { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = Convert.ToUInt32(indexSheet), Name = "Sheet" + indexSheet };

                    sheets.Append(sheet);
                    indexSheet++;

                    Row headerRow = new Row();

                    List<string> columns = new List<string>();
                    foreach (DataColumn column in table.Columns)
                    {
                        columns.Add(column.ColumnName);

                        Cell cell = new Cell();
                        cell.DataType = CellValues.String;
                        cell.CellValue = new CellValue(column.ColumnName);
                        headerRow.AppendChild(cell);
                    }

                    sheetData.AppendChild(headerRow);

                    foreach (DataRow dsrow in table.Rows)
                    {
                        Row newRow = new Row();
                        foreach (string col in columns)
                        {
                            Cell cell = new Cell();
                            cell.DataType = CellValues.String;
                            cell.CellValue = new CellValue(dsrow[col].ToString());
                            newRow.AppendChild(cell);
                        }

                        sheetData.AppendChild(newRow);
                    }
                }
               
               
                workbookPart.Workbook.Save();
            }

        }
            #endregion

        private static void WriteTxtFile(List<DataTable> tables)
        {
            foreach (var table in tables)
            {
                while (File.Exists(@ConfigurationManager.AppSettings["FileOutput"] + ".txt"))
                {
                    string pathFile = ConfigurationManager.AppSettings["FileOutput"];
                    pathFile = pathFile.Remove(pathFile.Length - 1, 1) + indexFile;
                    ConfigurationManager.AppSettings["FileOutput"] = pathFile;
                    indexFile++;
                }
                using (StreamWriter sw = File.CreateText(@ConfigurationManager.AppSettings["FileOutput"] + ".txt"))
                {
                    string columns = "";
                    foreach(DataColumn col in table.Columns)
                    {
                        columns += col + " || ";
                    }
                    sw.WriteLine(columns);
                    foreach (DataRow row in table.Rows)
                    {
                        sw.WriteLine(String.Join(" || ", row.ItemArray));
                    }
                }
            }
        }
    }
}
