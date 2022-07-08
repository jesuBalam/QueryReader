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


namespace QueryReader
{
    public static class QueryReader
    {
        public static int indexSheet = 1;
        public static int indexFile = 1;
        public static List<DataTable> dataTables = new List<DataTable>();

        public static Microsoft.Office.Interop.Excel.Application xlApp;
        public static Microsoft.Office.Interop.Excel.Workbook worKbooK;
        public static Microsoft.Office.Interop.Excel.Worksheet worKsheeT;
        public static Microsoft.Office.Interop.Excel.Range celLrangE;
        public static void ReadFile(string path)
        {
            


            string script = File.ReadAllText(path);
            string regexSemicolon = ";(?=(?:[^\"]*\"[^\"]*\")*(?![^\"]*\"))";
            string regexGo = @"^\s*GO\s*$";
            Console.WriteLine("Reading data");
            IEnumerable<string> commandStrings = Regex.Split(script, regexSemicolon, RegexOptions.Multiline | RegexOptions.IgnoreCase);
            Console.WriteLine("Executing querys");
            SqlConnection connection = new SqlConnection(string.Format("Data Source={0};database={1}; User ID={2};Password={3}", ConfigurationManager.AppSettings["ServerDatabase"], ConfigurationManager.AppSettings["Database"], ConfigurationManager.AppSettings["User"], ConfigurationManager.AppSettings["Pass"]));
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
            while(File.Exists(@ConfigurationManager.AppSettings["FileOutput"]+".xlsx"))
            {
                string pathFile = ConfigurationManager.AppSettings["FileOutput"];
                pathFile = pathFile.Remove(pathFile.Length-1,1) + indexFile;
                ConfigurationManager.AppSettings["FileOutput"] = pathFile;
                indexFile++;
            }                
            WriteExcelFile(@ConfigurationManager.AppSettings["FileOutput"]+".xlsx", dataTables);
            connection.Close();
        }

        private static void WriteExcelFile(string path, List<DataTable> tables)
        {
            Console.WriteLine("Writing file");
            xlApp = new Microsoft.Office.Interop.Excel.Application();
            worKbooK = xlApp.Workbooks.Add(Type.Missing);



            //using (SpreadsheetDocument document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook))
            //{
            //WorkbookPart workbookPart = document.AddWorkbookPart();
            //workbookPart.Workbook = new Workbook();

            foreach (var table in tables)
            {
                #region Excel
                worKsheeT = (Microsoft.Office.Interop.Excel.Worksheet)worKbooK.Worksheets.Add();
                worKsheeT.Name = "Sheet" + indexSheet;
                //worKsheeT = (Microsoft.Office.Interop.Excel.Worksheet)worKbooK.ActiveSheet; worKsheeT.Name = "Sheet" + indexSheet;
                for (int col = 1; col < table.Columns.Count + 1; col++)
                {
                    worKsheeT.Cells[1, col] = table.Columns[col - 1].ColumnName;
                    for (int row = 2; row < table.Rows.Count + 2; row++)
                    {
                        worKsheeT.Cells[row, col] = table.Rows[row - 2][col - 1];
                    }
                }
                indexSheet++;
                #endregion

                #region OpenXML (alternative)
                //WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                //var sheetData = new SheetData();
                //worksheetPart.Worksheet = new Worksheet(sheetData);

                //Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
                //Sheet sheet = new Sheet()
                //{ Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = Convert.ToUInt32(indexSheet), Name = "Sheet" + indexSheet };

                //sheets.Append(sheet);
                //indexSheet++;

                //Row headerRow = new Row();

                //List<string> columns = new List<string>();
                //foreach (DataColumn column in table.Columns)
                //{
                //    columns.Add(column.ColumnName);

                //    Cell cell = new Cell();
                //    cell.DataType = CellValues.String;
                //    cell.CellValue = new CellValue(column.ColumnName);
                //    headerRow.AppendChild(cell);
                //}

                //sheetData.AppendChild(headerRow);

                //foreach (DataRow dsrow in table.Rows)
                //{
                //    Row newRow = new Row();
                //    foreach (string col in columns)
                //    {
                //        Cell cell = new Cell();
                //        cell.DataType = CellValues.String;
                //        cell.CellValue = new CellValue(dsrow[col].ToString());
                //        newRow.AppendChild(cell);
                //    }

                //    sheetData.AppendChild(newRow);
                //}
                #endregion
            }
            worKbooK.SaveAs(path); ;
            worKbooK.Close();
            xlApp.Quit();

            //workbookPart.Workbook.Save();
            //}
        }
    }
}
