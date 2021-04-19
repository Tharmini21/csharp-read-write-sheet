using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
// Add nuget reference to smartsheet-csharp-sdk (https://www.nuget.org/packages/smartsheet-csharp-sdk/)
using Smartsheet.Api;
using Smartsheet.Api.Models;
using Smartsheet.Api.OAuth;
using MySql.Data.MySqlClient;
using System.Data;
using System.Data.SqlClient;

namespace csharp_read_write_sheet
{
    class Program
    {
        static Dictionary<string, long> columnMap = new Dictionary<string, long>();
        static void Main(string[] args)
        {
            try
            {
                var datasource = @"ILD-CHN-LAP-024\SQLEXPRESS";
                var database = "EmployeeDB"; 
                var username = "sa"; 
                var password = "Test@123"; 
                string cs = @"Data Source=" + datasource + ";Initial Catalog="+ database + ";Persist Security Info=True;User ID=" + username + ";Password=" + password;
                SqlConnection conn = new SqlConnection(cs);
                DataTable dt = new DataTable();
                try
                {
                    conn.Open();
                    string stm = "select * from Employee";
                    SqlCommand cmd = new SqlCommand(stm, conn);
                    SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd.CommandText, conn);
                    dataAdapter.Fill(dt);
                }
                catch (MySqlException ex)
                {
                    Console.WriteLine("Error: {0}", ex.ToString());
                }
                string smartsheetAPIToken = ConfigurationManager.AppSettings["AccessToken"];
                Token token = new Token();
                token.AccessToken = smartsheetAPIToken;
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(token.AccessToken).Build();
                //Sheet sheet = smartsheet.SheetResources.ImportXlsSheet("../../../EmployeeSheet _updated.xlsx", null, 0, null);
                //sheet = smartsheet.SheetResources.GetSheet(sheet.Id.Value, null, null, null, null, null, null, null);
                long SheetId = Convert.ToInt64(ConfigurationManager.AppSettings["SheetId"]);
                Sheet sheet = smartsheet.SheetResources.GetSheet(SheetId, null, null, null, null, null, null, null);
                Console.WriteLine("Loaded " + sheet.Rows.Count + " rows from sheet: " + sheet.Name);
               

                foreach (Column column in sheet.Columns)
                    columnMap.Add(column.Title, (long)column.Id);
               // if(dt.Rows.Count!=sheet.Rows.Count)
                if(dt.Rows.Count!=sheet.Rows.Count)
                {
                    Cell[] cellsA = null;
                    for (int i = dt.Rows.Count - 1; i >= 0; i--)
                    {
                        for (int j = 0; j < sheet.Columns.Count; j++)
                        {
                            cellsA = new Cell[]
                            {
                        //new Cell.AddCellBuilder(sheet.Columns[j].Id,dt.Rows[i].ItemArray.GetValue(j)).Build(),
                        //new Cell.AddCellBuilder(columnMap["EmployeeId"],dt.Rows[i].ItemArray.GetValue(j)).Build(),
                        new Cell.AddCellBuilder(columnMap["EmployeeId"],dt.Rows[i][0]).Build(),
                        new Cell.AddCellBuilder(columnMap["FirstName"],dt.Rows[i][1]).Build(),
                        new Cell.AddCellBuilder(columnMap["LastName"],dt.Rows[i][2]).Build(),
                        new Cell.AddCellBuilder(columnMap["Email"],dt.Rows[i][3]).Build(),
                        new Cell.AddCellBuilder(columnMap["Address"],dt.Rows[i][4]).Build(),
                            };
                            //Row rowA = new Row.AddRowBuilder(true, null, null, null, null).SetCells(cellsA).Build();
                            //smartsheet.SheetResources.RowResources.AddRows(sheet.Id.Value, new Row[] { rowA });
                        }
                        Row rowA = new Row.AddRowBuilder(true, null, null, null, null).SetCells(cellsA).Build();
                        smartsheet.SheetResources.RowResources.AddRows(sheet.Id.Value, new Row[] { rowA });
                    }
                }
               
                    List<Row> rowsToUpdate = new List<Row>();
                    foreach (Row row in sheet.Rows)
                    {
                        Row rowToUpdate = evaluateRowAndBuildUpdates(row);
                        if (rowToUpdate != null)
                            rowsToUpdate.Add(rowToUpdate);
                    }
                    Console.WriteLine("Writing " + rowsToUpdate.Count + " rows back to sheet id " + sheet.Id);
                    smartsheet.SheetResources.RowResources.UpdateRows(sheet.Id.Value, rowsToUpdate);
                   
                    //smartsheet.SheetResources.RowResources.DeleteRows(sheet.Id.Value, new long[] { 3413974005114756, 5366296352450436 }, true);  
                    smartsheet.SheetResources.RowResources.DeleteRows(sheet.Id.Value, new long[] { 4207621315291012, 8144091986061188 }, true);  
                    Console.WriteLine("Done...");
                    Console.ReadLine();
                               
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        
        static Row evaluateRowAndBuildUpdates(Row sourceRow)
        {
            Row rowToUpdate = null;
            Cell statusCell = getCellByColumnName(sourceRow, "LastName");
            if (statusCell.DisplayValue == "Clastname")
            {
                Cell remainingCell = getCellByColumnName(sourceRow, "LastName");
                if (remainingCell.DisplayValue != null)                 
                {
                    Console.WriteLine("Need to update row # " + sourceRow.RowNumber);

                    var cellToUpdate = new Cell
                    {
                        ColumnId = columnMap["LastName"],
                        Value = "UpdateCCLastName"
                    };

                    var cellsToUpdate = new List<Cell>();
                    cellsToUpdate.Add(cellToUpdate);

                    rowToUpdate = new Row
                    {
                        Id = sourceRow.Id,
                        Cells = cellsToUpdate
                    };
                }
            }
            return rowToUpdate;
        }

        // Helper function to find cell in a row
        static Cell getCellByColumnName(Row row, string columnName)
        {
            return row.Cells.First(cell => cell.ColumnId == columnMap[columnName]);
        }
    }
}
