using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Npgsql;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Dapper;
using System.Transactions;

namespace ExcelReader
{
    class Program
    {
        static void Main(string[] args)
        {
            /**************************************************************************
             * Reading from excel file
            **************************************************************************/

            //Create COM Objects. Create a COM object for everything that is referenced
            Application xlApp = new Application();
            Workbook xlWorkbook = xlApp.Workbooks.Open(@"D:\test.xlsx");
            Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            List<object> excelData = new List<object>();

            //iterate over the rows and columns and print to the console as it appears in the file
            //excel is not zero based!!!
            for (int i = 3; i <= rowCount; i++)
            {
                for (int j = 1; j <= colCount; j++)
                {
                    //write the value to the console
                    // if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                    //    Console.WriteLine(xlRange.Cells[i, j].Value2.ToString());

                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                        excelData.Add(new { Text = xlRange.Cells[i, j].Value2.ToString() });

                }
            }

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);

            /**************************************************************************
             * Storing data from excel file to database (PostgreSQL) by Dapper
            **************************************************************************/
            try
            {
                using (TransactionScope scope = new TransactionScope())
                {
                    using (var connection = new NpgsqlConnection("Host=localhost;Username=postgres;Password=XXX;Database=TestDb"))
                    {
                        connection.Execute("INSERT INTO \"Answers\" (\"Id\", \"CreateTime\", \"ModifyTime\", \"Text\", \"Value\") VALUES(uuid_generate_v4(), CURRENT_TIMESTAMP, CURRENT_TIMESTAMP, @Text, 99)", excelData);
                        
                        connection.Execute("INSERT INTO \"QuestionAnswer\" (\"Id\", \"CreateTime\", \"ModifyTime\", \"AnswerId\", \"QuestionId\")" +
                                            "SELECT uuid_generate_v4(), CURRENT_TIMESTAMP, CURRENT_TIMESTAMP, \"Id\", '123456789'" +
                                            "FROM \"Answers\"" +
                                            "WHERE \"Value\" = 99");

                        connection.Execute("UPDATE \"Answers\" SET \"Value\" = 0 WHERE \"Value\" = 99");

                        scope.Complete();
                    }
                }
            }
            catch (Exception e)
            {
                throw new Exception("Storing data from excel file to database failed!\n" + e.Message);
            }
        }
    }
}
