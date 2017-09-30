using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Net.Http;
using Newtonsoft.Json.Linq;
using Nito.AsyncEx;
using System.Runtime.InteropServices;
using System.Net;

namespace BulkIpGeoLocator
{
    class Program
    {
        const int CONCURRENT_TASKS = 5000;
        const string workbookPath = "E:\\src\\BulkIpGeoLocator\\20170925.xlsx";
        private static Excel.Application excelApp;
        private static Excel.Workbook excelWorkbook;
        private static Excel.Sheets excelSheets;
        static async Task<int> MainAsync(string[] args)
        {
            GeoLocator geo = new GeoLocator();
            
            string currentSheet = "Sheet4";
            Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelSheets.get_Item(currentSheet);
            //int x = 26720;
            int x = 0;
            try
            {
                
                while (x < excelWorksheet.Rows.Count)
                {
                    var tasks = new List<Task>();
                    // tasks.Add(geo.RunGeoLocator(excelWorksheet, x, "B"));
                    for (int y = 0; y < CONCURRENT_TASKS; y++)
                    {
                        try
                        {
                            tasks.Add(geo.RunGeoLocator(excelWorksheet, x + y, "B"));
                        }
                        catch (Exception exception)
                        {
                            Console.WriteLine(exception);
                        }
                    }

                    await Task.WhenAll(tasks);

                    excelWorkbook.Save();
                    //Parallel.Invoke(tasks);

                    x += CONCURRENT_TASKS;

                }
                return x;
                
               
            }
            catch (Exception ex)
            {
                return x;
            }
            finally
            {
                if (excelWorkbook != null) excelWorkbook.Close(0);
                excelApp.Quit();
            }
           
            
        }

        static void Main(string[] args)
        {
            excelApp = new Excel.Application();

            excelWorkbook = excelApp.Workbooks.Open(workbookPath,
                    0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "",
                    true, false, 0, true, false, false);
            excelSheets = excelWorkbook.Worksheets;
            int rowsProcessed = 0;
            var sheet = excelWorkbook.Worksheets["20170925"];
            int rowCount = sheet.Cells.Find("*", System.Reflection.Missing.Value,
                               System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                               Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                               false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;
            try
            {
                System.Net.ServicePointManager.DefaultConnectionLimit = 10;
                System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                rowsProcessed = AsyncContext.Run(() => MainAsync(args));
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine(ex);
                
                //return -1;
            }
            finally
            {
                //if (excelWorkbook != null) excelWorkbook.Close();
                excelApp.Quit();
                //excelWorkbook = null;
                //excelApp = null;

                Marshal.ReleaseComObject(excelSheets);
                Marshal.ReleaseComObject(excelWorkbook);
                Marshal.ReleaseComObject(excelApp);
            }

            Console.ReadKey();

            
        }

        




        
    }
}
