using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;       

namespace Sandbox
{
  
    public class Read_From_Excel
    {
        public static void Main()
        {

            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\Users\User\Desktop\Parsor\ExcelTest.csv");
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            List<string> AllValues = new List<string>();
            List<string> Dupes = new List<string>();
            
            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            


            for (int i = 1; i <= rowCount; i++)
            {
                for (int j = 1; j <= colCount; j++)
                {
                   

                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                 
                        AllValues.Add(xlRange.Cells[i, j].Value2.ToString()); 




                }
            }

          

            Dupes = AllValues.GroupBy(a => a).SelectMany(ab => ab.Skip(1).Take(1)).ToList();
            foreach (string g in Dupes)
            {

                Console.WriteLine(g);
            }
            
            GC.Collect();
            GC.WaitForPendingFinalizers();

 

           
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }
    }
}