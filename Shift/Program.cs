using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace Shift
{
    static class Program
    {
        // Creating reference objects
        static Excel.Application xlApp = new Excel.Application();
        // TODO figure out the path problem, bitch
        static Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\Users\Ryan\Documents\GitHub\Shift\Shift\Responses.xlsx");
        static Excel.Worksheet xlWorksheet = xlWorkbook.Sheets[1];
        static Excel.Range xlRange = xlWorksheet.UsedRange;
        

        [STAThread]
        static void Main()
        {
            System.Windows.Forms.Application.EnableVisualStyles();
            System.Windows.Forms.Application.SetCompatibleTextRenderingDefault(false);
            System.Windows.Forms.Application.Run(new Form1());
            XlSetup();
            XlCleanup();
        }


        static void XlSetup()
        {
            int rows = xlRange.Rows.Count;
            int cols = xlRange.Columns.Count;
            Console.WriteLine("rows: " + rows);
            Console.WriteLine("cols: " + cols);
        }

        // TODO kill processes when starting to prevent read onlyt issue
        static void ExcelProcessCleaner()
        {
            foreach (var process in Process.GetProcessesByName("Microsoft Excel"))
            {
                process.Kill();
                Console.WriteLine("killed a process");
            }
        }

        // General cleanup. Releases Com objects. Takes care of Excel instance staying open after program finishes
        static void XlCleanup()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
            Console.WriteLine("objects released");
        }
        


    }
}
