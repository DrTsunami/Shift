using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace Shift
{
    public class SheetProcessor
    {
        public SheetProcessor() { }

        public void Reformat(Excel.Workbook wb)
        {
            ////////////////////////////////////////////////////////////////
            // EXCEL Setup
            ////////////////////////////////////////////////////////////////

            // kills any excel remaining processes.
            foreach (Process p in Process.GetProcesses())
            {
                // Console.WriteLine(p.ProcessName);
                if (p.ProcessName == "EXCEL")
                {
                    p.Kill();
                    Console.WriteLine("killed process");
                }
            }

            // TODO allow the program to pick a file.
            // TODO make a verification process (check a cell for a certain value) to check if the file is a valid one.
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = wb;
            Excel.Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            ////////////////////////////////////////////////////////////////
            // VARS
            ////////////////////////////////////////////////////////////////

            // Classes
            DataProcessor dp = new DataProcessor();
            Scheduler s = new Scheduler();

            int personCount = 28;
            Person[] persons = new Person[personCount];


            // Excel Data
            int timestampCol = 1;
            int nameCol = 2;
            int prefCol = 3;
            int seniorityCol = 4;




            ////////////////////////////////////////////////////////////////
            // CLEANUP
            ////////////////////////////////////////////////////////////////

            XlCleanup(xlApp, xlWorkbook, xlWorksheet, xlRange);

        }

        // TODO General cleanup. Releases Com objects. Takes care of Excel instance staying open after program finishes
        static void XlCleanup(Excel.Application xlApp, Excel.Workbook xlWorkbook, Excel.Worksheet xlWorksheet, Excel.Range xlRange)
        {
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
