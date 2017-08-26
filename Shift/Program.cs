using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

////////////////////////////////
// TODO list
////////////////////////////////

// make sure to account for "coord shifts" and how that will fit into your program. Probably just a ghost person with low priority

////////////////////////////////


namespace Shift
{
    static class Program
    {

        [STAThread]
        static void Main()
        {
            // System.Windows.Forms.Application.EnableVisualStyles();
            // System.Windows.Forms.Application.SetCompatibleTextRenderingDefault(false);
            // System.Windows.Forms.Application.Run(new Form1());

            // Set up Excel
            // TODO figure out the path problem, bitch. Not really because you need to allow the program to pick a file.
            // TODO make a verification process (check a cell for a certain value) to check if the file is a valid one.
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\Users\Ryan\Documents\GitHub\Shift\Shift\sheets\Responses.xlsx");
            // Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"D:\Ryan\Documents\GitHub\Shift\Shift\sheets\Responses.xlsx");
            Excel.Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            ////////////////////////////////
            // Vars
            ////////////////////////////////

            // Excel Data
            int timestampCol = 1;
            int nameCol = 2;
            int prefCol = 3;
            int seniorityCol = 4;
            String[] names = getStringData(xlWorksheet, nameCol);
            String[] prefs = getStringData(xlWorksheet, prefCol);
            DateTime[] timestamps = getTimestampData(xlWorksheet, timestampCol);
            int[] seniority = getIntData(xlWorksheet, seniorityCol);

            Person[] persons = new Person[28];

            ////////////////////////////////////////////////////////////////

            createPersons(names, prefs, timestamps, seniority);
            DataProcessing dp = new DataProcessing();
            dp.parse();


            checkRowCol(xlRange);
            XlCleanup(xlApp, xlWorkbook, xlWorksheet, xlRange);

            ////////////////////////////////////////////////////////////////
            // TESTING
            ////////////////////////////////////////////////////////////////
            
            

            ////////////////////////////////////////////////////////////////
        }


        static void checkRowCol(Excel.Range range)
        {
            int rows = range.Rows.Count;
            int cols = range.Columns.Count;
            Console.WriteLine("rows: " + rows);
            Console.WriteLine("cols: " + cols);
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

        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        // Excel Data Processing
        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        static String[] getStringData(Excel.Worksheet xlWorksheet, int col)
        {
            int rowStart = 2;
            int rowEnd = 29;
            int colNumber = col;

            String[] data = new String[28];
            for (int i = rowStart; i < rowEnd; i++)
            {
                data[i - rowStart] = xlWorksheet.Cells[i, colNumber].Value;
            }
            return data;
        }

        static int[] getIntData(Excel.Worksheet xlWorksheet, int col)
        {
            int rowStart = 2;
            int rowEnd = 29;
            int colNumber = col;

            int[] data = new int[28];
            for (int i = rowStart; i < rowEnd; i++)
            {
                data[i - rowStart] = (int) xlWorksheet.Cells[i, colNumber].Value;
            }
            return data;
        }

        static DateTime[] getTimestampData(Excel.Worksheet xlWorksheet, int col)
        {
            int rowStart = 2;
            int rowEnd = 29;
            int colNumber = col;
            
            DateTime[] data = new DateTime[28];
            for (int i = rowStart; i < rowEnd; i++)
            {
                data[i - rowStart] = xlWorksheet.Cells[i, colNumber].Value;
            }
            return data;
        }

        static Person[] createPersons(String[] names, String[] prefs, DateTime[] timestamps, int[] seniority)
        {
            Person[] persons = new Person[28];

            for (int i = 0; i < 28; i++)
            {
                persons[i] = new Person(names[i], timestamps[i], seniority[i]);
                Console.WriteLine(persons[i].printName());
            }
            return persons;
        }

        
        
        /////////////////////////////////////////////////////////////////////////////////////////////////////////////////




    }
}
