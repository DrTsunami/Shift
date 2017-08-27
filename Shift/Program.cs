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

            // TODO figure out the path problem, bitch. Not really because you need to allow the program to pick a file.
            // TODO make a verification process (check a cell for a certain value) to check if the file is a valid one.
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\Users\Ryan\Documents\GitHub\Shift\Shift\sheets\Responses.xlsx");
            // Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"D:\Ryan\Documents\GitHub\Shift\Shift\sheets\Responses.xlsx");
            Excel.Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            ////////////////////////////////////////////////////////////////
            // VARS
            ////////////////////////////////////////////////////////////////

            // Excel Data
            int timestampCol = 1;
            int nameCol = 2;
            int prefCol = 3;
            int seniorityCol = 4;
            String[] names = GetStringData(xlWorksheet, nameCol);
            String[] prefs = GetStringData(xlWorksheet, prefCol);
            DateTime[] timestamps = GetTimestampData(xlWorksheet, timestampCol);
            int[] seniority = GetIntData(xlWorksheet, seniorityCol);

            Person[] persons = new Person[28];

            // Data processing
            DataProcessing dp = new DataProcessing();
            persons = CreatePersons(dp, names, prefs, timestamps, seniority);
            
            CheckDataRange(xlRange);

            ////////////////////////////////////////////////////////////////
            // TESTING
            ////////////////////////////////////////////////////////////////

            // print people
            ShowPeople(persons);

            ////////////////////////////////////////////////////////////////
            // CLEANUP
            ////////////////////////////////////////////////////////////////

            XlCleanup(xlApp, xlWorkbook, xlWorksheet, xlRange);
        }

        // checks and prints data range in excel sheet
        static void CheckDataRange(Excel.Range range)
        {
            int rows = range.Rows.Count;
            int cols = range.Columns.Count;
            Console.WriteLine("rows: " + rows);
            Console.WriteLine("cols: " + cols);
        }

        // returns names of all entries based on Persons created
        static void PrintNamesOfPersons (Person[] persons)
        {
            foreach (Person p in persons)
            {
                Console.WriteLine(p.GetName());
            }
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

        static String[] GetStringData(Excel.Worksheet xlWorksheet, int col)
        {
            int rowStart = 2;
            int rowEnd = 29;
            int colNumber = col;

            String[] data = new String[28];
            for (int i = rowStart; i < rowEnd + 1; i++)
            {
                data[i - rowStart] = xlWorksheet.Cells[i, colNumber].Value;
            }

            return data;
        }

        static int[] GetIntData(Excel.Worksheet xlWorksheet, int col)
        {
            int rowStart = 2;
            int rowEnd = 29;
            int colNumber = col;

            int[] data = new int[28];
            for (int i = rowStart; i < rowEnd + 1; i++)
            {
                data[i - rowStart] = (int) xlWorksheet.Cells[i, colNumber].Value;
            }
            return data;
        }

        static DateTime[] GetTimestampData(Excel.Worksheet xlWorksheet, int col)
        {
            int rowStart = 2;
            int rowEnd = 29;
            int colNumber = col;
            
            DateTime[] data = new DateTime[28];
            for (int i = rowStart; i < rowEnd + 1; i++)
            {
                data[i - rowStart] = xlWorksheet.Cells[i, colNumber].Value;
            }
            return data;
        }

        static Person[] CreatePersons(DataProcessing dp, String[] names, String[] stringPrefs, DateTime[] timestamps, int[] seniority)
        {
            Person[] persons = new Person[28];

            for (int i = 0; i < 28; i++)
            {
                int[] prefs = dp.Parse(stringPrefs[i]);

                // creates person
                persons[i] = new Person(names[i], prefs, timestamps[i], seniority[i]);
            }
            return persons;
        }

        static void ShowPeople (Person[] persons)
        {
            foreach (Person p in persons)
            {
                p.Print();
            }
        }

        
        
        /////////////////////////////////////////////////////////////////////////////////////////////////////////////////




    }
}
