using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace Shift
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        public static void Start()
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
            // Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\Users\Ryan\Documents\GitHub\Shift\Shift\sheets\Responses.xlsx");
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"D:\Ryan\Documents\GitHub\Shift\Shift\sheets\Responses.xlsx");
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
            String[] names = s.GetStringData(xlWorksheet, nameCol, personCount);
            String[] prefs = s.GetStringData(xlWorksheet, prefCol, personCount);
            DateTime[] timestamps = s.GetTimestampData(xlWorksheet, timestampCol, personCount);
            int[] seniority = s.GetIntData(xlWorksheet, seniorityCol, personCount);



            ////////////////////////////////////////////////////////////////
            // METHODS
            ////////////////////////////////////////////////////////////////

            // Diagnostics
            s.CheckDataRange(xlRange);

            // Create and print people
            persons = s.CreatePersons(names, prefs, timestamps, seniority);
            // DEBUG test results using random seniorities
            Random r = new Random();
            foreach (Person p in persons)
            {
                p.RandomizeSeniority(r);
            }
            // end DEBUG
            s.ShowPeople(persons);

            // DEBUG-ish: print supposedly working calender #firsttry????
            Calendar testCalendar = dp.SortMostPreferred(persons);
            testCalendar.ConsoleOut();

            // Initialize vars to start assigning shifts
            Calendar shiftCalendar = new Calendar();
            List<int> unassigned = new List<int>();
            s.AssignShifts(testCalendar, shiftCalendar, persons, unassigned);

            // Print final data
            Console.WriteLine("----------------------------------------");
            Console.WriteLine("RESULTS");
            shiftCalendar.ConsoleOut();
            Console.WriteLine("Unassigned People: ");
            foreach (int i in unassigned)
            {
                Console.WriteLine(persons[i].name);
            }

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
