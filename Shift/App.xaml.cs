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
        // TODO make this an input box
        public static int thisYear = 2017; // CHANGE THIS WHEN YEAR CHANGES
        public static int thisSeason = 4; // 1. winter, 2. spring, 3. summer, 4. fall

        public static void Start(String path)
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
            
            // TODO make a verification process (check a cell for a certain value) to check if the file is a valid one.
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path);

            // creates the copy and makes it the active one
            xlWorkbook = xlApp.Workbooks.Open(CreateCopy(xlWorkbook));            

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

            // EDIT
            int timestampCol = 1;
            int nameCol = 2;
            int prefCol = 4;   
            int seniorityCol = 3;

            dp.ConvertAndWriteSeniority(xlWorksheet, seniorityCol, personCount);

            // create arrays of data from the excel sheet
            String[] names = new String[personCount];
            String[] prefs = new String[personCount];
            DateTime[] timestamps = new DateTime[personCount];
            int[] seniority = new int[personCount];

            try
            {
                names = s.GetStringData(xlWorksheet, nameCol, personCount);
                prefs = s.GetStringData(xlWorksheet, prefCol, personCount);
                timestamps = s.GetTimestampData(xlWorksheet, timestampCol, personCount);
                seniority = s.GetIntData(xlWorksheet, seniorityCol, personCount);

            }

            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
            {
                Console.WriteLine("ERROR: you have a null value");
            }
            


            ////////////////////////////////////////////////////////////////
            // METHODS
            ////////////////////////////////////////////////////////////////

            // Diagnostics
            s.CheckDataRange(xlRange);

            // Create and print people
            persons = s.CreatePersons(names, prefs, timestamps, seniority);
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

            SaveOutput(xlWorkbook);
            XlCleanup(xlApp, xlWorkbook, xlWorksheet, xlRange);
        }

        private static void SaveOutput(Excel.Workbook wb)
        {
            wb.Save();
        }

        private static void XlCleanup(Excel.Application xlApp, Excel.Workbook xlWorkbook, Excel.Worksheet xlWorksheet, Excel.Range xlRange)
        {
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
            Console.WriteLine("objects released");
        }

        private static String CreateCopy(Excel.Workbook wb)
        {
            // TODO give a file location
            String path = @"C:\Users\Ryan\Desktop\ShiftOutput.xlsx";
            wb.SaveCopyAs(path);
            return path;
        }
    }
}
