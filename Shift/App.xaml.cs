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

            // EDIT: change the column numbers here
            int timestampCol = 1;
            int nameCol = 2;
            int seniorityCol = 3;
            int primaryPrefCol = 4;
            int secondaryPrefCol = 5;

            dp.ConvertAndWriteSeniority(xlWorksheet, seniorityCol, personCount);

            // create arrays of data from the excel sheet
            String[] names = new String[personCount];
            String[] primaryPrefs = new String[personCount];
            String[] secondaryPrefs = new String[personCount];
            DateTime[] timestamps = new DateTime[personCount];
            int[] seniority = new int[personCount];

            try
            {
                names = s.GetStringData(xlWorksheet, nameCol, personCount);
                primaryPrefs = s.GetStringData(xlWorksheet, primaryPrefCol, personCount);
                secondaryPrefs = s.GetStringData(xlWorksheet, secondaryPrefCol, personCount);
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
            persons = s.CreatePersons(names, primaryPrefs, secondaryPrefs, timestamps, seniority);
            s.ShowPeople(persons);

            // DEBUG-ish: print supposedly working calender #firsttry????
            Calendar prefCal = dp.SortMostPreferred(persons);
            prefCal.ConsoleOut();

            // Initialize vars to start assigning shifts
            Calendar shiftCalendar = new Calendar();
            List<int> noAssignment = new List<int>();
            s.AssignShifts(prefCal, shiftCalendar, persons, noAssignment, out List<int> queue);

            // Print final data
            Console.WriteLine("----------------------------------------");
            Console.WriteLine("PRIMARY RESULTS");
            shiftCalendar.ConsoleOut();
            Console.WriteLine("Unassigned People: ");
            foreach (int i in noAssignment)
            {
                Console.WriteLine(persons[i].name);
            }

            // if the calendar still doesn't work... go to secondary preferences.
            if (noAssignment.Count > 0)
            {
                s.AssignSecondaryShifts(prefCal, shiftCalendar, persons, noAssignment, queue);
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
            // String path = @"C:\Users\Ryan\Desktop\ShiftOutput.xlsx"
            String path = @"C:\Users\rvtsa\Desktop\ShiftOutput.xlsx";
            wb.SaveCopyAs(path);
            return path;
        }
    }
}
