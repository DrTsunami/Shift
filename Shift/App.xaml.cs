/**
  *
  * todo list
  *      
  */

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
        public static int thisYear = 2017; // HACK THIS: CHANGE THIS WHEN YEAR CHANGES
        public static int thisSeason = 4; // 1. winter, 2. spring, 3. summer, 4. fall
        public static String outPath = "";

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

            // Print primary data
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
                s.AssignSecondaryShifts(prefCal, shiftCalendar, persons, out noAssignment, queue);
            }

            // Print secondary data
            Console.WriteLine("----------------------------------------");
            Console.WriteLine("SECONDARY RESULTS");
            shiftCalendar.ConsoleOut();
            Console.WriteLine("Unassigned People: ");
            foreach (int i in noAssignment)
            {
                Console.WriteLine(persons[i].name);
            }

            ////////////////////////////////////////////////////////////////
            // OUTPUT AND CLEANUP
            ////////////////////////////////////////////////////////////////

            WriteOutput(xlWorkbook, shiftCalendar, persons);
            SaveOutput(xlWorkbook);
            XlCleanup(xlApp, xlWorkbook, xlWorksheet, xlRange);
        }

        private static void WriteOutput(Excel.Workbook wb, Calendar shiftCal, Person[] persons)
        {
            Excel.Worksheet outSheet;
            outSheet = (Excel.Worksheet)wb.Worksheets.Add();

            // fill out the day and times
            (outSheet.Cells[1, 2] as Excel.Range).Value = "Mon";
            (outSheet.Cells[1, 3] as Excel.Range).Value = "Tue";
            (outSheet.Cells[1, 4] as Excel.Range).Value = "Wed";
            (outSheet.Cells[1, 5] as Excel.Range).Value = "Thur";
            (outSheet.Cells[1, 6] as Excel.Range).Value = "Fri";
            (outSheet.Cells[1, 7] as Excel.Range).Value = "Sat";
            (outSheet.Cells[1, 8] as Excel.Range).Value = "Sun";

            (outSheet.Cells[2, 1] as Excel.Range).Value = "8am-12pm";
            (outSheet.Cells[3, 1] as Excel.Range).Value = "12pm-4pm";
            (outSheet.Cells[4, 1] as Excel.Range).Value = "4pm-8pm";
            (outSheet.Cells[5, 1] as Excel.Range).Value = "8pm-12am";

            // fill out people into the sheet
            int day = 1;
            int counter = 0;
            for (int i = 0; i < shiftCal.shifts.Length; i++)
            {
                int rowTime;    // the row
                int colDay;     // the time

                // counter will be either 0, 1, 2, 3 (early -> late)
                counter = i % 4;

                // assign the day
                if (counter == 0)
                {
                    day++;
                }

                colDay = day;
                rowTime = counter + 2;

                // assign the person. i is the index of the shift. 
                if (shiftCal.shifts[i] != -1)
                {
                    Person p = persons[shiftCal.shifts[i]];
                    (outSheet.Cells[rowTime, colDay] as Excel.Range).Value = p.name;
                } else
                {
                    (outSheet.Cells[rowTime, colDay] as Excel.Range).Value = "unassigned";
                }
                
                
            }
            

            // styles
            (outSheet.Cells[1, 2] as Excel.Range).EntireRow.Font.Bold = true;
            (outSheet.Cells[2, 1] as Excel.Range).EntireColumn.Font.Bold = true;

            int activeRange = 9;
            for (int i = 1; i < activeRange; i++)
            {
                (outSheet.Cells[1, i] as Excel.Range).EntireColumn.AutoFit();
            }


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
            // String path = @"C:\Users\Ryan\Desktop\ShiftOutput.xlsx";
            String path = @"C:\Users\rvtsa\Desktop\ShiftOutput.xlsx";
            wb.SaveCopyAs(path);

            /*
            if (outPath != null)
            {
                String outputPath = (outPath + "ShiftOutput.xlsx");
                wb.SaveCopyAs(@outputPath);
                return @outputPath;
            } else
            {
                Console.WriteLine("ERROR: you don't have a path selected");
                return "null";
            }
            */
            
            // HACK please print out who is unassigned to the excel sheet

            return path;
        }
    }
}
