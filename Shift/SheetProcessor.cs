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

        public void Reformat(string path)
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
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path);
            Excel.Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;


            ////////////////////////////////////////////////////////////////
            // VARS
            ////////////////////////////////////////////////////////////////

            // Classes
            DataProcessor dp = new DataProcessor();

            int personCount = 28;
            Person[] persons = new Person[personCount];


            // Excel Data
            int timestampCol = 1;
            int nameCol = 2;
            int prefCol = 4;
            int seniorityCol = 3;
            String[] names = GetStringData(xlWorksheet, nameCol, personCount);
            String[] prefs = GetStringData(xlWorksheet, prefCol, personCount);
            DateTime[] timestamps = GetTimestampData(xlWorksheet, timestampCol, personCount);
            String[] seniorData = GetStringData(xlWorksheet, seniorityCol, personCount);


            // DEBUG get dual entries
            /*
            for (int i = 0; i < names.Length; i++)
            {
                String thisName = names[i];

                for (int j = 0; j < names.Length; j++)
                {
                    if (j != i)
                    {
                        if (thisName.Equals(names[j]))
                        {
                            Console.WriteLine("Double entry line: " + j + " and " + i);
                            Console.WriteLine("NOTE: the numbers are 1 or 2 off);
                        }
                    }
                }
            }
            */

            ////////////////////////////////////////////////////////////////
            // PARSE SENIORITY
            ////////////////////////////////////////////////////////////////

            // Parse the data for seniority

            int thisYear = 2017; // CHANGE THIS WHEN YEAR CHANGES
            int thisSeason = 4; // 1. winter, 2. spring, 3. summer, 4. fall

            char[] delim = { ',', ' ' };
            int[] year = new int[seniorData.Length];
            int[] season = new int[seniorData.Length];
            int[] seniority = new int[seniorData.Length];

            for (int i = 0; i < seniorData.Length; i++)
            {
                String[] split = seniorData[i].Split(delim, System.StringSplitOptions.RemoveEmptyEntries);

                int.TryParse(split[1], out year[i]);

                switch (split[0])
                {
                    case ("Winter"):
                        season[i] = 1;
                        break;
                    case ("Spring"):
                        season[i] = 2;
                        break;
                    case ("Summer"):
                        season[i] = 3;
                        break;
                    case ("Fall"):
                        season[i] = 4;
                        break;
                    default:
                        Console.WriteLine("ERROR: invalid entry for seniority season");
                        break;
                }
            }
           

            for (int i = 0; i < seniorData.Length; i++)
            {
                int yearsBetweenModifier = (thisYear - year[i]) * 10;   // multiply by 10 to add the weight needed
                int seasonDifference = thisSeason - season[i];

                if (yearsBetweenModifier == 0)
                {
                    seniority[i] = seasonDifference;
                } else
                {
                    seniority[i] = yearsBetweenModifier - seasonDifference;
                }

                if (seniority[i] < 0)
                {
                    // DEBUG
                    Console.WriteLine("ERROR: someone is committing an act of tomfoolery. he/she is doomed to the last possible priority");
                }

            }

            foreach (int i in seniority)
            {
                Console.WriteLine(i);
            }
            


            ////////////////////////////////////////////////////////////////
            // CLEANUP
            ////////////////////////////////////////////////////////////////

            XlCleanup(xlApp, xlWorkbook, xlWorksheet, xlRange);

        }

        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        // Excel Data Processing
        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        /**
         * Returns string array of any specified col with Strings as inputs
         * NOTE: rowStart indicates when the data begins
         * 
         * @param xlWorksheet Excel Worksheet
         * @param col integer referencing the col we want to grab data from
         * @param personCount integer representing the amount of people/entries 
         */
        public String[] GetStringData(Excel.Worksheet xlWorksheet, int col, int personCount)
        {
            int rowStart = 2;
            int rowEnd = 29;
            int colNumber = col;

            String[] data = new String[personCount];
            for (int i = rowStart; i < rowEnd + 1; i++)
            {
                data[i - rowStart] = (xlWorksheet.Cells[i, colNumber] as Excel.Range).Value;
            }

            return data;
        }

        /**
         * Returns int array of any specified col with int as inputs
         * NOTE: rowStart indicates when the data begins
         * 
         * @param xlWorksheet Excel Worksheet
         * @param col integer referencing the col we want to grab data from
         * @param personCount integer representing the amount of people/entries 
         */
        public int[] GetIntData(Excel.Worksheet xlWorksheet, int col, int personCount)
        {
            int rowStart = 2;
            int rowEnd = 29;
            int colNumber = col;

            int[] data = new int[personCount];
            for (int i = rowStart; i < rowEnd + 1; i++)
            {
                data[i - rowStart] = (int)xlWorksheet.Cells[i, colNumber].Value;
            }
            return data;
        }

        /**
         * Returns DateTime array of any specified col with DateTime as inputs
         * NOTE: rowStart indicates when the data begins
         * 
         * @param xlWorksheet Excel Worksheet
         * @param col integer referencing the col we want to grab data from
         * @param personCount integer representing the amount of people/entries 
         */
        public DateTime[] GetTimestampData(Excel.Worksheet xlWorksheet, int col, int personCount)
        {
            int rowStart = 2;
            int rowEnd = 29;
            int colNumber = col;

            DateTime[] data = new DateTime[personCount];
            for (int i = rowStart; i < rowEnd + 1; i++)
            {
                data[i - rowStart] = xlWorksheet.Cells[i, colNumber].Value;
            }
            return data;
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
