using System;
using System.Collections;
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

            // TODO allow the program to pick a file.
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

            DataProcessor dp = new DataProcessor();

            ////////////////////////////////////////////////////////////////
            // METHODS
            ////////////////////////////////////////////////////////////////

            // Diagnostics
            CheckDataRange(xlRange);

            // Create and print people
            persons = CreatePersons(dp, names, prefs, timestamps, seniority);
            // DEBUG test results using random seniorities
            Random r = new Random();
            foreach(Person p in persons)
            {
                p.RandomizeSeniority(r);
            }
            // end DEBUG
            ShowPeople(persons);

            // DEBUG-ish: print supposedly working calender #firsttry????
            Calendar testCalendar = dp.SortMostPreferred(persons);
            testCalendar.ConsoleOut();

            // HACK write code to start assigning shifts
            Calendar shiftCalendar = new Calendar();
            List<int> unassigned = new List<int>();
            AssignShifts(testCalendar, shiftCalendar, persons, dp, unassigned);

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

        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        // Shift Assigning
        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        // TODO make this return the calendar with shfits assigned
        static void AssignShifts(Calendar prefCal, Calendar shiftCal, Person[] persons, DataProcessor dp, List<int> unassigned)
        {
            // Cycles through 28 times (for each shift) searching in the order of the least preferred preferences first
            // i is the number of shifts this loop as completed
            // shiftIndex, in the scope of the smaller loop, is the index of the shfit number
            for (int i = 0; i < 28; i++)
            {
                int leastPreferred = 99;
                int shiftIndex = 0;     
                for (int pref = 0; pref < 28; pref++)
                {
                    if (prefCal.shifts[pref] < leastPreferred && prefCal.shifts[pref] >= 0)
                    {
                        leastPreferred = prefCal.shifts[pref];
                        shiftIndex = pref;
                    }
                }

                // Begin process of assigning shift
                // list of people who prefer the current shift being examined
                
                List<int> peoplePref = new List<int>();
                peoplePref = GetPeoplePref(persons, shiftIndex, dp);

                // allow for conflicts
                if (peoplePref.Count > 0)
                {
                    int personAssignedIndex = -1;

                    // Goes through each person in the array list and compares to the next until the final person is found
                    // the minus 1 is to account for the fact that you only go up to the second to last to compare with last person
                    // if then loop to account for if only one person available
                    if (peoplePref.Count == 1)
                    {
                        personAssignedIndex = peoplePref[0];
                    } else // contine to compare
                    {
                        for (int j = 0; j < (peoplePref.Count - 1); j++)
                        {
                            personAssignedIndex = ComparePeople(persons, peoplePref[j], peoplePref[j + 1]);
                            Console.WriteLine("personIndex: " + personAssignedIndex);
                        }
                    }
                    
                    // assign person
                    shiftCal.shifts[shiftIndex] = personAssignedIndex;

                    // remove 1 from each of preferences
                    foreach (int pref in persons[personAssignedIndex].prefs)
                    {
                        int prefNum = dp.ShiftToArrayNum(pref);
                        // to account for preferences that are already assigned
                        if (prefCal.shifts[prefNum] > 0)
                        {
                            // subtract the preference from the prefcal
                            prefCal.shifts[prefNum]--;
                        }
                    }

                    /* Person object destroyed
                     * 1. Person prefs set to none: effectively making it so that when 
                     *      we loop through to count preferences again (which i need 
                     *      to set up) it will not detect any preferences from this 
                     *      person
                     * 2. Timestamp and seniority set to blanks, although i don't think
                     *      this is a concern since if they have no preferences they 
                     *      shouldn't be a factor in deciding who gets a shift
                     * 3. Original data is stored in private variables for each object. 
                     *      allowing for each object to have data restored if need be.
                     * NOTE: potential problem... forgot what it was
                     */
                    persons[personAssignedIndex].Destroy();

                    // DEBUG
                    prefCal.shifts[shiftIndex] = -1;
                } else // if no people available to take shift 
                {
                    Console.WriteLine("CONFLICT");

                    // DEBUG
                    prefCal.shifts[shiftIndex] = -2;
                }
                
                // DEBUG
                Console.WriteLine("Least Preferred Index: " + shiftIndex + "\t" + "Least Preferred Count" + leastPreferred);
                prefCal.ConsoleOut();
            } // end for loop going through 28 times for each shift

            // cycle through all the people in persons and create list of people without assigned spots
            for (int i = 0; i < persons.Length; i++)
            {
                if (!persons[i].destroyed)
                {
                    unassigned.Add(i);
                }
            }

        } // end AssignShifts

        // Compares two people by means of their indexes in persons array. Returns index of prioritized person
        static int ComparePeople(Person[] persons, int index1, int index2)
        {
            Person p1 = persons[index1];
            Person p2 = persons[index2];
            int priority = -1;

            // check seniority first
            if (p1.seniority > p2.seniority)
            {
                priority = index1;
            } else if (p2.seniority > p1.seniority)
            {
                priority = index2;
            } else // if p1 and p2 seniorities equal
            {
                // if seniority ends in tie, then move on to timestamp
                int compareTime = DateTime.Compare(p1.timestamp, p2.timestamp);
                if (compareTime < 0)
                {
                    priority = index1;
                } else if (compareTime > 0)
                {
                    priority = index2;
                } else // if submission time equal 
                {
                    priority = index1;
                    Console.WriteLine("ERROR: Or sort of.... somehow the submission time exactly lined up.");
                }
            }

            return priority;
        }

        static List<int> GetPeoplePref (Person[] persons, int shiftNum, DataProcessor dp)
        {
            List<int> peoplePref = new List<int>();

            // for every person in people, run through every preference and if the preference matches the shiftNum, add to arraylist and return
            for (int i = 0; i < persons.Length; i++)
            {
                foreach (int pref in persons[i].prefs)
                {
                    if (dp.ShiftToArrayNum(pref) == shiftNum)
                    {
                        peoplePref.Add(i);
                    }
                }
            }

            return peoplePref;
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

        static Person[] CreatePersons(DataProcessor dp, String[] names, String[] stringPrefs, DateTime[] timestamps, int[] seniority)
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

        // checks and prints data range in excel sheet
        static void CheckDataRange(Excel.Range range)
        {
            int rows = range.Rows.Count;
            int cols = range.Columns.Count;
            Console.WriteLine("rows: " + rows);
            Console.WriteLine("cols: " + cols);
        }

        // returns names of all entries based on Persons created
        static void PrintNamesOfPersons(Person[] persons)
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



        /////////////////////////////////////////////////////////////////////////////////////////////////////////////////




    }
}
