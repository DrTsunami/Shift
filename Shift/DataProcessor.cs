using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;


namespace Shift
{
    class DataProcessor
    {

        public DataProcessor() { }

        // set of characters (delimiters) that will separate into substrings
        char[] commaDelim = { ',', ' ' };

        public int[] ParsePrefs(String prefs)
        {
            // begin parsing
            String[] splitString = prefs.Split(commaDelim, System.StringSplitOptions.RemoveEmptyEntries);
            int prefCount = (splitString.Length / 2);    // day and time count for 2 entries in the 
            String[] day = new String[prefCount];
            String[] time = new String[prefCount];

            // for each entry in the split string, sorts into a day and time array 
            for (int i = 0; i < splitString.Length; i++)
            {
                if (i == 0 || (i % 2) == 0)
                {
                    // if even entry in split string, indicating a day
                    int evenArrayNumber = (i / 2); // for even numbers
                    day[evenArrayNumber] = splitString[i];
                }
                else
                {
                    // if odd entry in split string, indicating a time
                    int oddArrayNumber = (int)(((float)i / 2f) - 0.5f);
                    time[oddArrayNumber] = splitString[i];
                }
            }

            // print preferences
            /*
            for (int i = 0; i < prefCount; i++)
            {
                System.Console.Write(day[i]);
                System.Console.WriteLine(" " + time[i]);
            }
            */

            // converts preferences into int array with preferences
            return PrefsToShiftNums(day, time);
        } // end parse

        /**
         * Converts the sheet data into seniority integers and writes to the excel file
         * 
         * @param worksheet Excel Worksheet output
         * @localparam rowStart integer representing the row number for which data starts
         */
        public void ConvertAndWriteSeniority(Excel.Worksheet worksheet, int seniorityCol, int personCount)
        {
            int rowStart = 2;

            // Parse the data for seniority

            String[] seniorData = GetStringData(worksheet, seniorityCol, personCount);
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
                // TODO stop this static reference to thisYear/szn
                int yearsBetweenModifier = (App.thisYear - year[i]) * 10;   // multiply by 10 to add the weight needed
                int seasonDifference = App.thisSeason - season[i];

                if (yearsBetweenModifier == 0)
                {
                    seniority[i] = seasonDifference;
                }
                else
                {
                    seniority[i] = yearsBetweenModifier - seasonDifference;
                }

                if (seniority[i] < 0)
                {
                    // DEBUG
                    Console.WriteLine("ERROR: someone is committing an act of tomfoolery. he/she is doomed to the last possible priority");
                }

            }

            // fill the data for seniority
            for (int i = rowStart; i < rowStart + personCount; i++)
            {
                (worksheet.Cells[i, seniorityCol] as Excel.Range).Value = seniority[i - rowStart];
            }
        }

        // Sorts the preferences of everyone into a calendar object and returns it
        public Calendar SortMostPreferred(Person[] persons)
        {
            Calendar prefCal = new Calendar();

            foreach (Person p in persons)
            { 
                if (!p.assigned)
                {
                    foreach (int shift in p.primaryPrefs)
                    {
                        // HACK need to figure out why the index is out of bounds here. happens on like the last iteration through the foreach loop.
                        prefCal.shifts[ShiftToArrayNum(shift)]++;
                    }
                }
            }

            // return calendar with indexes filled with preferences
            return prefCal;
        } // end SortMostPreferred



        // takes a one set of prefs (with days and times in separate arrays with same indexes) and converts to number system
        public int[] PrefsToShiftNums(String[] days, String[] times)
        {
            int prefCount = days.Length;
            int[] prefsAsShiftNums = new int[prefCount];

            for (int i = 0; i < prefCount; i++)
            {
                int shift;
                switch (days[i])
                {
                    case "Monday":
                        shift = 10;
                        break;
                    case "Tuesday":
                        shift = 20;
                        break;
                    case "Wednesday":
                        shift = 30;
                        break;
                    case "Thursday":
                        shift = 40;
                        break;
                    case "Friday":
                        shift = 50;
                        break;
                    case "Saturday":
                        shift = 60;
                        break;
                    case "Sunday":
                        shift = 70;
                        break;
                    default:
                        shift = 0;
                        break;
                }

                switch (times[i])
                {
                    case "8am-12pm":
                        shift = shift + 1;
                        break;
                    case "12pm-4pm":
                        shift = shift + 2;
                        break;
                    case "4pm-8pm":
                        shift = shift + 3;
                        break;
                    case "8pm-12am":
                        shift = shift + 4;
                        break;
                    default:
                        shift = shift + 0;
                        break;
                }
                prefsAsShiftNums[i] = shift;
                // write pref to the array

            } // end for loop for conversion of each entry into numbers

            // return ints
            return prefsAsShiftNums;
        } // end PrefsToShiftNums

        // takes shift number we converted to and converts it back to the calendar array compatible form
        public int ShiftToArrayNum(int shift)
        {
            int arrayNum;
            int day;
            int time;

            if (shift >= 10 && shift < 15)
            {
                // start of monday. day starts at 0. substact shift by 11 to make 8am 0
                day = 0;
                time = shift - 11;
                arrayNum = day + time;
            }
            else if (shift >= 20 && shift < 25)
            {
                // start of tues. day starts at 4. substact shift by 21 to make 8am 0
                day = 4;
                time = shift - 21;
                arrayNum = day + time;
            }
            else if (shift >= 30 && shift < 35)
            {
                // start of wed. day starts at 0. substact shift by 31 to make 8am 0
                day = 8;
                time = shift - 31;
                arrayNum = day + time;
            }
            else if (shift >= 40 && shift < 45)
            {
                // start of thuf. day starts at 0. substact shift by 41 to make 8am 0
                day = 12;
                time = shift - 41;
                arrayNum = day + time;
            }
            else if (shift >= 50 && shift < 55)
            {
                // start of fri. day starts at 0. substact shift by 51 to make 8am 0
                day = 16;
                time = shift - 51;
                arrayNum = day + time;
            }
            else if (shift >= 60 && shift < 65)
            {
                // start of sat. day starts at 0. substact shift by 61 to make 8am 0
                day = 20;
                time = shift - 61;
                arrayNum = day + time;
            }
            else if (shift >= 70 && shift < 75)
            {
                // start of sun. day starts at 0. substact shift by 71 to make 8am 0
                day = 24;
                time = shift - 71;
                arrayNum = day + time;
            }
            else
            {
                arrayNum = -1;
                Console.WriteLine("ERROR: invalid shift num");
            }

            return arrayNum;
        } // end ShiftToArrayNum

        private String[] GetStringData(Excel.Worksheet xlWorksheet, int col, int personCount)
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


    } // end DataProcessing
} // namespace
