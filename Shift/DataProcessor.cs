using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace Shift
{
    class DataProcessor
    {

        public DataProcessor() { }

        // set of characters (delimiters) that will separate into substrings
        char[] commaDelim = { ',', ' ' };

        public int[] Parse(String prefs)
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

        // Sorts the preferences of everyone into a calendar object and returns it
        public Calendar SortMostPreferred(Person[] persons)
        {
            Calendar prefCal = new Calendar();

            foreach (Person p in persons)
            { 
                foreach (int shift in p.prefs)
                {
                    prefCal.shifts[ShiftToArrayNum(shift)]++;
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


    } // end DataProcessing
} // namespace
