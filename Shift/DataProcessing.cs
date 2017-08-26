using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Shift
{
    class DataProcessing
    {
        // set of characters (delimiters) that will separate into substrings
        char[] commaDelim = {',', ' '};

        // public DateTime[] parse(String prefs)
        public void parse()
        {
            // DateTime[] prefArray;
            String testString = "TUES 4PM-8PM, TUES 8PM-12AM, WEDS 12PM-4PM, WEDS 4PM-8PM, WEDS 8PM-12AM";
        
            // begin parsing
            String[] splitString = testString.Split(commaDelim, System.StringSplitOptions.RemoveEmptyEntries);
            int prefCount = (testString.Length / 2);    // day and time count for 2 entries in the 
            String[] day = new String[prefCount];
            String[] time = new String[prefCount];


            for (int i = 0; i < splitString.Length; i++)
            {
                if (i == 0 || (i % 2) == 0) {
                    // if even entry in split string, indicating a day
                    int evenArrayNumber = (i / 2); // for even numbers
                    day[evenArrayNumber] = splitString[i];
                } else {
                    // if odd entry in split string, indicating a time
                    int oddArrayNumber = (int)(((float) i / 2f) - 0.5f);
                    time[oddArrayNumber] = splitString[i];
                }
            }


            // Print original string
            System.Console.WriteLine(testString);

            // print preferences
            for (int i = 0; i < prefCount; i++)
            {
                System.Console.Write(day[i]);
                System.Console.WriteLine(" " + time[i]);
            }

            // print them as numbers
            // TODO START HERE. NEED TO FIX DAY AND TIME TO ACTUALLY REFERENCE SOMETHING. FIX THE FUNCTION
            prefsToShiftNums(day, time);
            

            // return prefArray;
        }

        
        int[] prefsToShiftNums(String[] days, String[] times)
        {
            int prefCount = days.Length;
            int[] prefsAsShiftNums = new int[prefCount];

            for (int i = 0; i < prefCount; i++)
            {
                int shift;
                switch (days[i])
                {
                    case "MON":
                        shift = 10;
                        break;
                    case "TUES":
                        shift = 20;
                        break;
                    case "WEDS":
                        shift = 30;
                        break;
                    case "THURS":
                        shift = 40;
                        break;
                    case "FRI":
                        shift = 50;
                        break;
                    case "SAT":
                        shift = 60;
                        break;
                    case "SUN":
                        shift = 70;
                        break;
                    default:
                        shift = 0;
                        break;
                }
                System.Console.WriteLine("Day: " + shift);
                
                switch (times[i])
                {
                    case "8AM-12PM":
                        shift = shift + 1;
                        break;
                    case "12PM-4PM":
                        shift = shift + 2;
                        break;
                    case "4PM-8PM":
                        shift = shift + 3;
                        break;
                    case "8PM-12AM":
                        shift = shift + 4;
                        break;
                    default:
                        shift = shift + 0;
                        break;
                }

                System.Console.WriteLine("Day + Time: " + shift);
            }

            return prefsAsShiftNums;
        }
    }
}
