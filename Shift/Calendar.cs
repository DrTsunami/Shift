using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Shift
{
    class Calendar
    {
        String name;
        public int[] shifts;

        public Calendar(String name = "default")
        {
            this.name = name;
            shifts = new int[28];

            // for the default case
            if (name.Equals("test"))
            {
                for (int i = 0; i < 28; i++)
                {
                    shifts[i] = i;
                }
            }
        }

        // NOTE you're going to assign the index number of the person to the shift calendar
        // you'll just convert to names later. This way you don't really need to handle the people here
        public void WriteValue (int index, int value) 
        {
            shifts[index] = value;
        }

        public int GetValue (int index)
        {
            return (shifts[index]);
        }

        public void ConsoleOut ()
        {
            System.Console.WriteLine();

            System.Console.Write(name);

            System.Console.WriteLine();

            // print out each line of the calendar

            for (int i = 0; i < 28; i = i + 4)
            {
                System.Console.Write(shifts[i] + "\t");
            }
            System.Console.WriteLine();

            for (int i = 1; i <= 28; i = i + 4)
            {
                System.Console.Write(shifts[i] + "\t");
            }
            System.Console.WriteLine();

            for (int i = 2; i <= 28; i = i + 4)
            {
                System.Console.Write(shifts[i] + "\t");
            }
            System.Console.WriteLine();

            for (int i = 3; i <= 28; i = i + 4)
            {
                System.Console.Write(shifts[i] + "\t");
            }
            System.Console.WriteLine();
        }

    }
}
