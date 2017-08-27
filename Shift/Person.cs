using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Shift
{
    class Person
    {
        public String name;
        public int[] prefs;
        public DateTime timestamp;
        public int seniority;
        
        public Person(String name, int[] prefs, DateTime timestamp, int seniority)
        {
            this.name = name;
            this.prefs = prefs;
            this.timestamp = timestamp;
            this.seniority = seniority;
        }

        public String GetName()
        {
            return name;
        }

        public void Print()
        {
            System.Console.WriteLine(name);

            System.Console.Write("Preferences: ");

            foreach (int i in prefs)
            {
                System.Console.Write(i + " ");
            }

            System.Console.WriteLine();

            System.Console.WriteLine("Timestamp: " + timestamp);
            System.Console.WriteLine("Seniority: " + seniority);
            System.Console.WriteLine();
        }
    }
}
