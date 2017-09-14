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
        public bool assigned = false;
        public int shiftAssigned = -1;

        // backup vars
        public int[] prefsBak;
        public DateTime timestampBak;
        public int seniorityBak;



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

        // DEBUG temporary system to randomize the seniority. Will be replaced with real seniority
        public void RandomizeSeniority(Random r)
        {
            seniority = r.Next(6);
        }

        public int GetShift()
        {
            return shiftAssigned;
        }

        public void Assign(int shiftIndex)
        {
            shiftAssigned = shiftIndex;
        }

        public void Destroy()
        {
            // make backup
            prefsBak = prefs;
            timestampBak = timestamp;
            seniorityBak = seniority;

            // Destroy current values
            prefs = new int[0];
            timestamp = new DateTime();
            seniority = 0;
            assigned = true;
        }

        public void Restore()
        {
            if (assigned)
            {
                prefs = prefsBak;
                timestamp = timestampBak;
                seniority = seniorityBak;
            } else
            {
                Console.WriteLine("ERROR: You haven't destroyed this person yet!!");
            }
        }
    }
}
