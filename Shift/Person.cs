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

        // backup vars
        private int[] prefsBak;
        private DateTime timestampBak;
        private int seniorityBak;
        public bool destroyed = false;


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

        public void RandomizeSeniority(Random r)
        {
            seniority = r.Next(6);
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
            destroyed = true;
        }

        public void Restore()
        {
            if (destroyed)
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
