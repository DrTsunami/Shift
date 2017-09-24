using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


using System.IO;
using System.Runtime.Serialization.Formatters.Binary;

namespace Shift
{
    [Serializable]
    public class Person : ICloneable
    {
        public String name;
        public int[] prefs;
        public DateTime timestamp;
        public int seniority;
        public bool assigned = false;
        public int shiftAssigned = -1;

        // backup vars
        public int[] prefsBak;

        public Person() { }

        public Person(String name, int[] prefs, DateTime timestamp, int seniority)
        {
            this.name = name;
            this.prefs = prefs;
            this.timestamp = timestamp;
            this.seniority = seniority;

            prefsBak = new int[prefs.Length];
        }

        // Cloning method
        public object Clone()
        {
            using (MemoryStream stream = new MemoryStream())
            {
                if (this.GetType().IsSerializable)
                {
                    BinaryFormatter formatter = new BinaryFormatter();
                    formatter.Serialize(stream, this);
                    stream.Position = 0;
                    return formatter.Deserialize(stream);
                }

                // DEBUG
                Console.WriteLine("ERROR: Clone process failed");
                return null;
            } 
        }

        public int GetShift()
        {
            return shiftAssigned;
        }

        public void Assign(int shiftIndex)
        {
            shiftAssigned = shiftIndex;
        }

        public void HidePrefs()
        {
            // make backup
            prefsBak = prefs;

            // Destroy current pref
            prefs = new int[0];
            assigned = true;
        }

        public void Restore()
        {
            if (assigned)
            {
                prefs = prefsBak;
            }
            else
            {
                Console.WriteLine("ERROR: You haven't destroyed this person yet!!");
            }
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
