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
        public int[] primaryPrefs;
        public int[] secondaryPrefs;
        public DateTime timestamp;
        public int seniority;
        public bool assigned = false;
        public int shiftAssigned = -1;

        // backup vars
        public int[] primaryPrefsBak;
        public int[] secondaryPrefsBak;

        public Person() { }

        public Person(String name, int[] primaryPrefs, int[] secondaryPrefs, DateTime timestamp, int seniority)
        {
            this.name = name;
            this.primaryPrefs = primaryPrefs;
            this.secondaryPrefs = secondaryPrefs;
            this.timestamp = timestamp;
            this.seniority = seniority;

            primaryPrefsBak = new int[primaryPrefs.Length];
            secondaryPrefsBak = new int[secondaryPrefs.Length];
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

        public void HidePrimaryPrefs()
        {
            // make backup
            primaryPrefsBak = primaryPrefs;

            // Destroy current pref
            primaryPrefs = new int[0];
            assigned = true;
        }

        public void HideSecondaryPrefs()
        {
            // make backup, destroy prefs, and mark assigned
            secondaryPrefsBak = secondaryPrefs;
            secondaryPrefs = new int[0];
            assigned = true;
        }

        public void RestorePrimaryPrefs()
        {
            if (assigned)
            {
                primaryPrefs = primaryPrefsBak;
            }
            else
            {
                Console.WriteLine("ERROR: You haven't destroyed this person yet!!");
            }
        }

        public void Print()
        {
            System.Console.WriteLine(name);

            System.Console.Write("Primary Preferences: ");

            foreach (int i in primaryPrefs)
            {
                System.Console.Write(i + " ");
            }

            System.Console.WriteLine();

            System.Console.Write("Secondary Preferences: ");
            foreach (int i in secondaryPrefs)
            {
                System.Console.Write(i + " ");
            }

            System.Console.WriteLine("Timestamp: " + timestamp);
            System.Console.WriteLine("Seniority: " + seniority);
            System.Console.WriteLine();
        }
    }
}
