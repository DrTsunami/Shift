using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Shift
{
    class Person
    {
        String name;
        DateTime[] prefs;
        DateTime timestamp;
        int seniority;

        // TODO add prefs back in
        public Person(String name, DateTime timestamp, int seniority)
        {
            this.name = name;
            // this.prefs = prefs;
            this.timestamp = timestamp;
            this.seniority = seniority;
        }

        public String printName()
        {
            return name;
        }
    }
}
