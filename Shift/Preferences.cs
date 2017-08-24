using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Shift
{
    class Preferences
    {
        DateTime[] shiftTimes;

        public Preferences(DateTime[] shiftTimes)
        {
            this.shiftTimes = shiftTimes;
        }

        public DateTime[] getShiftTimes()
        {
            return shiftTimes;
        }

        
    }
}
