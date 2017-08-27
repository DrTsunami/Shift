using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Shift
{
    class Calendar
    {
        int[] shiftCalendar;

        // HACK create and mange this object. You're probably going to have 2 of these. one for the actual 
        //  schedule and one working one for operating your sorting process
        public Calendar()
        {
            shiftCalendar = new int[28];
        }
    }
}
