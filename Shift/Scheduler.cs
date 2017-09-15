using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace Shift
{
    class Scheduler
    {
        DataProcessor dp = new DataProcessor();

        /**
         * Init
         */
        public Scheduler() { }

        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        // Shift Assigning
        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        // TODO make this return the calendar with shfits assigned
        public void AssignShifts(Calendar prefCal, Calendar shiftCal, Person[] persons, List<int> unassigned)
        {
            /////////////////////////////////////////////
            // Vars
            /////////////////////////////////////////////

            List<int> queue = new List<int>();

            /////////////////////////////////////////////

            // Cycles through 28 times (for each shift) searching in the order of the least preferred preferences first
            // i is the number of shifts this loop as completed
            // shiftIndex, in the scope of the smaller loop, is the index of the shfit number
            for (int i = 0; i < prefCal.shifts.Length; i++)
            {
                int leastPreferred = 99;
                int shiftIndex = 0;
                for (int pref = 0; pref < prefCal.shifts.Length; pref++)
                {
                    if (prefCal.shifts[pref] < leastPreferred && prefCal.shifts[pref] >= 0)
                    {
                        leastPreferred = prefCal.shifts[pref];
                        shiftIndex = pref;
                    }
                }

                // Begin process of assigning shift
                // list of people who prefer the current shift being examined

                List<int> peoplePref = new List<int>();
                peoplePref = GetPeoplePref(persons, shiftIndex);

                // allow for conflicts
                if (peoplePref.Count > 0)
                {
                    int personAssignedIndex = GetTopPriorityPerson(peoplePref, persons);

                    // assign person. Set in calendar and update person with shift
                    shiftCal.shifts[shiftIndex] = personAssignedIndex;
                    persons[personAssignedIndex].Assign(shiftIndex);

                    // remove 1 from each of preferences and destroy person
                    CleanPerson(persons[personAssignedIndex], prefCal);

                    // set the shift as taken
                    prefCal.shifts[shiftIndex] = -1;
                }
                else // if no people available to take shift 
                {
                    Console.WriteLine("CONFLICT");

                    // DEBUG
                    // prefCal.shifts[shiftIndex] = -2;
                    ResolveConflict(shiftIndex, persons, queue, shiftCal, prefCal);

                }

                queue.Add(shiftIndex);
                // DEBUG
                Console.WriteLine("Least Preferred Index: " + shiftIndex + "\t" + "Least Preferred Count" + leastPreferred);
                prefCal.ConsoleOut();
            } // end for loop going through 28 times for each shift

            // cycle through all the people in persons and create list of people without assigned spots
            for (int i = 0; i < persons.Length; i++)
            {
                if (!persons[i].assigned)
                {
                    unassigned.Add(i);
                }
            }

        } // end AssignShifts

        /**
         * Clean Person
         * Cleans the traces of the person just assigned. Foreach pref of that person,
         * we convert them into prefnumbers, and subtract 1 from the preference counter
         * calendar.
         */

        public void CleanPerson (Person p, Calendar prefCal)
        {
            // remove 1 from each of preferences
            foreach (int pref in p.prefs)
            {
                int prefNum = dp.ShiftToArrayNum(pref);
                // to account for preferences that are already assigned
                if (prefCal.shifts[prefNum] > 0)
                {
                    // subtract the preference from the prefcal
                    prefCal.shifts[prefNum]--;
                }
            }

            /* Person object destroyed
             * 1. Person prefs set to none: effectively making it so that when 
             *      we loop through to count preferences again (which i need 
             *      to set up) it will not detect any preferences from this 
             *      person
             * 2. Timestamp and seniority set to blanks, although i don't think
             *      this is a concern since if they have no preferences they 
             *      shouldn't be a factor in deciding who gets a shift
             * 3. Original data is stored in private variables for each object. 
             *      allowing for each object to have data restored if need be.
             * NOTE: potential problem... forgot what it was
             */
             p.Destroy();
        }

        /**
         * Prioritizes the list of PeoplePref and 
         * 
         */
        private int GetTopPriorityPerson(List<int> peoplePref, Person[] persons)
        {
            int personAssignedIndex = -1;

            // Goes through each person in the array list and compares to the next until the final person is found
            // the minus 1 is to account for the fact that you only go up to the second to last to compare with last person
            // if then loop to account for if only one person available
            if (peoplePref.Count == 1)
            {
                personAssignedIndex = peoplePref[0];
            }
            else // continue to compare
            {
                for (int j = 0; j < (peoplePref.Count - 1); j++)
                {
                    personAssignedIndex = ComparePeople(persons, peoplePref[j], peoplePref[j + 1]);
                    Console.WriteLine("personIndex: " + personAssignedIndex);
                }
            }

            return personAssignedIndex;
        }

        // Compares two people by means of their indexes in persons array. Returns index of prioritized person
        public int ComparePeople(Person[] persons, int index1, int index2)
        {
            Person p1 = persons[index1];
            Person p2 = persons[index2];
            int priority = -1;

            // check seniority first
            if (p1.seniority > p2.seniority)
            {
                priority = index1;
            }
            else if (p2.seniority > p1.seniority)
            {
                priority = index2;
            }
            else // if p1 and p2 seniorities equal
            {
                // if seniority ends in tie, then move on to timestamp
                int compareTime = DateTime.Compare(p1.timestamp, p2.timestamp);
                if (compareTime < 0)
                {
                    priority = index1;
                }
                else if (compareTime > 0)
                {
                    priority = index2;
                }
                else // if submission time equal 
                {
                    priority = index1;
                    Console.WriteLine("ERROR: Or sort of.... somehow the submission time exactly lined up.");
                }
            }

            return priority;
        }


        public List<int> GetPeoplePref(Person[] persons, int shiftNum)
        {
            List<int> peoplePref = new List<int>();

            // for every person in people, run through every preference and if the preference matches the shiftNum, add to arraylist and return
            for (int i = 0; i < persons.Length; i++)
            {
                foreach (int pref in persons[i].prefs)
                {
                    if (dp.ShiftToArrayNum(pref) == shiftNum)
                    {
                        peoplePref.Add(i);
                    }
                }
            }

            return peoplePref;
        }


        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        // Conflict Resolution
        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        /**
         * Resolves Conflict 
         * the main entry point for resolving conflicts if shift has no available people.
         * Takes in the shiftindex for the empty shift, the persons array, the queue list,
         * and the shiftcalendar. Makes list of assignedpeople, if there are previously
         * assigned people, starts from the end of the queue and goes backwards to try to 
         * cover and slide up the next person in the vacated spot. Peopel who move to fill the
         * problem shift are called COVERERS, those who slide to fill the recent vacancy
         * are called SLIDERS
         * 
         * @return void
         */
        public void ResolveConflict(int shiftIndex, Person[] persons, List<int> queue, Calendar shiftCal, Calendar prefCal)
        {
            List<int> assignedPersons = MakeAssignedList(persons);

            if (assignedPersons.Count > 0)
            {
                // Once list is created that isn't empty, need to go backwards through it
                for (int i = queue.Count - 1; i > 0; i--)
                {
                    // quick variables
                    int thisShift = queue[i];
                    int currentPersonIndex = shiftCal.shifts[thisShift];

                    if (PersonCanFill(shiftIndex, i, queue, shiftCal, persons))
                    {
                        List<int> sliders = GetPossibleSliders(thisShift, persons);

                        if (sliders.Count > 0)
                        {
                            CoverAndSlide(shiftIndex, currentPersonIndex, sliders, thisShift, shiftCal, persons, prefCal);
                            break; // end loop
                        } else
                        {
                            prefCal.shifts[shiftIndex] = -2;
                            shiftCal.shifts[shiftIndex] = -1;

                        }
                    } else
                    {
                        // if no fill found
                        prefCal.shifts[shiftIndex] = -2;
                        shiftCal.shifts[shiftIndex] = -1;
                    }

                    // else move to next queue number
                }
            } else
            {
                prefCal.shifts[shiftIndex] = -2;
                shiftCal.shifts[shiftIndex] = -1;
            }

        }

        /**
         * Makes Assigned List
         * Makes a list of all assigned people to this point. Iterates through the provided persons array 
         * and checks bool assigned to see if assigned. If assigned, then adds to a List and returns that
         * list
         * 
         * @return list of all assigned people
         */

        private List<int> MakeAssignedList(Person[] persons)
        {
            List<int> assigned = new List<int>();    // list of the indeces of people who are assigned

            // cycles through persons and adds anyone assigned to assignedPersons list
            for (int i = 0; i < persons.Length; i++)
            {
                if (persons[i].assigned)
                {
                    assigned.Add(i);
                }
            }

            return assigned;
        }

        /**
         * Checks if Person can fill
         * Foreach pref in the person currently occupying thisShift currently being looked at in the queue, 
         * we check to see if the person can cover.
         * 
         * @return true if person can cover, false otherwise
         */
        private bool PersonCanFill(int shiftToFillIndex, int qIndex, List<int> q, Calendar shiftCal, Person[] persons)
        {
            // check the person assigned at the current queue index
            /*
             * HACK - fix null pointer
             *  The problem here is that if a shift hasn't been assigned, it is 0, and basically it still uses that 0 value,
             *  so it is incorrectly trying to access person 0 right here. So to change it you can either
             *  1) check to make sure this shift on prefCal isn't -2 which means it can't be assigned 
             *  2) or make sure that when you can't assign a shift, instead of leaving it as 0 on shiftCal, you could change it 
             *     on shiftCal to something like -1 and check for that here. 
             */
            int thisShift = q[qIndex];
            int potentialFillIndex = shiftCal.shifts[thisShift];
            bool canCover = false;


            if (shiftCal.shifts[thisShift] != -1)
            {
                foreach (int pref in persons[potentialFillIndex].prefsBak)
                {
                    int convertedPref = dp.ShiftToArrayNum(pref);
                    if (convertedPref == shiftToFillIndex)   // if potentialFill had preferred shiftToFill
                    {
                        canCover = true;
                        break;
                    }
                }
            }

            return canCover;
        }

       

        /**
         * Covers the Empty spot, Slides up an available person
         * 
         */
        private void CoverAndSlide(int coverIndex, int covererIndex, List<int> sliders, 
            int slideToIndex, Calendar shiftCal, Person[] persons, Calendar prefCal)
        {
            // move coverer to the new shift
            shiftCal.shifts[coverIndex] = covererIndex;
            persons[covererIndex].Assign(coverIndex);

            // assign previous shift. Need to perform the same checks as before. Verify who is at top of priority
            int sliderIndex = GetTopPriorityPerson(sliders, persons);
            shiftCal.shifts[slideToIndex] = sliderIndex;
            persons[sliderIndex].Assign(slideToIndex);

            // clean newly assigned person
            CleanPerson(persons[sliderIndex], prefCal);

            // set the shift as taken
            prefCal.shifts[slideToIndex] = -1;

            // HACK combine previous checks, namely perosoncanslideup to help with the exchange here

        }

        /**
        * Checks and Gets if people can Slide up
        * Calls the getPeoplePref method to cycle through all available people. If there
        * is someone then 
        * 
        * @return List of people who can slide up
        */
        private List<int> GetPossibleSliders(int thisShiftIndex, Person[] persons)
        {
            List<int> slideable = GetPeoplePref(persons, thisShiftIndex);
            return slideable;
        }

        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        // Excel Data Processing
        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        public String[] GetStringData(Excel.Worksheet xlWorksheet, int col, int personCount)
        {
            int rowStart = 2;
            int rowEnd = 29;
            int colNumber = col;

            String[] data = new String[personCount];
            for (int i = rowStart; i < rowEnd + 1; i++)
            {
                data[i - rowStart] = xlWorksheet.Cells[i, colNumber].Value;
            }

            return data;
        }

        public int[] GetIntData(Excel.Worksheet xlWorksheet, int col, int personCount)
        {
            int rowStart = 2;
            int rowEnd = 29;
            int colNumber = col;

            int[] data = new int[personCount];
            for (int i = rowStart; i < rowEnd + 1; i++)
            {
                data[i - rowStart] = (int)xlWorksheet.Cells[i, colNumber].Value;
            }
            return data;
        }

        public DateTime[] GetTimestampData(Excel.Worksheet xlWorksheet, int col, int personCount)
        {
            int rowStart = 2;
            int rowEnd = 29;
            int colNumber = col;

            DateTime[] data = new DateTime[personCount];
            for (int i = rowStart; i < rowEnd + 1; i++)
            {
                data[i - rowStart] = xlWorksheet.Cells[i, colNumber].Value;
            }
            return data;
        }

        public Person[] CreatePersons(String[] names, String[] stringPrefs, DateTime[] timestamps, int[] seniority)
        {
            Person[] persons = new Person[names.Length];

            for (int i = 0; i < names.Length; i++)
            {
                int[] prefs = dp.Parse(stringPrefs[i]);

                // creates person
                persons[i] = new Person(names[i], prefs, timestamps[i], seniority[i]);
            }
            return persons;
        }

        public void ShowPeople(Person[] persons)
        {
            foreach (Person p in persons)
            {
                p.Print();
            }
        }

        // checks and prints data range in excel sheet
        public void CheckDataRange(Excel.Range range)
        {
            int rows = range.Rows.Count;
            int cols = range.Columns.Count;
            Console.WriteLine("rows: " + rows);
            Console.WriteLine("cols: " + cols);
        }

        // returns names of all entries based on Persons created
        public void PrintNamesOfPersons(Person[] persons)
        {
            foreach (Person p in persons)
            {
                Console.WriteLine(p.GetName());
            }
        }
    }
}
