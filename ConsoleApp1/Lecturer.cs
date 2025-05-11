using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1
{
    public class Lecturer
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public int MaxHoursPerWeek { get; set; }
        public List<string> PreferredSubjects { get; set; } = new List<string>();
        public List<(DateTime Date, TimeSpan Start, TimeSpan End)> BusyTimes { get; set; } = new List<(DateTime, TimeSpan, TimeSpan)>();
        public int CurrentHours { get; set; } = 0;

        public bool IsAvailable(DateTime date, TimeSpan start, TimeSpan end)
        {
            foreach (var busy in BusyTimes)
            {
                if (busy.Date == date && !(end <= busy.Start || start >= busy.End))
                    return false;
            }
            return true;
        }
    }
}