using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1
{
    public class Auditorium
    {
        public int Id { get; set; }
        public string Number { get; set; }
        public int Capacity { get; set; }
        public string Equipment { get; set; }
        public List<(DateTime Date, TimeSpan Start, TimeSpan End)> BusyTimes { get; set; } = new List<(DateTime, TimeSpan, TimeSpan)>();
        public string Type { get; set; }

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