using System;
using System.Collections.Generic;

namespace ConsoleApp1
{
    public class Group
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public int StudentsCount { get; set; }
        public int Year { get; set; }
        public int DepartmentId { get; set; } // Изменили с string на int
        public int StreamId { get; set; }

        public List<(DateTime Date, TimeSpan Start, TimeSpan End)> BusyTimes { get; set; } = new List<(DateTime, TimeSpan, TimeSpan)>();

        public bool HasFreeSlot(DateTime date, TimeSpan start, TimeSpan end)
        {
            return !BusyTimes.Any(bt => bt.Date.Date == date.Date && !(bt.End <= start || bt.Start >= end));
        }
    }
}