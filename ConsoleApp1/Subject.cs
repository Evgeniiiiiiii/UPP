using System;
using System.Collections.Generic;

namespace ConsoleApp1
{
    public enum LessonType
    {
        Лекция,
        Практика,
        Лаборатория
    }

    public class Subject
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public int Duration { get; set; }
        public int WeeklyFrequency { get; set; }
        public int LecturerId { get; set; }
        public int GroupId { get; set; }
        public int DepartmentId { get; set; } // Изменили с string на int
        public LessonType LessonType { get; set; }
        public bool IsForStream { get; set; }
        public List<int> AvailableAuditoriums { get; set; } = new List<int>();
        public string WeekType { get; set; } = "каждая";

        public double FreedomScore
        {
            get
            {
                if (WeeklyFrequency == 0) return double.MaxValue;
                return (double)AvailableAuditoriums.Count / WeeklyFrequency;
            }
        }
    }
}