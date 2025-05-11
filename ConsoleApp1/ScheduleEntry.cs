using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1
{
    public class ScheduleEntry
    {
        public int Id { get; set; }
        public int GroupId { get; set; } // Оставляем для обратной совместимости
        public List<int> GroupIds { get; set; } = new List<int>(); // Добавляем для поддержки потоков
        public int SubjectId { get; set; }
        public int LecturerId { get; set; }
        public int AuditoriumId { get; set; }
        public DateTime Date { get; set; }
        public TimeSpan Time { get; set; }
        public TimeSpan EndTime { get; set; }

        public bool ConflictsWith(ScheduleEntry other)
        {
            if (Date != other.Date) return false;
            if (EndTime <= other.Time || Time >= other.EndTime) return false;
            if (LecturerId == other.LecturerId) return true;
            if (AuditoriumId == other.AuditoriumId) return true;
            if (GroupIds.Any(g => other.GroupIds.Contains(g))) return true;
            return false;
        }
    }
}