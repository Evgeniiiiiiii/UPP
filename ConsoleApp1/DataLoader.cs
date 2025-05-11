using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ConsoleApp1
{
    public static class DataLoader
    {
        public static List<Department> LoadDepartments(IXLWorksheet worksheet)
        {
            var departments = new List<Department>();
            foreach (var row in worksheet.RowsUsed().Skip(1))
            {
                var department = new Department
                {
                    Id = row.Field<int>("Id"),
                    Name = row.Field<string>("Name")
                };
                departments.Add(department);
            }
            return departments;
        }

        public static List<Group> LoadGroups(IXLWorksheet worksheet)
        {
            var groups = new List<Group>();
            foreach (var row in worksheet.RowsUsed().Skip(1))
            {
                var group = new Group
                {
                    Id = row.Field<int>("Id"),
                    Name = row.Field<string>("Name"),
                    StudentsCount = row.Field<int>("StudentsCount"),
                    Year = row.Field<int>("Year"),
                    DepartmentId = row.Field<int>("DepartmentId"),
                    StreamId = row.Field<int>("StreamId")
                };
                groups.Add(group);
            }
            return groups;
        }

        public static List<Stream> LoadStreams(IXLWorksheet worksheet, List<Group> groups)
        {
            var streams = new List<Stream>();
            foreach (var row in worksheet.RowsUsed().Skip(1))
            {
                var stream = new Stream
                {
                    Id = row.Field<int>("Id"),
                    DepartmentId = row.Field<int>("DepartmentId")
                };
                stream.Groups = groups.Where(g => g.StreamId == stream.Id).ToList();
                streams.Add(stream);
            }
            return streams;
        }

        public static List<Lecturer> LoadLecturers(IXLWorksheet worksheet)
        {
            var lecturers = new List<Lecturer>();
            foreach (var row in worksheet.RowsUsed().Skip(1))
            {
                var lecturer = new Lecturer
                {
                    Id = row.Field<int>("Id"),
                    Name = row.Field<string>("Name"),
                    MaxHoursPerWeek = row.Field<int>("MaxHoursPerWeek")
                };
                lecturers.Add(lecturer);
            }
            return lecturers;
        }

        public static List<Auditorium> LoadAuditoriums(IXLWorksheet worksheet)
        {
            var auditoriums = new List<Auditorium>();
            foreach (var row in worksheet.RowsUsed().Skip(1))
            {
                var auditorium = new Auditorium
                {
                    Id = row.Field<int>("Id"),
                    Number = row.Field<string>("Number"),
                    Capacity = row.Field<int>("Capacity"),
                    Type = row.Field<string>("Type"),
                    Equipment = row.Field<string>("Equipment")
                };
                auditoriums.Add(auditorium);
            }
            return auditoriums;
        }

        public static List<Subject> LoadSubjects(IXLWorksheet worksheet)
        {
            var subjects = new List<Subject>();
            foreach (var row in worksheet.RowsUsed().Skip(1))
            {
                var subject = new Subject
                {
                    Id = row.Field<int>("Id"),
                    Name = row.Field<string>("Name"),
                    WeeklyFrequency = row.Field<int>("WeeklyFrequency"),
                    Duration = row.Field<int>("Duration"),
                    LecturerId = row.Field<int>("LecturerId"),
                    GroupId = row.Field<int>("GroupId"),
                    DepartmentId = row.Field<int>("DepartmentId"),
                    LessonType = Enum.Parse<LessonType>(row.Field<string>("LessonType")),
                    IsForStream = row.Field<string>("IsForStream") == "Да",
                    WeekType = row.Field<string>("WeekType")
                };

                var auditoriums = row.Field<string>("AvailableAuditoriums").Split(',').Select(int.Parse).ToList();
                subject.AvailableAuditoriums = auditoriums;

                subjects.Add(subject);
            }
            return subjects;
        }

        private static T Field<T>(this IXLRow row, string columnName)
        {
            // Поиск столбца, игнорируя пробелы и регистр
            var column = row.Worksheet.ColumnsUsed()
                .FirstOrDefault(c => c.Cell(1).GetValue<string>()?.Replace(" ", "").ToLower() == columnName.Replace(" ", "").ToLower());
            if (column == null)
            {
                throw new ArgumentException($"Column '{columnName}' not found in the worksheet.");
            }
            var cellValue = row.Cell(column.ColumnNumber()).GetValue<string>()?.Trim() ?? "";
            if (string.IsNullOrEmpty(cellValue) && default(T) is not null)
            {
                return default(T); // Возвращаем значение по умолчанию
            }
            return (T)Convert.ChangeType(cellValue, typeof(T));
        }
    }
}