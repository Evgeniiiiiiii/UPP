using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;

namespace ConsoleApp1
{
    class Program
    {
        static bool IsEvenWeek(DateTime date)
        {
            DateTime referenceDate = new DateTime(2025, 4, 1); // Считаем 01.04.2025 чётной неделей
            int daysDifference = (date - referenceDate).Days;
            int weekNumber = daysDifference / 7 + 1;
            return weekNumber % 2 == 0;
        }

        static bool CheckScheduleConflicts(List<ScheduleEntry> schedule)
        {
            for (int i = 0; i < schedule.Count; i++)
            {
                for (int j = i + 1; j < schedule.Count; j++)
                {
                    if (schedule[i].ConflictsWith(schedule[j]))
                    {
                        Console.WriteLine($"Конфликт в расписании: Занятия {i + 1} и {j + 1} пересекаются.");
                        return false;
                    }
                }
            }
            return true;
        }

        static bool CheckGroupDailyLimit(List<ScheduleEntry> schedule, List<Group> groups, int maxLessonsPerDay = 4)
        {
            foreach (var group in groups)
            {
                var groupEntries = schedule.Where(e => e.GroupIds.Contains(group.Id))
                                           .GroupBy(e => e.Date.Date)
                                           .ToList();

                foreach (var dailyEntries in groupEntries)
                {
                    if (dailyEntries.Count() > maxLessonsPerDay)
                    {
                        Console.WriteLine($"Группа {group.Name} имеет слишком много занятий ({dailyEntries.Count()}) в день {dailyEntries.Key:dd.MM.yyyy}. Максимум: {maxLessonsPerDay}");
                        return false;
                    }
                }
            }
            return true;
        }

        static bool CheckGroupBreakTime(List<ScheduleEntry> schedule, List<Group> groups, TimeSpan minBreak = default)
        {
            if (minBreak == default) minBreak = TimeSpan.FromMinutes(10);

            foreach (var group in groups)
            {
                var groupEntries = schedule.Where(e => e.GroupIds.Contains(group.Id))
                                           .OrderBy(e => e.Date)
                                           .ThenBy(e => e.Time)
                                           .ToList();

                for (int i = 0; i < groupEntries.Count - 1; i++)
                {
                    var currentEntry = groupEntries[i];
                    var nextEntry = groupEntries[i + 1];

                    if (currentEntry.Date != nextEntry.Date) continue;

                    TimeSpan breakTime = nextEntry.Time - currentEntry.EndTime;
                    if (breakTime < minBreak)
                    {
                        Console.WriteLine($"Недостаточный перерыв для группы {group.Name} между занятиями " +
                                          $"{currentEntry.Date:dd.MM.yyyy} {currentEntry.Time:hh\\:mm}-{currentEntry.EndTime:hh\\:mm} и " +
                                          $"{nextEntry.Date:dd.MM.yyyy} {nextEntry.Time:hh\\:mm}-{nextEntry.EndTime:hh\\:mm}. " +
                                          $"Перерыв: {breakTime.TotalMinutes} минут, минимум: {minBreak.TotalMinutes} минут");
                        return false;
                    }
                }
            }
            return true;
        }

        static void EvaluateGroupWindows(List<ScheduleEntry> schedule, List<Group> groups, TimeSpan maxAcceptableWindow)
        {
            Console.WriteLine("\nОценка 'окон' для групп:");
            foreach (var group in groups)
            {
                var groupEntries = schedule.Where(e => e.GroupIds.Contains(group.Id))
                                           .OrderBy(e => e.Date)
                                           .ThenBy(e => e.Time)
                                           .ToList();

                for (int i = 0; i < groupEntries.Count - 1; i++)
                {
                    var currentEntry = groupEntries[i];
                    var nextEntry = groupEntries[i + 1];

                    if (currentEntry.Date != nextEntry.Date) continue;

                    TimeSpan window = nextEntry.Time - currentEntry.EndTime;
                    if (window > maxAcceptableWindow)
                    {
                        Console.WriteLine($"Большое 'окно' для группы {group.Name} в {currentEntry.Date:dd.MM.yyyy}: " +
                                          $"между {currentEntry.Time:hh\\:mm}-{currentEntry.EndTime:hh\\:mm} и " +
                                          $"{nextEntry.Time:hh\\:mm}-{nextEntry.EndTime:hh\\:mm}. " +
                                          $"Окно: {window.TotalHours:F1} часов");
                    }
                }
            }
        }

        static void EvaluateLecturerWindows(List<ScheduleEntry> schedule, List<Lecturer> lecturers, TimeSpan maxAcceptableWindow)
        {
            Console.WriteLine("\nОценка 'окон' для преподавателей:");
            foreach (var lecturer in lecturers)
            {
                var lecturerEntries = schedule.Where(e => e.LecturerId == lecturer.Id)
                                             .OrderBy(e => e.Date)
                                             .ThenBy(e => e.Time)
                                             .ToList();

                for (int i = 0; i < lecturerEntries.Count - 1; i++)
                {
                    var currentEntry = lecturerEntries[i];
                    var nextEntry = lecturerEntries[i + 1];

                    if (currentEntry.Date != nextEntry.Date) continue;

                    TimeSpan window = nextEntry.Time - currentEntry.EndTime;
                    if (window > maxAcceptableWindow)
                    {
                        Console.WriteLine($"Большое 'окно' для преподавателя {lecturer.Name} в {currentEntry.Date:dd.MM.yyyy}: " +
                                          $"между {currentEntry.Time:hh\\:mm}-{currentEntry.EndTime:hh\\:mm} и " +
                                          $"{nextEntry.Time:hh\\:mm}-{nextEntry.EndTime:hh\\:mm}. " +
                                          $"Окно: {window.TotalHours:F1} часов");
                    }
                }
            }
        }

        static void EvaluateGroupLoadDistribution(List<ScheduleEntry> schedule, List<Group> groups, int totalDays)
        {
            Console.WriteLine("\nОценка равномерности нагрузки групп:");
            foreach (var group in groups)
            {
                var groupEntries = schedule.Where(e => e.GroupIds.Contains(group.Id))
                                           .GroupBy(e => e.Date.Date)
                                           .ToList();

                int totalLessons = groupEntries.Sum(g => g.Count());
                double idealLessonsPerDay = (double)totalLessons / totalDays;
                double variance = 0;

                for (int day = 0; day < totalDays; day++)
                {
                    DateTime date = schedule.Min(e => e.Date.Date).AddDays(day);
                    int lessonsOnDay = groupEntries.FirstOrDefault(g => g.Key == date)?.Count() ?? 0;
                    variance += Math.Pow(lessonsOnDay - idealLessonsPerDay, 2);
                }

                variance /= totalDays;
                Console.WriteLine($"Группа {group.Name}: Всего занятий: {totalLessons}, Идеально в день: {idealLessonsPerDay:F1}, Дисперсия: {variance:F2}");
                if (variance > 2)
                {
                    Console.WriteLine($"Нагрузка группы {group.Name} неравномерна! Рекомендуется распределить занятия более равномерно.");
                }
            }
        }

        static void EvaluateLecturerLoadDistribution(List<ScheduleEntry> schedule, List<Lecturer> lecturers, int totalDays)
        {
            Console.WriteLine("\nОценка равномерности нагрузки преподавателей:");
            foreach (var lecturer in lecturers)
            {
                var lecturerEntries = schedule.Where(e => e.LecturerId == lecturer.Id)
                                             .GroupBy(e => e.Date.Date)
                                             .ToList();

                int totalLessons = lecturerEntries.Sum(g => g.Count());
                if (totalLessons == 0) continue;

                double idealLessonsPerDay = (double)totalLessons / totalDays;
                double variance = 0;

                for (int day = 0; day < totalDays; day++)
                {
                    DateTime date = schedule.Min(e => e.Date.Date).AddDays(day);
                    int lessonsOnDay = lecturerEntries.FirstOrDefault(g => g.Key == date)?.Count() ?? 0;
                    variance += Math.Pow(lessonsOnDay - idealLessonsPerDay, 2);
                }

                variance /= totalDays;
                Console.WriteLine($"Преподаватель {lecturer.Name}: Всего занятий: {totalLessons}, Идеально в день: {idealLessonsPerDay:F1}, Дисперсия: {variance:F2}");
                if (variance > 2)
                {
                    Console.WriteLine($"Нагрузка преподавателя {lecturer.Name} неравномерна! Рекомендуется распределить занятия более равномерно.");
                }
            }
        }

        static void EvaluateAuditoriumUsage(List<ScheduleEntry> schedule, List<Auditorium> auditoriums, Dictionary<int, int> auditoriumUsage)
        {
            Console.WriteLine("\nОценка использования аудиторий:");
            foreach (var usage in auditoriumUsage)
            {
                var auditorium = auditoriums.Find(a => a.Id == usage.Key);
                Console.WriteLine($"Аудитория {auditorium.Number}: Использовано {usage.Value} раз");
            }

            int totalLessons = schedule.Count;
            double idealUsagePerAuditorium = (double)totalLessons / auditoriums.Count;
            double variance = 0;

            foreach (var auditorium in auditoriums)
            {
                int usageCount = auditoriumUsage.ContainsKey(auditorium.Id) ? auditoriumUsage[auditorium.Id] : 0;
                variance += Math.Pow(usageCount - idealUsagePerAuditorium, 2);
            }

            variance /= auditoriums.Count;
            Console.WriteLine($"Идеальное использование на аудиторию: {idealUsagePerAuditorium:F1}, Дисперсия: {variance:F2}");
            if (variance > 2)
            {
                Console.WriteLine("Использование аудиторий неравномерно! Рекомендуется распределить занятия по аудиториям более равномерно.");
            }
        }

        static void Main(string[] args)
        {
            string excelPath = @"schedule_data.xlsx";

            List<Department> departments;
            List<ConsoleApp1.Stream> streams;
            List<Lecturer> lecturers;
            List<Auditorium> auditoriums;
            List<Group> groups;
            List<Subject> subjects;

            try
            {
                using (var workbook = new XLWorkbook(excelPath))
                {
                    departments = DataLoader.LoadDepartments(workbook.Worksheet("Departments"));
                    groups = DataLoader.LoadGroups(workbook.Worksheet("Groups"));
                    streams = DataLoader.LoadStreams(workbook.Worksheet("Streams"), groups);
                    lecturers = DataLoader.LoadLecturers(workbook.Worksheet("Lecturers"));
                    auditoriums = DataLoader.LoadAuditoriums(workbook.Worksheet("Auditoriums"));
                    subjects = DataLoader.LoadSubjects(workbook.Worksheet("Subjects"));
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка загрузки данных: {ex.Message}");
                return;
            }

            Console.WriteLine($"Загружено: Departments={departments.Count}, Streams={streams.Count}, Lecturers={lecturers.Count}, Auditoriums={auditoriums.Count}, Groups={groups.Count}, Subjects={subjects.Count}");

            var slots = new List<(DateTime Date, TimeSpan Start, TimeSpan End)>();
            DateTime startDate = new DateTime(2025, 4, 21);
            TimeSpan[] slotStartsFull = new TimeSpan[]
            {
        new TimeSpan(9, 0, 0),
        new TimeSpan(10, 50, 0),
        new TimeSpan(12, 30, 0),
        new TimeSpan(14, 55, 0),
        new TimeSpan(16, 45, 0),
        new TimeSpan(18, 30, 0),
        new TimeSpan(20, 5, 0)
            };
            TimeSpan[] slotStartsSaturday = new TimeSpan[]
            {
        new TimeSpan(9, 0, 0),
        new TimeSpan(10, 50, 0),
        new TimeSpan(12, 30, 0),
        new TimeSpan(14, 55, 0)
            };

            for (int day = 0; day < 6; day++)
            {
                DateTime date = startDate.AddDays(day);
                var slotStarts = (day == 5) ? slotStartsSaturday : slotStartsFull;
                foreach (var start in slotStarts)
                {
                    var end = start.Add(new TimeSpan(1, 30, 0));
                    slots.Add((date, start, end));
                }
            }

            var sortedSubjects = subjects.OrderBy(s => s.FreedomScore).ToList();
            Console.WriteLine("Отсортированные предметы по FreedomScore:");
            foreach (var subject in sortedSubjects)
            {
                Console.WriteLine($"Subject: {subject.Name}, FreedomScore={subject.FreedomScore}, WeeklyFrequency={subject.WeeklyFrequency}, AvailableAuditoriums=[{string.Join(",", subject.AvailableAuditoriums)}]");
            }

            var schedule = new List<ScheduleEntry>();
            int entryId = 1;
            Random random = new Random();
            var auditoriumUsage = auditoriums.ToDictionary(a => a.Id, a => 0); // Отслеживание использования аудиторий

            foreach (var subject in sortedSubjects)
            {
                int placed = 0;
                var lecturer = lecturers.Find(l => l.Id == subject.LecturerId);
                List<Group> targetGroups;

                if (subject.IsForStream)
                {
                    var stream = streams.Find(s => s.Groups.Any(g => g.Id == subject.GroupId));
                    targetGroups = stream?.Groups ?? new List<Group>();
                }
                else
                {
                    targetGroups = new List<Group> { groups.Find(g => g.Id == subject.GroupId) };
                }

                if (lecturer == null)
                {
                    Console.WriteLine($"Не найден преподаватель с Id={subject.LecturerId} для предмета {subject.Name}");
                    continue;
                }
                if (!targetGroups.Any())
                {
                    Console.WriteLine($"Не найдены группы для предмета {subject.Name}");
                    continue;
                }

                Console.WriteLine($"Планирование предмета: {subject.Name}, WeeklyFrequency={subject.WeeklyFrequency}, Groups={string.Join(",", targetGroups.Select(g => g.Name))}");

                var availableSlots = slots.Where(slot =>
                {
                    bool isEvenWeek = IsEvenWeek(slot.Date);
                    if (subject.WeekType == "четная" && !isEvenWeek) return false;
                    if (subject.WeekType == "нечетная" && isEvenWeek) return false;
                    if (!lecturer.IsAvailable(slot.Date, slot.Start, slot.End)) return false;
                    if (!targetGroups.All(g => g.HasFreeSlot(slot.Date, slot.Start, slot.End))) return false;
                    if (lecturer.CurrentHours + (subject.Duration / 60) > lecturer.MaxHoursPerWeek) return false;
                    return true;
                }).ToList();

                var slotsOrdered = availableSlots
                    .Select(slot => new
                    {
                        Slot = slot,
                        LessonsOnDay = targetGroups.Sum(g => schedule.Count(e => e.GroupIds.Contains(g.Id) && e.Date.Date == slot.Date.Date)),
                        PreviousLessonEnd = targetGroups.Select(g => schedule.Where(e => e.GroupIds.Contains(g.Id) && e.Date.Date == slot.Date.Date && e.EndTime <= slot.Start)
                                                                    .OrderByDescending(e => e.EndTime)
                                                                    .FirstOrDefault()?.EndTime)
                                                    .Where(t => t.HasValue)
                                                    .DefaultIfEmpty(null)
                                                    .FirstOrDefault(),
                        NextLessonStart = targetGroups.Select(g => schedule.Where(e => e.GroupIds.Contains(g.Id) && e.Date.Date == slot.Date.Date && e.Time >= slot.End)
                                                                    .OrderBy(e => e.Time)
                                                                    .FirstOrDefault()?.Time)
                                                    .Where(t => t.HasValue)
                                                    .DefaultIfEmpty(null)
                                                    .FirstOrDefault()
                    })
                    .OrderBy(x => x.LessonsOnDay <= 2) // Предпочтение дням с меньшей нагрузкой
                    .ThenBy(x => x.PreviousLessonEnd.HasValue ? (x.Slot.Start - x.PreviousLessonEnd.Value).TotalHours : double.MaxValue) // Минимизировать окно после предыдущего занятия
                    .ThenBy(x => x.NextLessonStart.HasValue ? (x.NextLessonStart.Value - x.Slot.End).TotalHours : double.MaxValue) // Минимизировать окно до следующего занятия
                    .ThenBy(x => x.Slot.Date)
                    .ThenBy(x => x.Slot.Start) // Ранние слоты предпочтительнее
                    .Select(x => x.Slot)
                    .ToList();

                foreach (var slot in slotsOrdered)
                {
                    if (placed >= subject.WeeklyFrequency) break;

                    // Сортировка аудиторий по количеству использований (меньше — лучше)
                    var sortedAuditoriums = subject.AvailableAuditoriums
                        .Select(audId => new { AuditoriumId = audId, Usage = auditoriumUsage[audId] })
                        .OrderBy(a => a.Usage)
                        .ThenBy(_ => random.Next())
                        .Select(a => a.AuditoriumId)
                        .ToList();

                    bool auditoriumFound = false;

                    foreach (var auditoriumId in sortedAuditoriums)
                    {
                        var auditorium = auditoriums.Find(a => a.Id == auditoriumId);
                        if (auditorium == null) continue;

                        if (!subject.AvailableAuditoriums.Contains(auditoriumId)) continue;

                        if (!auditorium.IsAvailable(slot.Date, slot.Start, slot.End)) continue;

                        int requiredCapacity = targetGroups.Sum(g => g.StudentsCount);
                        if (auditorium.Capacity < requiredCapacity) continue;

                        if (subject.LessonType == LessonType.Лаборатория && auditorium.Type != "лаборатория") continue;
                        if (subject.LessonType != LessonType.Лаборатория && auditorium.Type != "лекционная") continue;

                        var entry = new ScheduleEntry
                        {
                            Id = entryId++,
                            GroupId = subject.GroupId,
                            GroupIds = targetGroups.Select(g => g.Id).ToList(),
                            SubjectId = subject.Id,
                            LecturerId = subject.LecturerId,
                            AuditoriumId = auditoriumId,
                            Date = slot.Date,
                            Time = slot.Start,
                            EndTime = slot.End
                        };

                        lecturer.BusyTimes.Add((slot.Date, slot.Start, slot.End));
                        foreach (var group in targetGroups)
                        {
                            group.BusyTimes.Add((slot.Date, slot.Start, slot.End));
                        }
                        auditorium.BusyTimes.Add((slot.Date, slot.Start, slot.End));
                        lecturer.CurrentHours += subject.Duration / 60;

                        auditoriumUsage[auditoriumId]++; // Обновляем счётчик использования аудитории

                        schedule.Add(entry);
                        placed++;
                        auditoriumFound = true;
                        Console.WriteLine($"Размещено занятие: {subject.Name} в {slot.Date:dd.MM.yyyy} {slot.Start:hh\\:mm}-{slot.End:hh\\:mm}, Аудитория: {auditorium.Number}, Группы: {string.Join(",", targetGroups.Select(g => g.Name))}");
                        break;
                    }

                    if (!auditoriumFound)
                    {
                        Console.WriteLine($"Не удалось найти свободную аудиторию для {subject.Name} в слоте {slot.Date:dd.MM.yyyy} {slot.Start:hh\\:mm}-{slot.End:hh\\:mm}");
                    }
                }

                if (placed < subject.WeeklyFrequency)
                {
                    Console.WriteLine($"Не удалось разместить все занятия для {subject.Name} (размещено {placed} из {subject.WeeklyFrequency})");
                }
            }

            Console.WriteLine("\nПроверка расписания на конфликты:");
            bool isValid = true;

            if (!CheckScheduleConflicts(schedule))
            {
                isValid = false;
                Console.WriteLine("Обнаружены конфликты в расписании!");
            }

            if (!CheckGroupDailyLimit(schedule, groups))
            {
                isValid = false;
                Console.WriteLine("Обнаружены нарушения дневного лимита занятий для групп!");
            }

            if (!CheckGroupBreakTime(schedule, groups))
            {
                isValid = false;
                Console.WriteLine("Обнаружены нарушения минимального перерыва между занятиями!");
            }

            if (isValid)
            {
                Console.WriteLine("Расписание прошло все проверки!");
            }

            TimeSpan maxAcceptableWindow = TimeSpan.FromHours(1);
            int totalDays = 6;

            EvaluateGroupWindows(schedule, groups, maxAcceptableWindow);
            EvaluateLecturerWindows(schedule, lecturers, maxAcceptableWindow);
            EvaluateGroupLoadDistribution(schedule, groups, totalDays);
            EvaluateLecturerLoadDistribution(schedule, lecturers, totalDays);
            EvaluateAuditoriumUsage(schedule, auditoriums, auditoriumUsage);

            Console.WriteLine("\nРасписание:");
            if (schedule.Count == 0)
            {
                Console.WriteLine("Расписание пусто.");
            }
            else
            {
                foreach (var entry in schedule.OrderBy(e => e.Date).ThenBy(e => e.Time))
                {
                    var subject = subjects.Find(s => s.Id == entry.SubjectId);
                    var lecturer = lecturers.Find(l => l.Id == entry.LecturerId);
                    var auditorium = auditoriums.Find(a => a.Id == entry.AuditoriumId);
                    var groupNames = entry.GroupIds.Select(gId => groups.Find(g => g.Id == gId).Name);
                    Console.WriteLine($"{entry.Date:dd.MM.yyyy} {entry.Time:hh\\:mm}-{entry.EndTime:hh\\:mm}: " +
                                      $"{subject.Name} (Группы: {string.Join(",", groupNames)}, Преподаватель: {lecturer.Name}, Аудитория: {auditorium.Number})");
                }
            }

            Console.WriteLine("Нажмите любую клавишу, чтобы закрыть...");
            Console.ReadKey();
        }
    }
    }
