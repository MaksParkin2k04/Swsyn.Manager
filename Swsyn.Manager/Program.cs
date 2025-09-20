using Swsyn.Manager.ModelConfiguration;
using Microsoft.Extensions.Configuration;
using System.Globalization;
using DocumentFormat.OpenXml.Wordprocessing;
using Timesheets;
using Timesheets.OpenXml;

namespace Swsyn.Manager
{
    public class Program
    {
        static void Main(string[] args)
        {
            try
            {
                IConfigurationRoot config = new ConfigurationBuilder().AddJsonFile("appsettings.json").Build();
                AppSettings? settings = config.Get<AppSettings>();
                ITimesheetRepository timesheetRepository = new TimesheetRepository(settings.SourcePath);

                if (settings == null)
                {
                    throw new Exception("Settings Error");
                }

                string[]? projectNames = GetProjectNames(settings);

                DateTime[] specifyPeriods = ProjectDataProcessing();

                GenerateTimesheet(timesheetRepository, projectNames, specifyPeriods, settings);
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(ex.ToString());
            }

            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }


        private static DateTime[] ProjectDataProcessing()
        {
            Console.WriteLine("Would you like to specify period? [y/n]");

            string? choice = Console.ReadLine();

            if (choice == "y")
            {
                Console.WriteLine("Type week Monday dates (e.g. 11.08.2025, 18.08.2025):");

                string targetDateFormat = "dd.MM.yyyy";
                DateTime data;

                string enteredDateString = Console.ReadLine();

                data = DateTime.ParseExact(enteredDateString, targetDateFormat, CultureInfo.InvariantCulture);

                if (data.DayOfWeek == DayOfWeek.Monday)
                {

                    return new DateTime[] { data };
                }
                else
                {
                    Console.WriteLine("Введен не понедельник");
                }
            }

            else if (choice == "n" || string.IsNullOrEmpty(choice))
            {
                DateTime today = DateTime.Today;

                int daysSinceMonday = ((int)today.DayOfWeek - 1 + 7) % 7;
                DateTime monday = today.AddDays(-daysSinceMonday);

                DateTime sunday = monday.AddDays(6);

                Console.WriteLine($"Создаем отчет для текущей недели: с {monday:dd.MM.yyyy} по {sunday:dd.MM.yyyy}");

                return new DateTime[] { monday };
            }

            return null;
        }

        private static string[]? GetProjectNames(AppSettings settings)
        {
            Console.WriteLine("Would you like to specify projects? [y/n]");

            string? choice = Console.ReadLine();

            if (choice == "y")
            {
                Console.WriteLine("Type project names (comma separated):");
                // Названия проектов через запятую
                string? projects = Console.ReadLine();
                // Возвращаем названия проектов, введенных пользователем
                return projects != null && projects != string.Empty ? projects.Split(',') : null;
            }

            else if (choice == "n" || string.IsNullOrEmpty(choice))
                {
                if (settings.Include == null)
                {
                    //  Console.WriteLine("There are no projects specified (Проекты не указаны в include)");
                    string[] error = { "There are no projects specified" };
                    return error;
                }

                else if (settings.Include != null)
                {
                    Console.WriteLine("Генерируем отчеты для проектов из include");
                    Console.WriteLine($"Include:{string.Join("\n\t", settings.Include.Select(i => $"'{i}'"))}");
                    return settings.Include;

                }

            }
            return null;
        }
        private static void GenerateTimesheet(ITimesheetRepository timesheetRepository, string[] projectNames, DateTime[] mondayDates, AppSettings settings)
        {
            foreach (DateTime mondayDate in mondayDates)
            {
                // Получаем отчеты для недели
                Timesheet[] tables = timesheetRepository.GetTimesheet(mondayDate);

                foreach (string projectName in projectNames)
                {
                    try
                    {
                        // Когда начинает генерироваться отчет выводится сообщение:
                        Console.WriteLine($"[{DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss")}] Creating timesheet [{projectName} {mondayDate.ToString("dd.MM.yyyy")}]…");

                        // Отфильтровываем задачи относящиеся к текущему проекту
                        List<ProjectTask> projectTasks = new List<ProjectTask>();
                        foreach (Timesheet table in tables)
                        {
                            ProjectTask[] tasks = table.Tasks.Where(t => t.Project == projectName).ToArray();
                            projectTasks.AddRange(tasks);
                        }

                        // Получаем настройки проекта
                        ProjectOption option = settings.Projects[projectName];

                        // Создаем сводный табель учета рабочего времени для проекта
                        Timesheet writeTable = new Timesheet(mondayDate, projectTasks);

                        // Сохраняем сводный табель учета рабочего времени для проекта
                        timesheetRepository.WriteTimesheet(option.TargetPath, projectName, writeTable, option.Contractor, option.Customer);

                        // Если процесс завершился успешно
                        Console.WriteLine($"[{DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss")}] Timesheet [{projectName} {mondayDate.ToString("dd.MM.yyyy")}] was generated successfully.");
                    }
                    catch (Exception exception)
                    {
                        // Если произошла какая-либо другая непредвиденная ошибка
                        throw new Exception($"[{DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss")}] Failed to create timesheet [{projectName} {mondayDate.ToString("dd.MM.yyyy")}].", exception);
                    }
                }
            }
        }
    }
}
