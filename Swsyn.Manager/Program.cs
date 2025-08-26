using Swsyn.Manager.ModelConfiguration;
using Microsoft.Extensions.Configuration;
using System.Globalization;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Swsyn.Manager
{
    internal class Program
    {
        static void Main(string[] args)
        {
            try
            {
                IConfigurationRoot config = new ConfigurationBuilder().AddJsonFile("appsettings.json").Build();
                AppSettings? settings = config.Get<AppSettings>();

                if (settings == null)
                {
                    throw new Exception("Settings Error");
                }

                Launch(settings);
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(ex.ToString());
            }
        }

        static void CurrentWeek()
        {
            DateTime today = DateTime.Today;

            int daysSinceMonday = ((int)today.DayOfWeek - 1 + 7) % 7;
            DateTime monday = today.AddDays(-daysSinceMonday);

            DateTime sunday = monday.AddDays(6);

            Console.WriteLine($"Создаем отчет для текущей недели: с {monday:dd.MM.yyyy} по {sunday:dd.MM.yyyy}");
        }

        static void InputWeek()
        {
            Console.WriteLine("Выберите дату");

            string targetDateFormat = "dd.MM.yyyy";
            DateTime data;

            Console.WriteLine($"Enter a date in the format {targetDateFormat}");
            string enteredDateString = Console.ReadLine();

            data = DateTime.ParseExact(enteredDateString, targetDateFormat, CultureInfo.InvariantCulture);

            if (data.DayOfWeek == DayOfWeek.Monday)
            {
                Console.WriteLine(data.DayOfWeek);
                Console.WriteLine(data);
                Console.WriteLine("Создаем отчет для недели", data.ToString("dd.MM.yyyy"));
            }
            else
            {
                Console.WriteLine("Введен не понедельник");
            }
        }

        static void ProjectDataProcessing()
        {

            Console.WriteLine("Would you like to specify period? [y/n]");

            ConsoleKey choice;
            do
            {
                choice = Console.ReadKey(true).Key;
                switch (choice)
                {
                    // Y ! key
                    case ConsoleKey.Y:
                        InputWeek();
                        break;
                    //N @ key
                    case ConsoleKey.N:
                        CurrentWeek();
                        break;
                    //Enter @ Key
                    case ConsoleKey.Enter:
                        CurrentWeek();
                        break;
                }
            }
            while (choice != ConsoleKey.Y && choice != ConsoleKey.N && choice != ConsoleKey.Enter);
        }

        static void GeneratingProjectInclude(AppSettings settings)
        {

            if (settings.Include == null)
            {
                Console.WriteLine("There are no projects specified (Проекты не указаны в include)");
            }

            else if (settings.Include != null)
            {
                Console.WriteLine("Генерируем отчеты для проектов из include");
                Console.WriteLine($"Include:{string.Join("\n\t", settings.Include.Select(i => $"'{i}'"))}");

                ProjectDataProcessing();
            }
        }

        static void EnteringProjects()
        {
            Console.WriteLine("Type project names (comma separated) (Введите названия проектов через запятую и нажмите Enter):");
            string input = Console.ReadLine();

            if (string.IsNullOrWhiteSpace(input))
            {
                Console.WriteLine("Вы не ввели ни одного проекта!");
                return;
            }

            string[] projects = input.Split(',')
                                   .Select(project => project.Trim())
                                   .Where(project => !string.IsNullOrEmpty(project))
                                   .ToArray();

            Console.WriteLine("\nСписок проектов для генерации отчетов:");
            foreach (string project in projects)
            {
                Console.WriteLine($"- {project}");
            }

            ProjectDataProcessing();
        }

        static void Launch(AppSettings settings)
        {
            Console.WriteLine("Would you like to specify projects? (Хотели бы вы указать проекты?) [y/n]");

            ConsoleKey choice;
            do
            {
                choice = Console.ReadKey(true).Key;
                switch (choice)
                {
                    // Y ! key
                    case ConsoleKey.Y:
                        Console.Clear();
                        EnteringProjects();
                        break;
                    //N @ key
                    case ConsoleKey.N:
                        Console.Clear();
                        GeneratingProjectInclude(settings);
                        break;
                    //Enter @ Key
                    case ConsoleKey.Enter:
                        Console.Clear();
                        GeneratingProjectInclude(settings);
                        break;
                }
            }
            while (choice != ConsoleKey.Y && choice != ConsoleKey.N && choice != ConsoleKey.Enter);
        }
    }
}
