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

                // Получаем названия проектов для которых будут генерироваться отчеты.
                string[]? projectNames = GetProjectNames(settings);

                DateTime[] specifyPeriods = ProjectDataProcessing();

                Console.WriteLine(projectNames);
              //  Launch(settings);
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(ex.ToString());
            }
        }

        private static DateTime[] CurrentWeek()
        {
           DateTime today = DateTime.Today;

            int daysSinceMonday = ((int)today.DayOfWeek - 1 + 7) % 7;
            DateTime monday = today.AddDays(-daysSinceMonday);

            DateTime sunday = monday.AddDays(6);

            Console.WriteLine($"Создаем отчет для текущей недели: с {monday:dd.MM.yyyy} по {sunday:dd.MM.yyyy}");

            return new DateTime[] { monday };
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

        private static DateTime[] ProjectDataProcessing()
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

                            return new DateTime[] { data } ;
                        }
                        else
                        {
                            Console.WriteLine("Введен не понедельник");
                        }
                        break;
                    //N @ key
                    case ConsoleKey.N:
                        DateTime today = DateTime.Today;

                        int daysSinceMonday = ((int)today.DayOfWeek - 1 + 7) % 7;
                        DateTime monday = today.AddDays(-daysSinceMonday);

                        DateTime sunday = monday.AddDays(6);

                        Console.WriteLine($"Создаем отчет для текущей недели: с {monday:dd.MM.yyyy} по {sunday:dd.MM.yyyy}");

                        return new DateTime[] { monday };
                        break;
                    //Enter @ Key
                    case ConsoleKey.Enter:
                        CurrentWeek(); //Исправить 
                        break;
                }
            }
            while (choice != ConsoleKey.Y && choice != ConsoleKey.N && choice != ConsoleKey.Enter);

            return null;
        }

        private static string[]? GetProjectNames(AppSettings settings)
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

                        Console.WriteLine("Type project names (comma separated):");
                        // Названия проектов через запятую
                        string? projects = Console.ReadLine();
                        // Возвращаем названия проектов, введенных пользователем
                        return projects != null && projects != string.Empty ? projects.Split(',') : null;

                    //N @ key
                    case ConsoleKey.N:

                        Console.Clear();

                        if (settings.Include == null)
                        {
                          //  Console.WriteLine("There are no projects specified (Проекты не указаны в include)");
                            string[] error = { "There are no projects specified (Проекты не указаны в include)" };
                            return error;
                        }

                        else if (settings.Include != null)
                        {
                            Console.WriteLine("Генерируем отчеты для проектов из include");
                            Console.WriteLine($"Include:{string.Join("\n\t", settings.Include.Select(i => $"'{i}'"))}");
                            return settings.Include;

                            ProjectDataProcessing();
                        }
                        break;
                    //Enter @ Key
                    case ConsoleKey.Enter:

                        Console.Clear();

                        if (settings.Include == null)
                        {
                            Console.WriteLine("There are no projects specified (Проекты не указаны в include)");
                        }

                        else if (settings.Include != null)
                        {
                            Console.WriteLine("Генерируем отчеты для проектов из include");
                            Console.WriteLine($"Include:{string.Join("\n\t", settings.Include.Select(i => $"'{i}'"))}");
                            return settings.Include;

                            ProjectDataProcessing();
                        }

                        break;
                }
            }
            while (choice != ConsoleKey.Y && choice != ConsoleKey.N && choice != ConsoleKey.Enter);

            return null;

            
        }
    }
}
