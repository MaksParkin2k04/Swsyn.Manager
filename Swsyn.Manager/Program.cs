using Swsyn.Manager.ModelConfiguration;
using Microsoft.Extensions.Configuration;

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

            //Launch();
        }
        static void GeneratingProjectInclude(AppSettings settings)
        {
        
            if (settings.Include == null)
            {
                Console.WriteLine("There are no projects specified (Проекты не указаны в include)");
            }

            else if(settings.Include != null)
            {
                Console.WriteLine("Генерируем отчеты для проектов из include");
                Console.WriteLine($"Include:{string.Join("\n\t", settings.Include.Select(i => $"'{i}'"))}");
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
