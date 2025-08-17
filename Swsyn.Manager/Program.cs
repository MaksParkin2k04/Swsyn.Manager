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

                if (settings == null) { throw new Exception("Settings Error"); }
               // View(settings);
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(ex.ToString());
            }

           

            Launch();
        }

        //static void View(AppSettings settings)
        //{
        //    Console.WriteLine($"SourcePath: '{settings.SourcePath}'.");
        //    Console.WriteLine();

        //    Console.WriteLine("Projects:");
        //    Console.WriteLine();

        //    foreach (string projectName in settings.Projects.Keys)
        //    {
        //        ProjectOption projectOption = settings.Projects[projectName];
        //        Console.WriteLine($"Project Name: '{projectName}'.");
        //        Console.WriteLine($"Target Path: '{projectOption.TargetPath}'.");
        //        Console.WriteLine($"Contractor: '{projectOption.Contractor}'.");
        //        Console.WriteLine($"Customer: '{projectOption.Customer}'.");
        //        Console.WriteLine();
        //    }

        //    Console.WriteLine($"Include: {string.Join(", ", settings.Include.Select(i => $"'{i}'"))}");
        //}

        static void Launch()
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
                        Console.WriteLine("Выбор нужных проектов");
                        break;
                    //N @ key
                    case ConsoleKey.N:
                        Console.Clear();
                        Console.WriteLine("Генерируем проекты из include");
                        break;
                    //Enter @ Key
                    case ConsoleKey.Enter:
                        Console.Clear();
                        Console.WriteLine("Генерируем проекты из include");
                        break;
                }
            }
            while (choice != ConsoleKey.Y && choice != ConsoleKey.N && choice != ConsoleKey.Enter);
        }
    }
}
