namespace Swsyn.Manager
{
    internal class Program
    {
        static void Main(string[] args)
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
