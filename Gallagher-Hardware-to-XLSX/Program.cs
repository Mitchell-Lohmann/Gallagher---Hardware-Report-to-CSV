namespace Gallagher_Hardware_to_XLSX
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Create new instances.
            Menu menu = new Menu();
            Transcoder transcoder = new Transcoder(); 

            // Main program loop. Display menu and get user choice.
            while (true)
            {
                menu.DisplayMenu();
                int choice = menu.GetUserChoice();

                menu.PerformUserAction(choice, ref transcoder);
            }
        }
    }
}