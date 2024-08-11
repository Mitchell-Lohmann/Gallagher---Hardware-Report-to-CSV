namespace Gallagher_Hardware_to_XLSX
{
    internal class Menu
    {

        //Display menu
        public void DisplayMenu()
        {

            Console.WriteLine("\nConvert Gallagher Hardware Report to .xlsx");
            Console.WriteLine("--------------------------------------------------");
            Console.WriteLine("1. Convert Doors");
            Console.WriteLine("2. Convert Inputs");
            Console.WriteLine("0. Exit");
            Console.WriteLine("--------------------------------------------------");
            Console.Write("Enter your choice (1-2) or press 0 to exit: ");
        }

        //Get the user's choice
        public int GetUserChoice()
        {
            int choice;
            while (true)
            {
                try
                {
                    if (int.TryParse(Console.ReadLine(), out choice))
                    {
                        if (choice == 0)
                        {
                            Console.WriteLine("Thanks for using this system!");
                            Environment.Exit(0);
                        }
                        else if (choice >= 1 || choice <= 2)
                        {
                            return choice;
                        }
                        else
                        {
                            Console.WriteLine();
                            Console.WriteLine("Invalid input. Please enter a valid choice (0-2).");
                        }
                    }
                    else
                    {
                        Console.WriteLine();
                        Console.WriteLine("Invalid input. Please enter a valid choice (0-2).");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Invalid input: {ex.Message}");
                    Console.WriteLine();
                }
            }
        }

        //Perform an action selected by the user
        public void PerformUserAction(int choice, ref Transcoder transcoder)
        {
            try
            {
                switch (choice)
                {
                    case 1:
                        string? inputFilename;
                        string? outputFilename;
                        string currentDirectory = Environment.CurrentDirectory;
                        string? formattedInputFilename;
                        string? FormattedOutputFilename;

                        Console.Write("Enter the relative path of the file to read from (including file extensions): ");
                        inputFilename = Console.ReadLine();

                        Console.Write("Enter the relative path of the generated .xlsx file (including file extensions): ");
                        outputFilename = Console.ReadLine();

                        if (inputFilename == null || outputFilename == null)
                        {
                            Console.WriteLine("Invalid input file name or output file name.");
                        }
                        else
                        {
                            formattedInputFilename = Path.Combine(currentDirectory, inputFilename);
                            FormattedOutputFilename = Path.Combine(currentDirectory, outputFilename);

                            transcoder.convertDoors(formattedInputFilename, FormattedOutputFilename);
                        }
                        break;

                    case 2:
                        Console.WriteLine("Conversion of Inputs not yet implemented.");
                        break;

                    default:
                        Console.WriteLine("Invalid selection");
                        break;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An exception occurred: {ex.Message}");
            }

        }
    }
}
