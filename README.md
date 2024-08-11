# Gallagher Hardware Report to XLSX

This project provides a console application that converts Gallagher hardware reports from `.txt` format to an `.xlsx` format. It currently supports the conversion of door data.

## Features

- **Convert Doors**: Converts Gallagher hardware door data from a `.txt` file to an `.xlsx` file.
- **Inputs Conversion**: Placeholder for future input conversion functionality.

## Requirements

- .NET Core or .NET Framework (for running the application)
- EPPlus library (for Excel file handling)

## Getting Started

1. **Clone the Repository**

   ```bash
   git clone https://github.com/yourusername/Gallagher_Hardware_to_XLSX.git
   cd Gallagher_Hardware_to_XLSX
   
2. **Install Dependencies**
   ```bash
   dotnet add package EPPlus
   
3. **Build the Project**
   ```bash
   dotnet build

4. **Run the Application**
   ```bash
   dotnet run

## Usage
When prompted by the application:

1. Enter the relative path of the .txt file you want to read from.
2. Enter the relative path of the .xlsx file you want to generate.
   
The application will then convert the data and save it to the specified .xlsx file.

## Example
If you have a file doors_report.txt with door data and you want to convert it to doors_report.xlsx, you would:

- Choose option 1 to convert doors.
- Enter doors_report.txt when prompted for the input file.
- Enter doors_report.xlsx when prompted for the output file.

## File Format
The .txt file should contain door data formatted as follows:

- Each door entry starts with Type,Door.
- Followed by attributes such as Name, Open sensor, Lock sensor, etc.
- Each door entry should end with New Page.

## Notes
- Ensure that your .txt file adheres to the expected format for successful conversion.
- The conversion of inputs is not yet implemented but is planned for future updates.

## License
This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

