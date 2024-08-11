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
   git clone https://github.com/Mitchell-Lohmann/Gallagher_Hardware_to_XLSX.git
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

1. Enter the relative path of the `.txt` file you want to read from.
2. Enter the relative path of the `.xlsx` file you want to generate.
   
The application will then convert the data and save it to the specified `.xlsx` file.

## Example
If you have a file doors_report.txt with door data and you want to convert it to doors_report.xlsx, you would:

- Choose option 1 to convert doors.
- Enter doors_report.txt when prompted for the input file.
- Enter doors_report.xlsx when prompted for the output file.

## File Format
The `.txt` file should contain door data formatted as follows:

- Each door entry starts with Type,Door.
- Followed by attributes such as Name, Open sensor, Lock sensor, etc.
- Each door entry should end with New Page.

## Notes
- Ensure that your `.txt` file adheres to the expected format for successful conversion.
- The conversion of inputs is not yet implemented but is planned for future updates.

### MIT License

MIT License

Copyright (c) [2024] [Mitchell Lohmann]

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
