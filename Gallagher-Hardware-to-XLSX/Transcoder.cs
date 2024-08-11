using System;
using System.IO;
using OfficeOpenXml;

public partial class Transcoder
{
    public bool convertDoors(string inputFilename, string outputFilename)
    {
        try
        {
            int count = 0;
            if (!File.Exists(inputFilename))
            {
                Console.WriteLine("Failed to transcode data. Input file does not exist.");
                return false;
            }

            using (StreamReader reader = new StreamReader(inputFilename))
            using (ExcelPackage package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Doors");

                worksheet.Cells[1, 1].Value = "Door Name";
                worksheet.Cells[1, 2].Value = "Open sensor";
                worksheet.Cells[1, 3].Value = "Lock sensor";
                worksheet.Cells[1, 4].Value = "Remote release";
                worksheet.Cells[1, 5].Value = "Emergency release";
                worksheet.Cells[1, 6].Value = "Unlock output";
                worksheet.Cells[1, 7].Value = "Entry Reader One";
                worksheet.Cells[1, 8].Value = "Entry Reader Two";
                worksheet.Cells[1, 9].Value = "Exit Reader One";
                worksheet.Cells[1, 10].Value = "Exit Reader Two";

                int row = 2;
                string? line;
                string doorName = string.Empty;
                string openSensor = "N/A";
                string lockSensor = "N/A";
                string remoteRelease = "N/A";
                string emergencyRelease = "N/A";
                string unlockOutput = "N/A";
                string entryReader1 = "N/A";
                string entryReader2 = "N/A";
                string exitReader1 = "N/A";
                string exitReader2 = "N/A";

                while ((line = reader.ReadLine()) != null)
                {
                    line = line.Trim();

                    if (line.StartsWith("Type,") && !line.Contains("Door"))
                    {
                        Console.WriteLine("Failed to transcode data. File contains elements tht are not of Type Door. Please refer to read me for instructions on exporting reports.");
                        return false;
                    }

                    if (line.StartsWith("Type,Door"))
                    {
                        count++;
                        // Reset fields for new door entry
                        doorName = string.Empty;
                        openSensor = "N/A";
                        lockSensor = "N/A";
                        remoteRelease = "N/A";
                        emergencyRelease = "N/A";
                        unlockOutput = "N/A";
                        entryReader1 = "N/A";
                        entryReader2 = "N/A";
                        exitReader1 = "N/A";
                        exitReader2 = "N/A";
                    }
                    else if (line.StartsWith("Name,"))
                    {
                        doorName = line.Substring(5).Trim();
                    }
                    else if (line.StartsWith("Open sensor"))
                    {
                        openSensor = "Not Tested";
                    }
                    else if (line.StartsWith("Lock sensor"))
                    {
                        lockSensor = "Not Tested";
                    }
                    else if (line.StartsWith("Remote release"))
                    {
                        remoteRelease = "Not Tested";
                    }
                    else if (line.StartsWith("Emergency release"))
                    {
                        emergencyRelease = "Not Tested";
                    }
                    else if (line.StartsWith("Unlock output"))
                    {
                        unlockOutput = "Not Tested";
                    }
                    else if (line.StartsWith("Entry Reader 1"))
                    {
                        entryReader1 = "Not Tested";
                    }
                    else if (line.StartsWith("Entry Reader 2"))
                    {
                        entryReader2 = "Not Tested";
                    }
                    else if (line.StartsWith("Exit Reader 1"))
                    {
                        exitReader1 = "Not Tested";
                    }
                    else if (line.StartsWith("Exit Reader 2"))
                    {
                        exitReader2 = "Not Tested";
                    }
                    else if (line.Contains("New Page"))
                    {
                        // End of a door entry, write the row to the Excel file
                        worksheet.Cells[row, 1].Value = doorName;
                        worksheet.Cells[row, 2].Value = openSensor;
                        worksheet.Cells[row, 3].Value = lockSensor;
                        worksheet.Cells[row, 4].Value = remoteRelease;
                        worksheet.Cells[row, 5].Value = emergencyRelease;
                        worksheet.Cells[row, 6].Value = unlockOutput;
                        worksheet.Cells[row, 7].Value = entryReader1;
                        worksheet.Cells[row, 8].Value = entryReader2;
                        worksheet.Cells[row, 9].Value = exitReader1;
                        worksheet.Cells[row, 10].Value = exitReader2;

                        row++;

                        // Debugging line
                        //Console.WriteLine($"{doorName}, {openSensor}, {lockSensor}, {remoteRelease}, {emergencyRelease}, {unlockOutput}, {entryReader1}, {entryReader2}, {exitReader1}, {exitReader2}");
                    }
                }

                // Write the last door's details if there's no "New Page" after it
                if (!string.IsNullOrEmpty(doorName))
                {
                    worksheet.Cells[row, 1].Value = doorName;
                    worksheet.Cells[row, 2].Value = openSensor;
                    worksheet.Cells[row, 3].Value = lockSensor;
                    worksheet.Cells[row, 4].Value = remoteRelease;
                    worksheet.Cells[row, 5].Value = emergencyRelease;
                    worksheet.Cells[row, 6].Value = unlockOutput;
                    worksheet.Cells[row, 7].Value = entryReader1;
                    worksheet.Cells[row, 8].Value = entryReader2;
                    worksheet.Cells[row, 9].Value = exitReader1;
                    worksheet.Cells[row, 10].Value = exitReader2;

                    // Debugging line
                    //Console.WriteLine($"{doorName}, {openSensor}, {lockSensor}, {remoteRelease}, {emergencyRelease}, {unlockOutput}, {entryReader1}, {entryReader2}, {exitReader1}, {exitReader2}");
                }

                FileInfo excelFile = new FileInfo(outputFilename);
                package.SaveAs(excelFile);
            }

            Console.WriteLine($"Successfully transcoded {count} doors to {outputFilename}");
            return true;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to transcode data. An error occurred: {ex.Message}");
            return false;
        }
    }
}
