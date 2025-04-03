using System;
using System.IO;

// Syncfusion Licensing Information:
// 1. Add reference: using Syncfusion.Licensing;
// 2. Register license before using any Syncfusion components:
//    SyncfusionLicenseProvider.RegisterLicense("YOUR_LICENSE_KEY_HERE");
// 3. License should be registered at application startup
// 4. Get your license key from https://www.syncfusion.com/account/downloads
// 5. Store license key in configuration or environment variables for production use

class Program
{
    static void Main(string[] args)
    {
        // TODO: Register Syncfusion license here before using any Syncfusion components
        // SyncfusionLicenseProvider.RegisterLicense("YOUR_LICENSE_KEY_HERE");
        
        Console.WriteLine("Excel to CSV Converter Demo");
        Console.WriteLine("==========================");
        
        try
        {
            // Path to the Excel file
            Console.Write("Enter path to Excel file: ");
            string excelFilePath = Console.ReadLine();
            
            if (string.IsNullOrEmpty(excelFilePath))
            {
                Console.WriteLine("No file path provided. Exiting...");
                return;
            }
            
            // Read the Excel file into a byte array
            byte[] excelData = File.ReadAllBytes(excelFilePath);
            
            // Convert Excel to CSV
            Console.WriteLine("Converting Excel to CSV...");
            byte[] csvData = ExcelToCsvConverter.CreateCsvFromExcel<object>(excelData);
            
            // Save the CSV file
            string csvFilePath = Path.ChangeExtension(excelFilePath, ".csv");
            File.WriteAllBytes(csvFilePath, csvData);
            
            Console.WriteLine($"CSV file saved to: {csvFilePath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
            if (ex.InnerException != null)
            {
                Console.WriteLine($"Inner Exception: {ex.InnerException.Message}");
            }
        }
        
        Console.WriteLine("\nPress any key to exit...");
        Console.ReadKey();
    }
} 