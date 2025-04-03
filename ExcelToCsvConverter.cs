using System;
using System.IO;
using System.Text;
using Syncfusion.XlsIO;

public static class ExcelToCsvConverter
{
    /// <summary>
    /// Converts an Excel file to CSV format
    /// </summary>
    /// <typeparam name="T">Type parameter (not used in this implementation but kept for signature consistency)</typeparam>
    /// <param name="ExcelFileData">The Excel file as a byte array</param>
    /// <returns>The CSV data as a byte array</returns>
    public static byte[] CreateCsvFromExcel<T>(byte[] ExcelFileData)
    {
        // Initialize ExcelEngine
        using (ExcelEngine excelEngine = new ExcelEngine())
        {
            // Set application version
            IApplication application = excelEngine.Excel;
            application.DefaultVersion = ExcelVersion.Excel2016;

            // Create a workbook from the Excel data
            using (MemoryStream inputStream = new MemoryStream(ExcelFileData))
            {
                // Load the Excel document
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                
                // Access the first worksheet
                IWorksheet worksheet = workbook.Worksheets[0];
                
                // Create a memory stream to store CSV data
                using (MemoryStream outputStream = new MemoryStream())
                {
                    // Use a StreamWriter to write CSV data
                    using (StreamWriter writer = new StreamWriter(outputStream, Encoding.UTF8, 1024, true))
                    {
                        // Get used range of the worksheet
                        IRange usedRange = worksheet.UsedRange;
                        int rowCount = usedRange.LastRow;
                        int colCount = usedRange.LastColumn;
                        
                        // Process each row in the Excel file
                        for (int row = 1; row <= rowCount; row++)
                        {
                            StringBuilder rowData = new StringBuilder();
                            
                            // Process each column in the current row
                            for (int col = 1; col <= colCount; col++)
                            {
                                // Get cell value
                                string cellValue = worksheet[row, col].DisplayText?.ToString() ?? string.Empty;
                                
                                // Escape quotes and special characters
                                cellValue = cellValue.Replace("\"", "\"\"");
                                
                                // Add quotes around the value if it contains comma, newline, or quotes
                                if (cellValue.Contains(",") || cellValue.Contains("\"") || 
                                    cellValue.Contains("\n") || cellValue.Contains("\r"))
                                {
                                    cellValue = $"\"{cellValue}\"";
                                }
                                
                                // Append to row data
                                rowData.Append(cellValue);
                                
                                // Add comma separator if not the last column
                                if (col < colCount)
                                {
                                    rowData.Append(",");
                                }
                            }
                            
                            // Write row to the CSV
                            writer.WriteLine(rowData.ToString());
                        }
                        
                        // Flush the writer
                        writer.Flush();
                        
                        // Get the CSV data as a byte array
                        outputStream.Position = 0;
                        return outputStream.ToArray();
                    }
                }
            }
        }
    }
} 