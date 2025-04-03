using System;
using System.IO;
using System.Text;
using Xunit;
using Syncfusion.XlsIO;

namespace ExcelToCsv.Tests
{
    public class ExcelToCsvConverterTests
    {
        [Fact]
        public void CreateCsvFromExcel_WithValidExcelData_ReturnsCsvData()
        {
            // Arrange
            byte[] excelData = CreateSampleExcelFile();

            // Act
            byte[] csvData = ExcelToCsvConverter.CreateCsvFromExcel<object>(excelData);

            // Assert
            Assert.NotNull(csvData);
            Assert.True(csvData.Length > 0);
            
            string csvContent = Encoding.UTF8.GetString(csvData);
            Assert.Contains("Test1,Test2,Test3", csvContent);
            Assert.Contains("1,2,3", csvContent);
            Assert.Contains("4,5,6", csvContent);
        }

        [Fact]
        public void CreateCsvFromExcel_WithEmptyExcelFile_ReturnsResult()
        {
            // Arrange
            byte[] emptyExcelData = CreateEmptyExcelFile();

            // Act
            byte[] csvData = ExcelToCsvConverter.CreateCsvFromExcel<object>(emptyExcelData);

            // Assert
            Assert.NotNull(csvData);
            // We don't make assumptions about the exact output format,
            // just ensure the method executes without throwing exceptions
        }

        [Fact]
        public void CreateCsvFromExcel_WithSpecialCharacters_EscapesCorrectly()
        {
            // Arrange
            byte[] excelDataWithSpecialChars = CreateExcelFileWithSpecialCharacters();

            // Act
            byte[] csvData = ExcelToCsvConverter.CreateCsvFromExcel<object>(excelDataWithSpecialChars);

            // Assert
            Assert.NotNull(csvData);
            string csvContent = Encoding.UTF8.GetString(csvData);
            Assert.Contains("\"Text with, comma\"", csvContent);
            Assert.Contains("\"Text with \"\"quotes\"\"\"", csvContent);
        }

        [Fact]
        public void CreateCsvFromExcel_WithMultipleSheets_ProcessesFirstSheetOnly()
        {
            // Arrange
            byte[] multiSheetExcelData = CreateMultiSheetExcelFile();

            // Act
            byte[] csvData = ExcelToCsvConverter.CreateCsvFromExcel<object>(multiSheetExcelData);

            // Assert
            Assert.NotNull(csvData);
            string csvContent = Encoding.UTF8.GetString(csvData);
            Assert.Contains("Sheet1Data", csvContent);
            Assert.DoesNotContain("Sheet2Data", csvContent);
        }

        #region Helper Methods

        private byte[] CreateSampleExcelFile()
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Excel2016;
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet worksheet = workbook.Worksheets[0];

                // Add headers
                worksheet[1, 1].Text = "Test1";
                worksheet[1, 2].Text = "Test2";
                worksheet[1, 3].Text = "Test3";

                // Add data row 1
                worksheet[2, 1].Number = 1;
                worksheet[2, 2].Number = 2;
                worksheet[2, 3].Number = 3;

                // Add data row 2
                worksheet[3, 1].Number = 4;
                worksheet[3, 2].Number = 5;
                worksheet[3, 3].Number = 6;

                using (MemoryStream stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    return stream.ToArray();
                }
            }
        }

        private byte[] CreateEmptyExcelFile()
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Excel2016;
                IWorkbook workbook = application.Workbooks.Create(1);
                
                using (MemoryStream stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    return stream.ToArray();
                }
            }
        }

        private byte[] CreateExcelFileWithSpecialCharacters()
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Excel2016;
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet worksheet = workbook.Worksheets[0];

                // Add headers
                worksheet[1, 1].Text = "Column1";
                worksheet[1, 2].Text = "Column2";

                // Add data with special characters
                worksheet[2, 1].Text = "Text with, comma";
                worksheet[2, 2].Text = "Text with \"quotes\"";

                using (MemoryStream stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    return stream.ToArray();
                }
            }
        }

        private byte[] CreateMultiSheetExcelFile()
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Excel2016;
                IWorkbook workbook = application.Workbooks.Create(2);
                
                IWorksheet sheet1 = workbook.Worksheets[0];
                sheet1.Name = "Sheet1";
                sheet1[1, 1].Text = "Sheet1Data";

                IWorksheet sheet2 = workbook.Worksheets[1];
                sheet2.Name = "Sheet2";
                sheet2[1, 1].Text = "Sheet2Data";

                using (MemoryStream stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    return stream.ToArray();
                }
            }
        }

        #endregion
    }
} 