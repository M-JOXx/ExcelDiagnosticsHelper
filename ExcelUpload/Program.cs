using ExcelDiagnosticHelper;
using OfficeOpenXml;

namespace ExcelUpload
{
    internal class Program
    {

        static void Main(string[] args)
        {
            string filePath = "C:\\Users\\Adestra\\Documents\\SampleTransactions.xlsx";

            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Error: File not found at path '{filePath}'");
                return;
            }

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            string inputPath = filePath;
            string errorsOut = $"invoices_errors_{DateTime.Now:yyyy-MM-dd_HH-mm-ss}.xlsx";

            IExcelReader reader = new ExcelReader();

            var validators = new IRowValidator<ExcelInvoiceRow>[] { new SampleInvoiceRowValidator() };

            var result = reader.ReadExcel<ExcelInvoiceRow>(
                filePath: inputPath,
                headerRow: 1,
                rowValidators: new[] { new SampleInvoiceRowValidator() }
            );


            Console.WriteLine($"Total Rows: {result.TotalRows}");
            Console.WriteLine($"Valid Rows: {result.ValidRowsCount}");
            Console.WriteLine($"Invalid Rows: {result.InvalidRowsCount}");
            Console.WriteLine($"Cell Errors: {result.TotalCellErrorsCount}");
            Console.WriteLine($"Row Errors: {result.TotalRowErrorsCount}");

            // If you added warnings
            Console.WriteLine($"Cell Warnings: {result.TotalCellWarningsCount}");
            Console.WriteLine($"Row Warnings: {result.TotalRowWarningsCount}");



            // Generate separate Errors sheet file
            result.GenerateErrorsExcelFile(errorsOut);
            Console.WriteLine($"Errors file: {Path.GetFullPath(errorsOut)}");

        }

    }

}
