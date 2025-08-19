using OfficeOpenXml;
using System.Threading.Tasks;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using System.Reflection.PortableExecutable;
using ExcelDiagnostic.Core.Contracts;
using ExcelDiagnostic.Core.Templates;
using ExcelDiagnostic.Core.Validation;
using ExcelDiagnostic.Core.Extensions;
using ExcelDiagnostic.Core.Readers;
using Microsoft.Extensions.DependencyInjection.Extensions;

namespace ExcelUpload
{
    internal class Program
    {

        static async Task  Main(string[] args)
        {
            //Notice : For now i add this console appilcation  later i will push to repo new update contains like testing unit using (Xunit) for fast feedback and assertion to tests folder
            //,This is simple console application just for testing purposes .

            //You could Majd Add DI to your application WEB/API like this or inject it as <IExcelReader, ExcelReader>.

            /////////////////////////Option two///////////////////////////////
            string filePath = @"DeskTop\SampleTransactions.xlsx";// or try xlx extension file like .xls
            string errorsOutFileName = "invoices_errors";

            var errorsOut = $"{errorsOutFileName}_{DateTime.Now:yyyy-MM-dd_HH-mm-ss}.xlsx";

            //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using IHost host = Host.CreateDefaultBuilder(args)
                .ConfigureServices((context, services) =>
                {
                    services.AddExcelDiagnostics(); //i mean this line IServiceCollection or single 
                    //services.TryAddScoped<IExcelReader, ExcelReader>();// this line is not needed if you use AddExcelDiagnostics
                })
                .Build();

            using var scope = host.Services.CreateScope();
            var provider = scope.ServiceProvider;

            var reader = provider.GetRequiredService<IExcelReader>();

            var validators = new IRowValidator<ExcelInvoiceRow>[] { new InvoiceRowValidator() };

            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File does not exist: {filePath}");
                return;
            }

            var result = reader.ReadExcel<ExcelInvoiceRow>(filePath, rowValidators: validators);

            
            //Sample counter results for valid/invalid also over cells/rows 
            Console.WriteLine($"Total Rows: {result.TotalRows}");
            Console.WriteLine($"Valid Rows: {result.ValidRowsCount}");
            Console.WriteLine($"Invalid Rows: {result.InvalidRowsCount}");
            Console.WriteLine($"Cell Errors: {result.TotalCellErrorsCount}");
            Console.WriteLine($"Cell Warnings: {result.TotalCellWarningsCount}");
            Console.WriteLine($"Row Errors: {result.TotalRowErrorsCount}");

            result.GenerateErrorsExcelFile(errorsOut, filePath);
            Console.WriteLine($"Errors file saved: {Path.GetFullPath(errorsOut)}");

            await host.StopAsync();


            /////////////////////////Option two///////////////////////////////
            //or  majd you colud use this  simple instance creation 


            //string filePath = "DeskTop\SampleTransactions.xlsx";

            //if (!File.Exists(filePath))
            //{
            //    Console.WriteLine($"Error: File not found at path '{filePath}'");
            //    return;
            //}

            //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            //string inputPath = filePath;
            //string errorsOut = $"invoices_errors_{DateTime.Now:yyyy-MM-dd_HH-mm-ss}.xlsx";

            ////IExcelReader reader = new ExcelReader();
            //var validators = new List<IRowValidator<ExcelInvoiceRow>> { new SampleInvoiceRowValidator() };

            //var result = reader.ReadExcel<ExcelInvoiceRow>(
            //    filePath: inputPath,
            //    headerRow: 1,
            //    rowValidators: new[] { new SampleInvoiceRowValidator() }
            //);


            //Console.WriteLine($"Total Rows: {result.TotalRows}");
            //Console.WriteLine($"Valid Rows: {result.ValidRowsCount}");
            //Console.WriteLine($"Invalid Rows: {result.InvalidRowsCount}");
            //Console.WriteLine($"Cell Errors: {result.TotalCellErrorsCount}");
            //Console.WriteLine($"Row Errors: {result.TotalRowErrorsCount}");

            //Console.WriteLine($"Cell Warnings: {result.TotalCellWarningsCount}");
            //Console.WriteLine($"Row Warnings: {result.TotalRowWarningsCount}");



            //result.GenerateErrorsExcelFile(errorsOut);
            //Console.WriteLine($"Errors file: {Path.GetFullPath(errorsOut)}");

        }

    }

}
