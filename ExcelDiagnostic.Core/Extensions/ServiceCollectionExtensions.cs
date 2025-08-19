using ExcelDiagnostic.Core.Contracts;
using ExcelDiagnostic.Core.Readers;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection.Extensions;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDiagnostic.Core.Extensions
{
    public static class ServiceCollectionExtensions
    {
        public static IServiceCollection AddExcelDiagnostics(this IServiceCollection services)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;// Set the license context for EPPlus

            //You could only take this DI 
            services.TryAddScoped<IExcelReader, ExcelReader>();
            return services;
        }
    }
}
