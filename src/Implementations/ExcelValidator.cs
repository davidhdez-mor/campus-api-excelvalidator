using System;
using System.IO;
using System.Text;
using APIExcelValidator.Abstractions;
using ExcelDataReader;

namespace APIExcelValidator.Implementations
{
    public class ExcelValidator : IFileValidator
    {
        public void Validate(Stream file)
        {
            
            // Sets the encoding to newer version, default 1252
            var readerConfiguration = new ExcelReaderConfiguration()
            {
                FallbackEncoding = Encoding.GetEncoding(1252)
            };
            
            using (var excelFile = ExcelReaderFactory.CreateReader(file, readerConfiguration))
            {
                Console.WriteLine(excelFile.Name);
                Console.WriteLine(excelFile.ResultsCount);
            }
        }
    }
}