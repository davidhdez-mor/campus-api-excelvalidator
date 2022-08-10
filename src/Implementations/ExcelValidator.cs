using System;
using System.IO;
using System.Text;
using APIExcelValidator.Abstractions;
using ExcelDataReader;
using ExcelDataReader.Exceptions;
using Microsoft.AspNetCore.DataProtection.Repositories;

namespace APIExcelValidator.Implementations
{
    public class ExcelValidator : IFileValidator, IFieldValidator 
    {
        private readonly ExcelReaderConfiguration _excelConfig;
        private readonly ExcelDataSetConfiguration _excelDataSetConfig;

        public ExcelValidator()
        {
            // Sets the encoding to newer version, default 1252
            _excelConfig = new ExcelReaderConfiguration()
            {
                FallbackEncoding = Encoding.GetEncoding(1252)
            };

            _excelDataSetConfig = new ExcelDataSetConfiguration()
            {
                UseColumnDataType = true,
                ConfigureDataTable = _ => new ExcelDataTableConfiguration()
                {
                    UseHeaderRow = true
                }
            };
        }

        public bool Validate(Stream file)
        {
            try
            {
                using (var excelFile = ExcelReaderFactory.CreateReader(file, _excelConfig))
                {
                    // Console.WriteLine(excelFile.Name);
                    // Console.WriteLine(excelFile.ResultsCount);
                    return true;
                }
            }
            catch (HeaderException ex)
            {
                // Console.WriteLine($"El archivo no es una hoja de calculo: {ex.Message}");
                return false;
            }
        }

        public bool ValidateDescription(Stream file)
        {
            using (var reader = ExcelReaderFactory.CreateReader(file, _excelConfig))
            {
                int fromIndex = 4;
                int toIndex = 38;
                
                var dataset = reader.AsDataSet(_excelDataSetConfig);
                string match = dataset.Tables[0].Rows[fromIndex][0].ToString();

                for (int i = fromIndex + 1; i <= toIndex; i++)
                {
                    Console.WriteLine($"Fila {i}: {dataset.Tables[0].Rows[i][0]}");
                    string data = dataset.Tables[0].Rows[i][0].ToString();
                    if (data is not null && !data.Equals(match))
                    {
                        Console.Error.WriteLine($"Error en la columna: A fila: {i + 1}");
                        return false;
                    }
                }
            }

            return true;
        }

        public bool ValidateNumbers(Stream file)
        {
            return true;
        }
    }
}