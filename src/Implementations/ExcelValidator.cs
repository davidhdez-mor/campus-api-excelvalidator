using System;
using System.Data;
using System.IO;
using System.Text;
using APIExcelValidator.Abstractions;
using ExcelDataReader;
using ExcelDataReader.Exceptions;

namespace APIExcelValidator.Implementations
{
    public class ExcelValidator : IFileValidator, IFieldValidator
    {
        private readonly IExcelTable _excelTable;
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
                int toIndex = 37;

                var dataset = reader.AsDataSet(_excelDataSetConfig);
                var dataTable = dataset.Tables[0];
                // string match = dataTable.Rows[fromIndex][0].ToString();

                string match = "BANCO VE POR MAS, S.A. FIDEICO (UVCGLOBAL USD 630603 AMEX)";

                for (int i = fromIndex; i <= toIndex; i++)
                {
                    string data = dataTable.Rows[i][0].ToString();

                    if (data is not null && !data.Equals(match))
                    {
                        Console.Error.WriteLine($"Error en la columna A fila {i + 2}");
                        return false;
                    }
                }
            }

            return true;
        }

        public bool ValidateRangeField(DataTable table, int fromColIndex, int toColIndex, int fromRowIndex, int toRowIndex)
        {
            for (int col = fromColIndex; col <= toColIndex; col++)
            {
                for (int row = fromRowIndex; row <= toRowIndex; row++)
                {
                    try
                    {
                        float value = float.Parse(table.Rows[row][col].ToString());
                    }
                    catch (Exception ex)
                    {
                        Console.Error.WriteLine($"Error en la columna {col} fila {row + 2}");
                        return false;
                    }
                }
            }

            return true;
        }

        public bool ValidateNumbers(Stream file)
        {
            using (var reader = ExcelReaderFactory.CreateReader(file, _excelConfig))
            {
                int fromRowIndex = 4;
                int toRowIndex = 37;

                // K - P: 10 - 15
                // U - W: 20 - 22
                int fromColIndex = 10;
                int toColIndex = 15;

                var dataset = reader.AsDataSet(_excelDataSetConfig);
                var dataTable = dataset.Tables[0];

                bool firstNumberRange =
                    ValidateRangeField(dataTable, fromColIndex, toColIndex, fromRowIndex, toRowIndex);
                bool secondNumberRange = 
                    ValidateRangeField(dataTable, 20, 22, fromRowIndex, toRowIndex);
                return firstNumberRange && secondNumberRange;
            }
        }
    }
}