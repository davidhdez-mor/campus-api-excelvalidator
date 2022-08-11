using System;
using System.Data;
using System.IO;
using System.Text;
using APIExcelValidator.Abstractions;
using APIExcelValidator.Exceptions;
using ExcelDataReader;
using ExcelDataReader.Exceptions;

namespace APIExcelValidator.Implementations
{
    public class ExcelValidator : IFileValidator, IFieldValidator
    {
        private readonly IExcelTable _excelTable;
        private readonly ExcelReaderConfiguration _excelConfig;
        private readonly ExcelDataSetConfiguration _excelDataSetConfig;
        public string ErrorMessage { get; set; }

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
                ErrorMessage = "El archivo no es una hoja de calculo";
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

                try
                {
                    return ValidateRangeField.ValidateDescriptionFields(dataTable, match, 0, 1, fromIndex, toIndex);
                }
                catch (InvalidDescriptionTypeException ex)
                {
                    ErrorMessage = ex.Message;
                    return false;
                } 
            }
        }

        public bool ValidateNumbers(Stream file)
        {
            using (var reader = ExcelReaderFactory.CreateReader(file, _excelConfig))
            {
                // K - P: 10 - 15
                // U - W: 20 - 22
                
                int fromRowIndex = 4;
                int toRowIndex = 37;
                int fromColIndex = 10;
                int toColIndex = 15;

                var dataset = reader.AsDataSet(_excelDataSetConfig);
                var dataTable = dataset.Tables[0];

                try
                {
                    bool firstNumberRange =
                        ValidateRangeField.ValidateNumberFromFields(dataTable, fromColIndex, toColIndex, fromRowIndex,
                            toRowIndex);
                    bool secondNumberRange =
                        ValidateRangeField.ValidateNumberFromFields(dataTable, 20, 22, fromRowIndex, toRowIndex);

                    return firstNumberRange && secondNumberRange;
                }
                catch (InvalidNumberTypeException ex)
                {
                    ErrorMessage = ex.Message;
                    return false;
                }
            }
        }
    }
}