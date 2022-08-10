using System.IO;

namespace APIExcelValidator.Abstractions
{
    public interface IFieldValidator
    {
        public bool ValidateDescription(Stream file);
        public bool ValidateNumbers(Stream file);
    }
}