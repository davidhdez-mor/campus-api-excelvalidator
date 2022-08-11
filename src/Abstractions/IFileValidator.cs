using System.IO;
using System.Threading.Tasks;

namespace APIExcelValidator.Abstractions
{
    public interface IFileValidator
    {
        public string ErrorMessage { get; set; }
        public bool Validate(Stream file);
    }
}