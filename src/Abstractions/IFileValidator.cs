using System.IO;
using System.Threading.Tasks;

namespace APIExcelValidator.Abstractions
{
    public interface IFileValidator
    {
        public bool Validate(Stream file);
    }
}