using System.IO;
using System.Threading.Tasks;

namespace APIExcelValidator.Abstractions
{
    public interface IFileValidator
    {
        public void Validate(Stream file);
    }
}