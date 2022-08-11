using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using APIExcelValidator.Abstractions;
using APIExcelValidator.Implementations;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace APIExcelValidator.Controllers
{
    [Route("api/ExcelValidator")]
    [ApiController]
    public class ValidateExcelController : ControllerBase
    {
        private readonly IFileValidator _fileValidator;

        public ValidateExcelController(IFileValidator fileValidator)
        {
            _fileValidator = fileValidator;
        }

        // POST file api/ExcelValidator
        [HttpPost]
        public async Task<IActionResult> OnLoadFile(IFormFile file)
        {
            // IFormFile file = HttpContext.Request.Form.Files.FirstOrDefault();

            if (file is null)
                return BadRequest("El archivo esta vacio");

            if (!_fileValidator.Validate(file.OpenReadStream()))
                return BadRequest(((ExcelValidator)_fileValidator).ErrorMessage);

            if (!((IFieldValidator)_fileValidator).ValidateDescription(file.OpenReadStream()))
                return BadRequest(((ExcelValidator)_fileValidator).ErrorMessage);

            if (!((IFieldValidator)_fileValidator).ValidateNumbers(file.OpenReadStream()))
                return BadRequest(((ExcelValidator)_fileValidator).ErrorMessage);

            return Ok("El archivo es valido");
        }
    }
}