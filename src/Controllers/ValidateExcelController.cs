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
                return BadRequest("No se envio nada o el archivo esta vacio");
            
            if (!_fileValidator.Validate(file.OpenReadStream()))
                return BadRequest("El archivo no es una hoja de calculo");

            if (!((IFieldValidator) _fileValidator).ValidateDescription(file.OpenReadStream()))
                return BadRequest("Error en la columna x y la fila y descripcion invalida");
            
            if (!((IFieldValidator) _fileValidator).ValidateNumbers(file.OpenReadStream()))
                return BadRequest("Error en la columna x y la fila y: no es un numero");

            return Ok($"El archivo {file.Name} es valido");
        }
    }
}