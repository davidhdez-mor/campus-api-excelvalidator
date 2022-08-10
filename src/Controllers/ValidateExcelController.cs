using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using APIExcelValidator.Implementations;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace APIExcelValidator.Controllers
{
    [Route("api/ExcelValidator")]
    [ApiController]
    public class ValidateExcelController : ControllerBase
    {
        // IFormFile
        // Endpoints async
        
        // POST file api/ExcelValidator
        [HttpPost]
        public async Task<IActionResult> OnLoadFile()
        {
            IFormFile file = HttpContext.Request.Form.Files.FirstOrDefault();
            
            if (file is null)
                return BadRequest("Ningun archivo proporcioando");

            ExcelValidator excelValidator = new ExcelValidator();
            if (!excelValidator.Validate(file.OpenReadStream()))
                return BadRequest("El archivo no es una hoja de calculo");

            if (!excelValidator.ValidateDescription(file.OpenReadStream()))
                return BadRequest("Error en la columna x y la fila y");
            
            return Ok(file.Length);
        }
    }
}