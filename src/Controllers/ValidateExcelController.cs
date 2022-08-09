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
            excelValidator.Validate(file.OpenReadStream());
            return Ok(file.Length);
        }
    }
}