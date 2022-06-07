using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using System.IO;
using System.Threading.Tasks;
using XlsExport.Api.Services;

namespace XlsExport.Api.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class ExportController : ControllerBase
    {
        private readonly ILogger<ExportController> _logger;

        public ExportController(ILogger<ExportController> logger)
        {
            _logger = logger;
        }

        [HttpGet]
        public async Task<IActionResult> Download()
        {
            ExportService exportService = new ExportService();
            var teste = exportService.GetApplicantsStatistics();

            using MemoryStream stream = new MemoryStream();
            teste.SaveAs(stream);
            return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "teste-arquivo.xlsx");
        }
    }
}