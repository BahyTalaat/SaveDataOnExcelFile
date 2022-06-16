using ExcelSheet.Services;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace ExcelSheet.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelController : ControllerBase
    {
        WriteOnExcel writeOnExcel;
        public ExcelController()
        {
            writeOnExcel=new WriteOnExcel();
        }

        [HttpGet]
        [AllowAnonymous]
        public ActionResult Index()
        {
            writeOnExcel.write();
            return Ok();
        }
    }
}
