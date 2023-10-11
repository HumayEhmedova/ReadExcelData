using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using ReadExcelFiles.DTOs;
using static OfficeOpenXml.ExcelErrorValue;
using System.Xml.Linq;
using ReadExcelFiles.Models;

namespace ReadExcelFiles.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelController : ControllerBase
    {
        [HttpGet]
        public IActionResult Get()
        {
            var result = Excel.ReturnExcelDatas();
            return Ok(result.Take(1));
        }

        [HttpGet("GetByFilter")]
        public IActionResult GetByFilter( )
        {
           ExcelDataManipulationDto dto = new ExcelDataManipulationDto();
            var result = Excel.ReturExcelDataByFilter( dto);
            return Ok(result.Take(1));
        }
    }
}

