using Microsoft.AspNetCore.Mvc;
using Lab13_DiegoMedina.Services;

namespace Lab13_DiegoMedina.Controllers;

[ApiController]
[Route("api/[controller]")]
public class ReportesController : ControllerBase
{
    private readonly ReportesService _reportesService;

    public ReportesController(ReportesService reportesService)
    {
        _reportesService = reportesService;
    }

    [HttpGet("clientes")]
    public IActionResult Clientes()
    {
        var bytes = _reportesService.GenerarReporteClientes();
        return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "reporte-clientes.xlsx");
    }

    [HttpGet("pedidos")]
    public IActionResult Pedidos()
    {
        var bytes = _reportesService.GenerarReportePedidos();
        return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "reporte-pedidos.xlsx");
    }

    [HttpGet("modificar")]
    public IActionResult Modificar()
    {
        _reportesService.FirstExample();
        _reportesService.SecondExample();
        var bytes = System.IO.File.ReadAllBytes("reporteCreado.xlsx");
        return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "reporte-modificado.xlsx");
    }
}
