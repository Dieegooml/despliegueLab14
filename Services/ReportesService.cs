using ClosedXML.Excel;
using Lab13_DiegoMedina.Data;
using Microsoft.EntityFrameworkCore;

namespace Lab13_DiegoMedina.Services;

public class ReportesService
{
    private readonly AppDbContext _context;

    public ReportesService(AppDbContext context)
    {
        _context = context;
    }

    public void FirstExample()
    {
        using (var workbook = new XLWorkbook())
        {
            var worksheet = workbook.Worksheets.Add("Hoja Excell creada");

            worksheet.Cell(1, 1).Value = "Nombre";
            worksheet.Cell(1, 2).Value = "Edad";
            worksheet.Cell(2, 1).Value = "Diego";
            worksheet.Cell(2, 2).Value = 21;

            workbook.SaveAs("reporteCreado.xlsx");
        }
    }

    public void SecondExample()
    {
        using (var workbook = new XLWorkbook("reporteCreado.xlsx"))
        {
            var worksheet = workbook.Worksheet(1);

            worksheet.Cell(2, 2).Value = 30;

            workbook.Save();
        }
    }

    public byte[] GenerarReporteClientes()
    {
        var clientes = _context.Clients.OrderBy(c => c.Clientid).ToList();

        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add("Clientes");

        var headers = new[] { "ID", "Nombre", "Email" };
        for (var i = 0; i < headers.Length; i++)
        {
            worksheet.Cell(1, i + 1).Value = headers[i];
        }

        var headerRange = worksheet.Range(1, 1, 1, headers.Length);
        headerRange.Style.Font.Bold = true;
        headerRange.Style.Fill.BackgroundColor = XLColor.Cyan;
        headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

        var row = 2;
        foreach (var cliente in clientes)
        {
            worksheet.Cell(row, 1).Value = cliente.Clientid;
            worksheet.Cell(row, 2).Value = cliente.Name;
            worksheet.Cell(row, 3).Value = cliente.Email;
            row++;
        }

        if (clientes.Count > 0)
        {
            worksheet.Range(1, 1, clientes.Count + 1, headers.Length).CreateTable();
        }

        worksheet.Columns().AdjustToContents();
        worksheet.SheetView.FreezeRows(1);

        using var stream = new MemoryStream();
        workbook.SaveAs(stream);
        return stream.ToArray();
    }

    public byte[] GenerarReportePedidos()
    {
        var pedidos = _context.Orders
            .Include(o => o.Client)
            .Include(o => o.Orderdetails)
            .ThenInclude(d => d.Product)
            .OrderBy(o => o.Orderid)
            .ToList();

        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add("Pedidos");

        var headers = new[] { "ID Pedido", "Cliente", "Fecha", "Producto", "Cantidad", "Precio", "Subtotal" };
        for (var i = 0; i < headers.Length; i++)
        {
            worksheet.Cell(1, i + 1).Value = headers[i];
        }

        var headerRange = worksheet.Range(1, 1, 1, headers.Length);
        headerRange.Style.Font.Bold = true;
        headerRange.Style.Fill.BackgroundColor = XLColor.Cyan;
        headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

        var row = 2;
        foreach (var pedido in pedidos)
        {
            foreach (var detalle in pedido.Orderdetails)
            {
                worksheet.Cell(row, 1).Value = pedido.Orderid;
                worksheet.Cell(row, 2).Value = pedido.Client.Name;
                worksheet.Cell(row, 3).Value = pedido.Orderdate;
                worksheet.Cell(row, 4).Value = detalle.Product.Name;
                worksheet.Cell(row, 5).Value = detalle.Quantity;
                worksheet.Cell(row, 6).Value = detalle.Product.Price;
                worksheet.Cell(row, 7).Value = detalle.Quantity * detalle.Product.Price;
                row++;
            }
        }

        var lastRow = row - 1;
        if (lastRow >= 2)
        {
            worksheet.Range(1, 1, lastRow, headers.Length).CreateTable();
        }

        worksheet.Columns().AdjustToContents();
        worksheet.SheetView.FreezeRows(1);

        using var stream = new MemoryStream();
        workbook.SaveAs(stream);
        return stream.ToArray();
    }
}
