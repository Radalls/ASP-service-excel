﻿public class FooController : Controller
{
    private readonly IWebHostEnvironment _environment;
    private readonly DbContext _context;

    public FooController(IWebHostEnvironment environment, DbContext context)
    {
        _environment = environment;
        _context = context;
    }

    public IActionResult ExportExcel()
    {
        var excel = new ExcelService(_environment, _context);
        var content = excel.Export<Foo>();
        return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Foo.xlsx");
    }

    [HttpPost]
    public async Task<IActionResult> ImportExcel(IFormFile file)
    {
        var excel = new ExcelService(_environment, _context);
        var errorContent = await excel.Import<Foo>(file);
        if (errorContent is null)
        {
            return Redirect(Request.Headers["Referer"].ToString()); // back to index page
        }
        else
        {
            return File(errorContent, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Foo.xlsx"); // export error file
        }
    }
}
