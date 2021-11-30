public class ExcelService
{
    private readonly IWebHostEnvironment _environment;
    private readonly CompanyContext _context;

    private static readonly int nameRowNumber = 1;
    private static readonly int typeRowNumber = 2;
    private static readonly int displayNameRowNumber = 3;
    private static readonly int firstDataRowNumber = 4;
    private static readonly int firstOptionRowNumber = 1;
    private static readonly int limitRowNumber = 500;

    public ExcelService(IWebHostEnvironment environment, CompanyContext context)
    {
        _environment = environment;
        _context = context;
    }

    /// <summary>
    /// Creates an Excel workbook and converts a model and its properties into exploitable data in said workbook.
    /// </summary>
    /// <typeparam name="TEntity">The type of the model to export.</typeparam>
    /// <returns>
    /// The file content corresponding to the workbook content.
    /// </returns>
    /// <remarks>
    /// How it works :
    /// Each property of the model is converted into a column in the worksheet. 
    /// The first rows contains informations helping the user to fill out the data about the model.
    /// Each empty row after that correspond to a potential instance of the model.
    /// </remarks>
    public byte[] Export<TEntity>() where TEntity : class
    {
        var modelType = typeof(TEntity);
        var modelProperties = modelType.GetProperties();

        var workbook = new XLWorkbook();
        workbook.Worksheets.Add("Import - " + modelType.Name);
        workbook.Worksheets.Add("Options - " + modelType.Name);

        workbook.ToExcel(modelProperties, nameRowNumber, typeRowNumber, displayNameRowNumber, firstDataRowNumber, firstOptionRowNumber, limitRowNumber, _context);
        workbook.SetStyleAndFormat(nameRowNumber, typeRowNumber, displayNameRowNumber, firstDataRowNumber, limitRowNumber);

        return workbook.ToFileContent();
    }

    /// <summary>
    /// Validates and converts Excel data into model instances to be added to the application database.
    /// </summary>
    /// <typeparam name="TEntity">The type of the model.</typeparam>
    /// <param name="file">The Excel file to import.</param>
    /// <returns>
    /// The workbook if it contains invalid data.
    /// </returns>
    /// <remarks>
    /// How it works :
    /// The valid data rows are converted to model instances, if there is any invalid data, the workbook is returned to the user with errors shown in red.
    /// </remarks>
    public async Task<byte[]> Import<TEntity>(IFormFile file) where TEntity : class
    {
        if (!file.IsExcelFile())
        {
            return null;
        }

        var filePath = file.Upload(_environment).Result;
        var workbook = new XLWorkbook(filePath);

        var instances = workbook.ValidateImportData<TEntity>(nameRowNumber, firstDataRowNumber, true);
        if (instances is not null)
        {
            _context.AddRange(instances);
            await _context.SaveChangesAsync();
            return null; // success
        }
        else
        {
            return workbook.ToFileContent(); // error
        }
    }
}
