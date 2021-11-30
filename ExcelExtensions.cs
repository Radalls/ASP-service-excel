public static class ExcelExtensions
{
    /// <summary>
    /// Determines if a model property is a primary key.
    /// </summary>
    /// <param name="property">The property to verify.</param>
    /// <returns>
    /// Whether or not the property is a primary key.
    /// </returns>
    public static bool IsPrimaryKey(this PropertyInfo property)
    {
        return property.Name is "Id";
    }

    /// <summary>
    /// Determines if a model property is a foreign key.
    /// </summary>
    /// <param name="property">The property to verify.</param>
    /// <returns>
    /// Wheter or not the property is a foreign key.
    /// </returns>
    public static bool IsForeignKey(this PropertyInfo property)
    {
        return property.Name.Contains("Id") && property.Name is not "Id";
    }

    /// <summary>
    /// Determines if a model property stores a model instance corresponding to a foreign key of said model.
    /// </summary>
    /// <param name="f_property">The property of foreign type.</param>
    /// <param name="property">The foreign key property.</param>
    /// <returns>
    /// Whether or not the <c>PropertyType</c> of <c>f_property</c> corresponds to the foreign key.
    /// </returns>
    public static bool IsForeignModelFor(this PropertyInfo f_property, PropertyInfo property)
    {
        return property.Name.Contains(f_property.PropertyType.Name);
    }

    /// <summary>
    /// Determines if a model property can be converted to exploitable data in Excel.
    /// </summary>
    /// <param name="property">The property to verify.</param>
    /// <returns>
    /// Whether or not the property is compatible with Excel.
    /// </returns>
    public static bool IsExcelCompatible(this PropertyInfo property)
    {
        return property.PropertyType.IsPrimitive || property.PropertyType == typeof(string) || property.PropertyType == typeof(DateTime);
    }

    /// <summary>
    /// Determines if a model property will translate to a selectlist in Excel.
    /// </summary>
    /// <param name="property">The property to verify.</param>
    /// <returns>
    /// Whether or not the property will be converted to a selectlist.
    /// </returns>
    public static bool HasSelectOptionValues(this PropertyInfo property)
    {
        return Attribute.IsDefined(property, typeof(SelectOptionAttribute));
    }

    /// <summary>
    /// Determines if a file is an Excel file.
    /// </summary>
    /// <param name="file">The file to verify.</param>
    /// <returns>
    /// Whether or not the file is exploitable and also an Excel file.
    /// </returns>
    public static bool IsExcelFile(this IFormFile file)
    {
        return file is not null && file.Length is not 0 && Path.GetExtension(file.FileName).ToLower() is ".xlsx";
    }

    /// <summary>
    /// Gets the Excel exploitable name for a model property.
    /// </summary>
    /// <param name="property">The property to convert.</param>
    /// <returns>
    /// The Excel name for the property.
    /// </returns>
    public static string GetCellName(this PropertyInfo property)
    {
        var cellName = property.Name;
        if (Attribute.IsDefined(property, typeof(DisplayNameAttribute))) // attribute has a display name
        {
            cellName = (property.GetCustomAttribute(typeof(DisplayNameAttribute)) as DisplayNameAttribute).DisplayName;
        }
        if (Attribute.IsDefined(property, typeof(RequiredAttribute))) // attribute is required
        {
            cellName += " *";
        }
        return cellName;
    }

    /// <summary>
    /// Gets the Excel exploitable type name for a model property.
    /// </summary>
    /// <param name="property">The property to convert.</param>
    /// <returns>
    /// The simplified and translated type name for the property.
    /// </returns>
    public static string GetExcelTypeName(this PropertyInfo property)
    {
        return property.PropertyType.Name switch
        {
            "String" => "Texte",
            "Int32" => "Entier",
            "Boolean" => "O/N",
            "DateTime" => "Date (jj/mm/aaaa)",
            _ => property.PropertyType.Name,
        };
    }

    /// <summary>
    /// Creates a collection of Excel exploitable custom identifiers for all existing model instances of a certain type.
    /// </summary>
    /// <param name="context">The database context for this application.</param>
    /// <param name="modelType">The type of the model.</param>
    /// <returns>
    /// The list of all the model instances identifiers.
    /// </returns>
    /// <remarks>
    /// How this works :
    /// The foreign identifier for a model instance is composed of its database ID and also a more user-friendly ID determined beforehand in the model.
    /// </remarks>
    public static List<string> GetForeignIdentifiers(this DbContext context, Type modelType)
    {
        var identifiers = new List<string>();
        foreach (var instance in context.Set(modelType))
        {
            identifiers.Add(
                modelType.GetProperty("Id").GetValue(instance).ToString() 
                + " - " 
                + modelType.GetProperty(modelType.GetCustomAttribute<CustomIdentifierAttribute>().AttributeName).GetValue(instance).ToString()
            ); // id - sometext
        }
        return identifiers;
    }

    /// <summary>
    /// Gets a range in an Excel worksheet containing all the rows containing data to import.
    /// </summary>
    /// <param name="importWorksheet">The excel worksheet containing the data.</param>
    /// <param name="firstDataRowNumber">The number of the first row containing data in the worksheet.</param>
    /// <returns>
    /// The evaluated range.
    /// </returns>
    public static IXLRange GetDataRange(this IXLWorksheet importWorksheet, int firstDataRowNumber)
    {
        return importWorksheet.Range(
            importWorksheet.Row(firstDataRowNumber)
                            .FirstCell().Address,
            importWorksheet.LastRowUsed()
                            .Cell(importWorksheet.LastColumnUsed().ColumnNumber()).Address
        );
    }

    /// <summary>
    /// Gets the Excel cell data converted to an exploitable type in the application.
    /// </summary>
    /// <param name="cellData">The data to convert.</param>
    /// <param name="outputType">The type to which the data must be converted.</param>
    /// <returns>
    /// The converted cell data.
    /// </returns>
    public static object GetCellDataTo(this object cellData, Type outputType)
    {
        var cleanedCellValue = cellData.ToString().Trim().ToUpper();

        switch (outputType.Name)
        {
            case "Int32":
                if (cleanedCellValue is "")
                {
                    return 0;
                }
                else
                {
                    return int.Parse(cleanedCellValue);
                }
            case "Boolean":
                if (cleanedCellValue is "O")
                {
                    return true;
                }
                else
                {
                    return false;
                }
            case "DateTime":
                return DateTime.Parse(cleanedCellValue);
            default:
                if (cleanedCellValue is "")
                {
                    return null;
                }
                else
                {
                    return cellData.ToString();
                }
        }
    }

    /// <summary>
    /// Gets the Excel cell data converted to an exploitable ID in the application.
    /// </summary>
    /// <param name="cellData">The data to convert.</param>
    /// <returns>
    /// The database ID corresponding to the data.
    /// </returns>
    /// <remarks>
    /// How this works :
    /// The cell data is composed of a database ID and a custom ID, the method retrieves only the real database ID.
    /// </remarks>
    public static int GetForeignKeyData(this object cellData)
    {
        var foreignIdStr = cellData.ToString().Split()[0]; // id - sometext
        var foreignId = int.Parse(foreignIdStr);
        return foreignId;
    }

    /// <summary>
    /// Sets the indicative values for the name and type of a model property in a column of an Excel workbook.
    /// </summary>
    /// <param name="workbook">The Excel workbook to work on.</param>
    /// <param name="property">The property to convert.</param>
    /// <param name="nameRowNumber">The number of the row where the property name will be written.</param>
    /// <param name="typeRowNumber">The number of the row where the property type name will be written.</param>
    /// <param name="displayNameRowNumber">The number of the row where the property display name will be written.</param>
    /// <param name="currentColumnNumber">The number of the column where the property infos will be written.</param>
    public static void SetDataRowsNames(this XLWorkbook workbook, PropertyInfo property, int nameRowNumber, int typeRowNumber, int displayNameRowNumber, int currentColumnNumber)
    {
        var importWorksheet = workbook.Worksheet(1);

        importWorksheet.Cell(nameRowNumber, currentColumnNumber).Value = property.Name;
        importWorksheet.Cell(typeRowNumber, currentColumnNumber).Value = property.GetExcelTypeName();
        importWorksheet.Cell(displayNameRowNumber, currentColumnNumber).Value = property.GetCellName();
    }

    /// <summary>
    /// Sets the validation of an Excel workbook column to be a Selectlist. Also sets the values of that list.
    /// </summary>
    /// <param name="workbook">The Excel workbook to work on.</param>
    /// <param name="firstDataRowNumber">The number of the first row containing data in the import worksheet.</param>
    /// <param name="currentRowNumber">The number of the current row containing an option value in the options worksheet.</param>
    /// <param name="currentColumnNumber">The number of the current column holding the information of the property to set.</param>
    /// <param name="limitRowNumber">The number of the last row that will be affected by the modifications.</param>
    /// <param name="values">The option values of the Selectlist needed for validation.</param>
    public static void SetSelectValidation(this XLWorkbook workbook, int firstDataRowNumber, int currentRowNumber, int currentColumnNumber, int limitRowNumber, string[] values)
    {
        var importWorksheet = workbook.Worksheet(1);
        var optionsWorksheet = workbook.Worksheet(2);

        foreach (var value in values)
        {
            optionsWorksheet.Cell(currentRowNumber, currentColumnNumber).Value = value;
            currentRowNumber++;
        }

        var selectRange = importWorksheet.Range(
            importWorksheet.Column(currentColumnNumber)
                            .Cell(firstDataRowNumber).Address,
            importWorksheet.Column(currentColumnNumber)
                            .Cell(limitRowNumber).Address
        );
        var selectOptionsRange = optionsWorksheet.Range(
            optionsWorksheet.Column(currentColumnNumber)
                            .FirstCellUsed().Address,
            optionsWorksheet.Column(currentColumnNumber)
                            .LastCellUsed().Address
        );
        selectRange.SetDataValidation().List(selectOptionsRange, true);
    }

    /// <summary>
    /// Sets the style and type format for the content of an Excel workbook.
    /// </summary>
    /// <param name="workbook">The Excel workbook to work on.</param>
    /// <param name="nameRowNumber">The number of the row containing the names of the model properties.</param>
    /// <param name="typeRowNumber">The number of the row containing the type names of the model properties.</param>
    /// <param name="displayNameRowNumber">The number of the row containing the display names of the model properties.</param>
    /// <param name="firstDataRowNumber">The number of the first row containing data in the import worksheet.</param>
    /// <param name="limitRowNumber">The number of the last row that could contain data.</param>
    public static void SetStyleAndFormat(this XLWorkbook workbook, int nameRowNumber, int typeRowNumber, int displayNameRowNumber, int firstDataRowNumber, int limitRowNumber)
    {
        var importWorksheet = workbook.Worksheet(1);
        var optionsWorksheet = workbook.Worksheet(2);

        optionsWorksheet.Hide();
        importWorksheet.Row(nameRowNumber).Hide();
        importWorksheet.Row(typeRowNumber).CellsUsed().Style.Fill.SetBackgroundColor(XLColor.LightGray);
        importWorksheet.Row(displayNameRowNumber).CellsUsed().Style.Fill.SetBackgroundColor(XLColor.LightGray);

        var dataRange = importWorksheet.Range(
            importWorksheet.Row(firstDataRowNumber)
                            .FirstCell().Address,
            importWorksheet.Row(limitRowNumber)
                            .Cell(importWorksheet.LastColumnUsed().ColumnNumber()).Address
        );
        foreach (var cell in dataRange.Cells())
        {
            cell.Style.NumberFormat.Format = "@"; // change cell type format to text
        }

        foreach (var column in importWorksheet.ColumnsUsed())
        {
            column.AdjustToContents();
        }
    }

    /// <summary>
    /// Validates all the data in an Excel workbook. Also binds the valid data to a model instance.
    /// </summary>
    /// <typeparam name="TEntity">The model type.</typeparam>
    /// <param name="workbook">The Excel workbook to work on.</param>
    /// <param name="nameRowNumber">The number of the row containing the names of the model properties.</param>
    /// <param name="firstDataRowNumber">The number of the first row containing data in the import worksheet.</param>
    /// <param name="keepValidRows">Whether or not the application removes the valid rows of the potential error workbook.</param>
    /// <returns>
    /// The list of all the valid model instances to be added to the database.
    /// </returns>
    public static List<TEntity> ValidateImportData<TEntity>(this XLWorkbook workbook, int nameRowNumber, int firstDataRowNumber, bool keepValidRows) where TEntity : class
    {
        var importWorksheet = workbook.Worksheet(1);
        var dataRange = importWorksheet.GetDataRange(firstDataRowNumber);
        var dataIsValid = true;
        var instances = new List<TEntity>();

        foreach (var row in dataRange.Rows())
        {
            var modelType = typeof(TEntity);
            var modelProperties = modelType.GetProperties().ToList();
            var modelInstance = Activator.CreateInstance(modelType);

            var rowDataIsValid = workbook.ValidateRowData(nameRowNumber, row.RowNumber(), modelProperties, modelInstance);
            if (rowDataIsValid)
            {
                if (!keepValidRows)
                {
                    row.Delete();
                }

                instances.Add((TEntity)modelInstance);
            }
            else
            {
                dataIsValid = false;
            }
        }

        if (dataIsValid)
        {
            return instances;
        }
        else
        {
            return null;
        }
    }

    /// <summary>
    /// Validates the data of a row in an Excel workbook according to the data annotations of the model properties the row corresponds to.
    /// </summary>
    /// <param name="workbook">The Excel workbook to work on.</param>
    /// <param name="nameRowNumber">The number of the row containing the names of the model properties.</param>
    /// <param name="rowNumber">The number of the row to validate.</param>
    /// <param name="modelProperties">The model properties.</param>
    /// <param name="modelInstance">An instance of the model.</param>
    /// <returns>
    /// Whether or not all the data in the row is valid.
    /// </returns>
    public static bool ValidateRowData(this XLWorkbook workbook, int nameRowNumber, int rowNumber, List<PropertyInfo> modelProperties, object modelInstance)
    {
        var importWorksheet = workbook.Worksheet(1);
        var rowDataIsValid = true;
        foreach (var property in modelProperties)
        {
            if (!property.IsPrimaryKey() && property.IsExcelCompatible())
            {
                foreach (var column in importWorksheet.ColumnsUsed())
                {
                    var nameCell = importWorksheet.Cell(nameRowNumber, column.ColumnNumber());
                    if (nameCell.Value.ToString() == property.Name)
                    {
                        var cell = importWorksheet.Cell(rowNumber, column.ColumnNumber());
                        if (property.IsForeignKey())
                        {
                            var foreignKey = cell.Value.GetForeignKeyData();
                            property.SetValue(modelInstance, foreignKey);
                            break;
                        }
                        else
                        {
                            try
                            {
                                var cellData = cell.Value.GetCellDataTo(property.PropertyType);
                                var cellDataIsValid = cellData.ValidateCellData(property, modelInstance);
                                if (cellDataIsValid)
                                {
                                    property.SetValue(modelInstance, cellData);
                                    break;
                                }
                            }
                            catch (Exception)
                            {
                                cell.Style.Fill.BackgroundColor = XLColor.Red; // indicate error in excel file
                                rowDataIsValid = false;
                                break;
                            }
                        }
                    }
                }
            }
        }

        return rowDataIsValid;
    }

    /// <summary>
    /// Validates the data of a cell in an Excel workbook according to the validation context of the property the cell corresponds to.
    /// </summary>
    /// <param name="cellData">The cell data to validate.</param>
    /// <param name="property">The model property to validate.</param>
    /// <param name="modelInstance">An instance of the model.</param>
    /// <returns>
    /// Whether or not the data is valid.
    /// </returns>
    public static bool ValidateCellData(this object cellData, PropertyInfo property, object modelInstance)
    {
        var validationContext = new ValidationContext(modelInstance, null, null) { MemberName = property.Name };
        var validationResults = new List<ValidationResult>();
            
        return Validator.TryValidateProperty(cellData, validationContext, validationResults);
    }

    /// <summary>
    /// Converts a model and its properties to exploitable and exportable data in an Excel workbook.
    /// </summary>
    /// <param name="workbook">The Excel workbook to work on.</param>
    /// <param name="modelProperties">The properties of the model.</param>
    /// <param name="nameRowNumber">The number of the row containing the names of the model properties.</param>
    /// <param name="typeRowNumber">The number of the row containing the type names of the model properties.</param>
    /// <param name="displayNameRowNumber">The number of the row containing the display names of the model properties.</param>
    /// <param name="firstDataRowNumber">The number of the first row containing data in the import worksheet.</param>
    /// <param name="firstOptionRowNumber">The number of the first row containing data in the options worksheet.</param>
    /// <param name="limitRowNumber">The number of the last row that could contain data.</param>
    /// <param name="_context">The application database context.</param>
    public static void ToExcel(this XLWorkbook workbook, PropertyInfo[] modelProperties, int nameRowNumber, int typeRowNumber, int displayNameRowNumber, int firstDataRowNumber, int firstOptionRowNumber, int limitRowNumber, CompanyContext _context)
    {
        var currentColumnNumber = 1;
        foreach (var property in modelProperties)
        {
            if (!property.IsPrimaryKey() && property.IsExcelCompatible())
            {
                workbook.SetDataRowsNames(property, nameRowNumber, typeRowNumber, displayNameRowNumber, currentColumnNumber);

                if (property.IsForeignKey())
                {
                    foreach (var f_property in modelProperties)
                    {
                        if (f_property.IsForeignModelFor(property))
                        {
                            var foreignModelType = f_property.PropertyType;
                            var identifiers = _context.GetForeignIdentifiers(foreignModelType).ToArray();
                            workbook.SetSelectValidation(firstDataRowNumber, firstOptionRowNumber, currentColumnNumber, limitRowNumber, identifiers);
                        }
                    }
                }
                else if (property.HasSelectOptionValues())
                {
                    var options = (property.GetCustomAttribute(typeof(SelectOptionAttribute)) as SelectOptionAttribute).Values;
                    workbook.SetSelectValidation(firstDataRowNumber, firstOptionRowNumber, currentColumnNumber, limitRowNumber, options);
                }

                currentColumnNumber++;
            }
        }
    }

    /// <summary>
    /// Converts an Excel workbook into filecontent for download purposes.
    /// </summary>
    /// <param name="workbook">The workbook to convert.</param>
    /// <returns>
    /// The byte stream corresponding to the workbook content.
    /// </returns>
    public static byte[] ToFileContent(this XLWorkbook workbook)
    {
        var stream = new MemoryStream();
        workbook.SaveAs(stream);
        var content = stream.ToArray();
        return content;
    }
}
