public static class CommonsExtensions
{
    /// <summary>
    /// Creates on-the-fly EF dbset from a non-generic model type.
    /// </summary>
    /// <param name="context">The database context.</param>
    /// <param name="modelType">The model type.</param>
    /// <returns>
    /// A dynamic <c>Queryable</c> containing the data corresponding to <c>modelType</c>
    /// </returns>
    public static IQueryable<object> Set(this DbContext context, Type modelType)
    {
        return (IQueryable<object>)context.GetType().GetMethods()
            .Where(x => x.Name is "Set")
            .FirstOrDefault(x => x.IsGenericMethod)
            .MakeGenericMethod(modelType)
            .Invoke(context, null); // context.Set<modelType>()
    }

    /// <summary>
    /// Uploads a file to the webhost environment of the program.
    /// </summary>
    /// <param name="file">The file to upload.</param>
    /// <param name="_environment">The filehosting environment.</param>
    /// <returns>
    /// The filepath to the uploaded file.
    /// </returns>
    public static async Task<string> Upload(this IFormFile file, IWebHostEnvironment _environment)
    {
        var fileName = Path.GetFileName(file.FileName);
        var uploadPath = Path.Combine(_environment.WebRootPath, "uploads");
        var filePath = Path.Combine(uploadPath, fileName);
        using (var stream = new FileStream(filePath, FileMode.Create))
        {
            await file.CopyToAsync(stream);
        }
        return filePath;
    }
}