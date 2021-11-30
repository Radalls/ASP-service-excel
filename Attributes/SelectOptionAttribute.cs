[AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
public class SelectOptionAttribute : Attribute
{
    /// <summary>
    /// Gets or sets the options values that will be used in the select view
    /// </summary>
    /// <value>
    /// The select options values.
    /// </value>
    public string[] Values { get; set; }
}
