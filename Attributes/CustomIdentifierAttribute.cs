[AttributeUsage(AttributeTargets.Class, AllowMultiple = false)]
public class CustomIdentifierAttribute : Attribute
{
    /// <summary>
    /// Gets or sets the attribute name that will be used to create the identifier.
    /// </summary>
    /// <value>
    /// The name of the attribute.
    /// </value>
    public string AttributeName { get; set; }
}
