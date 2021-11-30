[AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
public sealed class RequiredIfAttribute : ValidationAttribute
{
    /// <summary>
    /// Gets or sets the other property name that will be used during validation.
    /// </summary>
    /// <value>
    /// The other property name.
    /// </value>
    public string OtherPropertyName { get; private set; }

    /// <summary>
    /// Gets or sets the display name of the other property.
    /// </summary>
    /// <value>
    /// The display name of the other property.
    /// </value>
    public string OtherPropertyDisplayName { get; set; }

    /// <summary>
    /// Gets or sets the other property value that will be relevant for validation.
    /// </summary>
    /// <value>
    /// The other property value.
    /// </value>
    public object OtherPropertyValue { get; private set; }

    /// <summary>
    /// Gets or sets a value indicating whether other property's value should match or differ from provided other property's value (default is <c>false</c>).
    /// </summary>
    /// <value>
    ///   <c>true</c> if other property's value validation should be inverted; otherwise, <c>false</c>.
    /// </value>
    /// <remarks>
    /// How this works
    /// - true: validated property is required when other property doesn't equal provided value
    /// - false: validated property is required when other property matches provided value
    /// </remarks>
    public bool ValuesMustMatch { get; set; }

    /// <summary>
    /// Gets a value that indicates whether the attribute requires validation context.
    /// </summary>
    /// <returns><c>true</c> if the attribute requires validation context; otherwise, <c>false</c>.</returns>
    public override bool RequiresValidationContext
    {
        get { return true; }
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="RequiredIfAttribute"/> class.
    /// </summary>
    /// <param name="otherProperty">The other property.</param>
    /// <param name="otherPropertyValue">The other property value.</param>
    public RequiredIfAttribute(string otherProperty, object otherPropertyValue)
        : base("'{0}' is required because '{1}' has a value {3}'{2}'.")
    {
        OtherPropertyName = otherProperty;
        OtherPropertyValue = otherPropertyValue;
        ValuesMustMatch = false;
    }

    /// <summary>
    /// Applies formatting to an error message, based on the data field where the error occurred.
    /// </summary>
    /// <param name="name">The name to include in the formatted message.</param>
    /// <returns>
    /// An instance of the formatted error message.
    /// </returns>
    public override string FormatErrorMessage(string name)
    {
        return string.Format(
            CultureInfo.CurrentCulture,
            ErrorMessageString,
            name,
            OtherPropertyDisplayName ?? OtherPropertyName,
            OtherPropertyValue,
            ValuesMustMatch ? "other than " : "of ");
    }

    /// <summary>
    /// Validates the specified value with respect to the current validation attribute.
    /// </summary>
    /// <param name="value">The value to validate.</param>
    /// <param name="validationContext">The context information about the validation operation.</param>
    /// <returns>
    /// An instance of the <see cref="T:System.ComponentModel.DataAnnotations.ValidationResult" /> class.
    /// </returns>
    protected override ValidationResult IsValid(object value, ValidationContext validationContext)
    {
        if (validationContext is null)
        {
            throw new ArgumentNullException(nameof(validationContext));
        }

        PropertyInfo otherProperty = validationContext.ObjectType.GetProperty(OtherPropertyName);
        if (otherProperty is null)
        {
            return new ValidationResult(
                string.Format(CultureInfo.CurrentCulture, "Could not find a property named '{0}'.", OtherPropertyName));
        }

        object otherValue = otherProperty.GetValue(validationContext.ObjectInstance);

        if (!ValuesMustMatch && Equals(otherValue, OtherPropertyValue) ||
            ValuesMustMatch && !Equals(otherValue, OtherPropertyValue))
        {
            if (value is null)
            {
                return new ValidationResult(FormatErrorMessage(validationContext.DisplayName));
            }
                
            if (value is string str && str.Trim().Length is 0)
            {
                return new ValidationResult(FormatErrorMessage(validationContext.DisplayName));
            }
                
            if (value is int i && i is 0)
            {
                return new ValidationResult(FormatErrorMessage(validationContext.DisplayName));
            }
        }

        return ValidationResult.Success;
    }
}
