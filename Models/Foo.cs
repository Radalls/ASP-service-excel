[CustomIdentifier(AttributeName = "FirstName")]
public class Foo
{
    [Key]
    public int Id { get; set; }

    [Required, MaxLength(30)]
    public string Firstname { get; set; }

    [Required, MaxLength(30)]
    public string LastName { get; set; }

    [Required, MaxLength(5), MinLength(5), RegularExpression("^[0-9]{5}$")]
    public string PostalCode { get; set; }

    [Required, SelectOption(Values = new string[] { "Toto", "Titi", "Tutu" })]
    public string FriendName { get; set; }

    [Required]
    public DateTime BirthDate { get; set; }

    [Required]
    public bool isMinor { get; set; }

    [RequiredIf("isMinor", true)]
    public int Age { get; set; }
}
