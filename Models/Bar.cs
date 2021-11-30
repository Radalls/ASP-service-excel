public class Bar
{
    [Key]
    public int Id { get; set; }

    [Required, MaxLength(50), MinLength(1)]
    public string Name { get; set; }

    [DisplayName("Foo")]
    public int FooId { get; set; }

    [NotMapped]
    public Foo Foo { get; set; }
}
