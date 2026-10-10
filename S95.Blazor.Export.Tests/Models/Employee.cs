using System.ComponentModel.DataAnnotations;
using System.Runtime.Serialization;

namespace S95.Blazor.Export.Tests.Models;

public class Employee
{
    [Display(Name = "Full name")]
    public string FullName { get; set; } = string.Empty;

    [Display(Name = "Years")]
    public int Age { get; set; }

    public string? City { get; set; }

    [IgnoreDataMember]
    public string Password { get; set; } = string.Empty;

    public List<string> Skills { get; set; } = [];
}