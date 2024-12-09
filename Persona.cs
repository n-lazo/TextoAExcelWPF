namespace TextoAExcelWPF;

/// <summary>
/// Clase para representar una persona
/// y generar los objetos necesarios
/// y listas de objetos para la exportación
/// Tenemos los nullables con ? desde C# 8.0
/// </summary>
public class Persona
{
    /// <summary>
    /// Propiedad para el nombre de la persona
    /// </summary>
    public string? Nombre { get; set; }
    /// <summary>
    /// Propiedad para el apellido de la persona
    /// </summary>
    public string? Apellido { get; set; }
    /// <summary>
    /// Propiedad para el nombre completo de la persona
    /// </summary>
    public string? NombreCompleto { get; set; }
}