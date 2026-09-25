namespace Lookout.Options;

/// <summary>
/// Configures contact listing.
/// </summary>
internal class ListContactsOptions
{
    /// <summary>
    /// The default maximum number of folder recursion levels.
    /// </summary>
    public const int DefaultMaxRecurseLevels = 10;

    /// <summary>
    /// Gets or sets the names of the contact properties to display.
    /// </summary>
    public string[] PropertyNames { get; set; } = [];

    /// <summary>
    /// Gets or sets the maximum number of folder recursion levels.
    /// </summary>
    public int MaxRecurseLevel { get; set; } = DefaultMaxRecurseLevels;
}
