namespace Lookout.Options;

/// <summary>
/// Configures attachment export.
/// </summary>
internal class ExportAttachmentsOptions
{
    /// <summary>
    /// The default maximum number of folder recursion levels.
    /// </summary>
    public const int DefaultMaxRecurseLevels = 10;

    /// <summary>
    /// Gets or sets the folder where attachments are exported.
    /// </summary>
    public DirectoryInfo TargetFolder { get; set; } = new DirectoryInfo(Directory.GetCurrentDirectory());

    /// <summary>
    /// Gets or sets the maximum number of folder recursion levels.
    /// </summary>
    public int MaxRecurseLevel { get; set; } = DefaultMaxRecurseLevels;

    /// <summary>
    /// Gets or sets a value indicating whether the export is a dry run.
    /// </summary>
    public bool WhatIf { get; set; }
}
