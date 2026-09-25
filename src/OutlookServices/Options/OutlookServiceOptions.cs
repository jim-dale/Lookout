namespace OutlookServices.Options;

/// <summary>
/// Configures how <see cref="OutlookServices.OutlookService"/> discovers and loads Outlook stores.
/// </summary>
public class OutlookServiceOptions
{
    /// <summary>
    /// Gets or sets a provider-specific descriptor that identifies the underlying storage for the entity.
    /// </summary>
    /// <remarks>
    /// May be the full or partial store name, id, or path. Can also be 'default' to match the default store.
    /// If not specified or '*' will match all stores.
    /// </remarks>
    public string? StoreDescriptor { get; set; }

    /// <summary>
    /// Gets or sets the additional Outlook data files to load.
    /// </summary>
    public List<FileInfo> AdditionalDataFiles { get; set; } = [];
}
