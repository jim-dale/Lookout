namespace OutlookServices;

using System.Diagnostics.CodeAnalysis;
using Microsoft.Extensions.Options;
using Microsoft.Office.Interop.Outlook;
using OutlookServices.Options;

/// <summary>
/// Provides access to the Outlook application, session, and configured data stores.
/// </summary>
public sealed class OutlookService : IDisposable
{
    private const string MatchAllStoresDescriptor = "*";
    private const string DefaultStoreDescriptor = "DEFAULT";
    private readonly OutlookServiceOptions options;
    private readonly List<DataFileStore> additionalStores = new();
    private Application? application;
    private NameSpace? session;

    /// <summary>
    /// Initializes a new instance of the <see cref="OutlookService"/> class.
    /// </summary>
    /// <param name="options">The configured Outlook service options.</param>
    public OutlookService(IOptions<OutlookServiceOptions> options)
    {
        ArgumentNullException.ThrowIfNull(options);
        ArgumentNullException.ThrowIfNull(options.Value);

        this.options = options.Value;
    }

    /// <summary>
    /// Initializes the Outlook application and session when they have not already been initialized.
    /// </summary>
    [MemberNotNull(nameof(application))]
    [MemberNotNull(nameof(session))]
    public void EnsureInitialized()
    {
        if (this.application is null || this.session is null)
        {
            this.application ??= new Application();
            this.session ??= this.application.GetNamespace(Constants.MapiNamespace);

            this.LoadAdditionalDataFileStore();
        }
    }

    /// <summary>
    /// Loads all additional data file stores configured in the service options.
    /// </summary>
    public void LoadAdditionalDataFileStore()
    {
        this.options.AdditionalDataFiles.ForEach((fi) => this.LoadDataFileStore(fi));
    }

    /// <summary>
    /// Loads an additional data file store from the specified file.
    /// </summary>
    /// <param name="path">The file containing the data store to load.</param>
    public void LoadDataFileStore(FileInfo path)
    {
        this.LoadDataFileStore(path.FullName);
    }

    /// <summary>
    /// Loads an additional data file store from the specified path.
    /// </summary>
    /// <param name="path">The path of the data store to load.</param>
    public void LoadDataFileStore(string path)
    {
        ArgumentNullException.ThrowIfNull(path);

        this.EnsureInitialized();

        DataFileStore item = new(this.session, path);

        if (item.TryAddStore())
        {
            this.additionalStores.Add(item);
        }
    }

    /// <summary>
    /// Gets all stores available in the Outlook session.
    /// </summary>
    /// <returns>All stores available in the Outlook session.</returns>
    public IEnumerable<Store> GetAllStores()
    {
        this.EnsureInitialized();

        return this.session.Stores.OfType<Store>();
    }

    /// <summary>
    /// Gets stores available according to the configured options.
    /// </summary>
    /// <returns>The stores available according to the configured options.</returns>
    public IEnumerable<Store> GetStores()
    {
        this.EnsureInitialized();

        string descriptor = string.IsNullOrWhiteSpace(this.options.StoreDescriptor) ? MatchAllStoresDescriptor : this.options.StoreDescriptor;

        if (descriptor == MatchAllStoresDescriptor)
        {
            foreach (Store item in this.GetAllStores())
            {
                yield return item;
            }
        }
        else if (string.Equals(descriptor, DefaultStoreDescriptor, StringComparison.InvariantCultureIgnoreCase))
        {
            yield return this.session.DefaultStore;
        }
        else
        {
            foreach (Store item in this.GetAllStores())
            {
                if (item.IsDescriptorMatch(descriptor))
                {
                    yield return item;
                }
            }
        }
    }

    /// <summary>
    /// Attempts to open a shared Outlook item at the specified <paramref name="path"/> using the
    /// provided Outlook session. This method is best-effort: failures are swallowed and reported
    /// via the returned value rather than propagated as exceptions.
    /// </summary>
    /// <param name="path">The filesystem or Outlook path of the shared item to open. Cannot be null or whitespace.</param>
    /// <exception cref="ArgumentException">Thrown when <paramref name="path"/> is null or whitespace.</exception>
    /// <returns>The <see cref="SharedItem"/> instance.</returns>
    public SharedItem CreateSharedItemFrom(string path)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);

        this.EnsureInitialized();

        SharedItem result = new();
        result.OpenSharedItem(path, this.session);

        return result;
    }

    /// <summary>
    /// Attempts to get the root folder for the specified store.
    /// </summary>
    /// <param name="store">The store whose root folder to get.</param>
    /// <param name="result">When this method returns, the store's root folder, if available.</param>
    /// <returns><see langword="true"/> if a root folder was found; otherwise, <see langword="false"/>.</returns>
    public bool TryGetRootFolder(Store store, [NotNullWhen(true)] out MAPIFolder result)
    {
        ArgumentNullException.ThrowIfNull(store);

        result = store.GetRootFolder();

        return result != default;
    }

    /// <summary>
    /// Releases the Outlook application, session, and additional data file stores.
    /// </summary>
    public void Dispose()
    {
        foreach (DataFileStore item in this.additionalStores)
        {
            item.Dispose();
        }

        this.application?.Quit();
        this.application = null;
        this.session = null;
    }
}
