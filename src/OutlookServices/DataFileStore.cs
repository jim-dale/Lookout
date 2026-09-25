namespace OutlookServices;

using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Outlook;

/// <summary>
/// Manages the temporary addition and removal of an Outlook data file store.
/// </summary>
/// <remarks>
/// The store is added when <see cref="TryAddStore"/> is called and removed when
/// <see cref="Dispose"/> or <see cref="EnsureUnloaded"/> is called.
/// </remarks>
public sealed class DataFileStore : IDisposable
{
    private readonly string path;
    private readonly NameSpace session;

    /// <summary>
    /// Initializes a new instance of the <see cref="DataFileStore"/> class.
    /// </summary>
    /// <param name="session">The Outlook namespace in which to add the store.</param>
    /// <param name="path">The path to the Outlook data file.</param>
    /// <exception cref="ArgumentNullException">
    /// <paramref name="session"/> is <see langword="null"/>.
    /// </exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="path"/> is empty or consists only of white-space characters.
    /// </exception>
    public DataFileStore(NameSpace session, string path)
    {
        ArgumentNullException.ThrowIfNull(session);
        ArgumentException.ThrowIfNullOrWhiteSpace(path);

        this.session = session;
        this.path = path;
    }

    /// <summary>
    /// Gets the Outlook store that was added, or <see langword="null"/> if no store is loaded.
    /// </summary>
    public Store? Store { get; private set; }

    /// <summary>
    /// Removes the added Outlook data file store, if present.
    /// </summary>
    public void Dispose()
    {
        this.EnsureUnloaded();
    }

    /// <summary>
    /// Attempts to add the Outlook data file store.
    /// </summary>
    /// <returns>
    /// <see langword="true"/> if the store was added; otherwise, <see langword="false"/>.
    /// </returns>
    public bool TryAddStore()
    {
        bool result = false;

        if (this.Store is null)
        {
            try
            {
                this.session.AddStoreEx(this.path, OlStoreType.olStoreUnicode);

                foreach (Store item in this.session.Stores.OfType<Store>())
                {
                    if (item.FilePath == this.path)
                    {
                        this.Store = item;
                        result = true;
                        break;
                    }
                }
            }
            catch (COMException)
            {
                // ignore COM exceptions
            }
        }

        return result;
    }

    /// <summary>
    /// Removes the Outlook data file store, if it was added, and clears the stored reference.
    /// </summary>
    public void EnsureUnloaded()
    {
        if (this.Store is not null)
        {
            Folder? folder = this.Store.GetRootFolder() as Folder;
            if (folder is not null)
            {
                this.session.RemoveStore(folder);
            }

            this.Store = null;
        }
    }
}
